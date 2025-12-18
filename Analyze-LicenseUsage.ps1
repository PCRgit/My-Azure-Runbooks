#Requires -Modules @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="2.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Analyzes Microsoft 365 license usage and identifies optimization opportunities.

.DESCRIPTION
    This runbook analyzes license assignments, identifies unused or underutilized licenses,
    and provides cost optimization recommendations. Generates an Excel report with
    conditional formatting and estimated cost savings.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Graph API Permissions Required:
        - User.Read.All
        - Organization.Read.All
        - Mail.Send (if using email notifications)
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("admin@yourdomain.com", "finance@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "License Usage Analysis Report - $(Get-Date -Format 'yyyy-MM-dd')"

# License Cost Configuration (Update with your actual costs per month)
$LicenseCosts = @{
    'Microsoft 365 E3' = 36.00
    'Microsoft 365 E5' = 57.00
    'Microsoft 365 Business Premium' = 22.00
    'Microsoft 365 Business Standard' = 12.50
    'Office 365 E3' = 23.00
    'Office 365 E5' = 38.00
    'Enterprise Mobility + Security E3' = 10.60
    'Enterprise Mobility + Security E5' = 16.40
    'Power BI Pro' = 9.99
    'Project Plan 3' = 30.00
    'Visio Plan 2' = 15.00
}

# Inactivity Threshold for identifying unused licenses
$InactivityThreshold = 90  # days

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "LicenseUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$FullReportPath = Join-Path -Path $ExportPath -ChildPath $ReportFileName
#endregion

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Output "[$timestamp] [$Level] $Message"
}

function Connect-MgGraphWithManagedIdentity {
    try {
        Write-Log "Connecting to Microsoft Graph using Managed Identity..."
        Connect-MgGraph -Identity -NoWelcome
        Write-Log "Successfully connected to Microsoft Graph"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-SubscribedSKUs {
    Write-Log "Retrieving subscribed SKUs..."
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        $skuDetails = @()
        
        foreach ($sku in $response.value) {
            $skuDetails += [PSCustomObject]@{
                'SkuId' = $sku.skuId
                'SkuPartNumber' = $sku.skuPartNumber
                'ProductName' = $sku.skuPartNumber -replace '_', ' '
                'TotalLicenses' = $sku.prepaidUnits.enabled
                'AssignedLicenses' = $sku.consumedUnits
                'AvailableLicenses' = $sku.prepaidUnits.enabled - $sku.consumedUnits
                'CapabilityStatus' = $sku.capabilityStatus
            }
        }
        
        Write-Log "Retrieved $($skuDetails.Count) SKUs"
        return $skuDetails
    }
    catch {
        Write-Log "Error retrieving SKUs: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-UserLicenseAssignments {
    Write-Log "Retrieving user license assignments..."
    
    try {
        $allUsers = @()
        $licenseAssignments = @()
        
        # Get all users with license details
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,assignedLicenses,signInActivity,createdDateTime,userType"
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allUsers += $response.value
            $uri = $response.'@odata.nextLink'
            
            if ($uri) {
                Start-Sleep -Milliseconds 100  # Rate limiting
            }
        } while ($uri)
        
        Write-Log "Retrieved $($allUsers.Count) users"
        
        # Get SKU details for name mapping
        $skus = Get-SubscribedSKUs
        $skuLookup = @{}
        foreach ($sku in $skus) {
            $skuLookup[$sku.SkuId] = $sku.ProductName
        }
        
        # Process each user
        foreach ($user in $allUsers) {
            if ($user.assignedLicenses.Count -gt 0) {
                $lastSignIn = $null
                $daysInactive = $null
                
                if ($null -ne $user.signInActivity -and $null -ne $user.signInActivity.lastSignInDateTime) {
                    $lastSignIn = [DateTime]$user.signInActivity.lastSignInDateTime
                    $daysInactive = [int]((Get-Date) - $lastSignIn).TotalDays
                }
                
                # Process each license
                foreach ($license in $user.assignedLicenses) {
                    $skuId = $license.skuId
                    $productName = if ($skuLookup.ContainsKey($skuId)) { $skuLookup[$skuId] } else { "Unknown SKU: $skuId" }
                    
                    # Determine cost
                    $monthlyCost = $null
                    foreach ($key in $LicenseCosts.Keys) {
                        if ($productName -like "*$key*") {
                            $monthlyCost = $LicenseCosts[$key]
                            break
                        }
                    }
                    
                    # Determine if license is potentially wasted
                    $isWasted = $false
                    $wasteReason = ''
                    
                    if (-not $user.accountEnabled) {
                        $isWasted = $true
                        $wasteReason = 'Account Disabled'
                    }
                    elseif ($null -ne $daysInactive -and $daysInactive -gt $InactivityThreshold) {
                        $isWasted = $true
                        $wasteReason = "Inactive for $daysInactive days"
                    }
                    
                    $licenseAssignments += [PSCustomObject]@{
                        'DisplayName' = $user.displayName
                        'UserPrincipalName' = $user.userPrincipalName
                        'LicenseName' = $productName
                        'AccountEnabled' = $user.accountEnabled
                        'LastSignIn' = if ($null -ne $lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
                        'DaysInactive' = if ($null -ne $daysInactive) { $daysInactive } else { 'N/A' }
                        'IsWasted' = $isWasted
                        'WasteReason' = $wasteReason
                        'MonthlyCost' = $monthlyCost
                        'AnnualCost' = if ($null -ne $monthlyCost) { $monthlyCost * 12 } else { $null }
                        'UserType' = $user.userType
                        'CreatedDate' = (Get-Date $user.createdDateTime).ToString('yyyy-MM-dd')
                        'UserId' = $user.id
                    }
                }
            }
        }
        
        Write-Log "Processed $($licenseAssignments.Count) license assignments"
        Write-Log "  - Potentially wasted licenses: $(($licenseAssignments | Where-Object {$_.IsWasted}).Count)"
        
        return $licenseAssignments
    }
    catch {
        Write-Log "Error retrieving user licenses: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Calculate-CostSavings {
    param([object[]]$WastedLicenses)
    
    Write-Log "Calculating potential cost savings..."
    
    $totalMonthlySavings = 0
    $totalAnnualSavings = 0
    $savingsByLicense = @{}
    
    foreach ($license in $WastedLicenses) {
        if ($null -ne $license.MonthlyCost) {
            $totalMonthlySavings += $license.MonthlyCost
            $totalAnnualSavings += $license.AnnualCost
            
            $licenseName = $license.LicenseName
            if (-not $savingsByLicense.ContainsKey($licenseName)) {
                $savingsByLicense[$licenseName] = @{
                    Count = 0
                    MonthlyCost = 0
                    AnnualCost = 0
                }
            }
            
            $savingsByLicense[$licenseName].Count++
            $savingsByLicense[$licenseName].MonthlyCost += $license.MonthlyCost
            $savingsByLicense[$licenseName].AnnualCost += $license.AnnualCost
        }
    }
    
    Write-Log "Potential savings - Monthly: `$$totalMonthlySavings, Annual: `$$totalAnnualSavings"
    
    return @{
        TotalMonthlySavings = $totalMonthlySavings
        TotalAnnualSavings = $totalAnnualSavings
        SavingsByLicense = $savingsByLicense
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$SKUs,
        [object[]]$LicenseAssignments,
        [hashtable]$CostSavings,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Create executive summary
        $wastedLicenses = $LicenseAssignments | Where-Object {$_.IsWasted}
        $totalAssigned = ($SKUs | Measure-Object -Property AssignedLicenses -Sum).Sum
        $totalAvailable = ($SKUs | Measure-Object -Property AvailableLicenses -Sum).Sum
        
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total License SKUs',
                'Total Licenses Assigned',
                'Total Licenses Available',
                'Potentially Wasted Licenses',
                'Waste Percentage',
                '--- Cost Optimization ---',
                'Potential Monthly Savings',
                'Potential Annual Savings',
                'Report Generated'
            )
            'Value' = @(
                $SKUs.Count,
                $totalAssigned,
                $totalAvailable,
                $wastedLicenses.Count,
                "$(if ($totalAssigned -gt 0) { [math]::Round(($wastedLicenses.Count / $totalAssigned) * 100, 2) } else { 0 })%",
                '',
                "`$$([math]::Round($CostSavings.TotalMonthlySavings, 2))",
                "`$$([math]::Round($CostSavings.TotalAnnualSavings, 2))",
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Export SKU details
        $SKUs | Sort-Object ProductName | Export-Excel -Path $FilePath -WorksheetName "License SKUs" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
            -ConditionalText $(
                New-ConditionalText -Text "Enabled" -Range "G:G" -BackgroundColor LightGreen
                New-ConditionalText -Text "Suspended" -Range "G:G" -BackgroundColor Orange
            )
        
        # Savings by license type
        if ($CostSavings.SavingsByLicense.Count -gt 0) {
            $savingsBreakdown = @()
            foreach ($key in $CostSavings.SavingsByLicense.Keys) {
                $savingsBreakdown += [PSCustomObject]@{
                    'LicenseType' = $key
                    'WastedCount' = $CostSavings.SavingsByLicense[$key].Count
                    'MonthlySavings' = "`$$([math]::Round($CostSavings.SavingsByLicense[$key].MonthlyCost, 2))"
                    'AnnualSavings' = "`$$([math]::Round($CostSavings.SavingsByLicense[$key].AnnualCost, 2))"
                }
            }
            
            $savingsBreakdown | Sort-Object {[decimal]($_.AnnualSavings -replace '[\$,]', '')} -Descending | 
                Export-Excel -Path $FilePath -WorksheetName "Savings by License" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Top 20 wasted licenses (highest cost first)
        $top20Wasted = $wastedLicenses | Where-Object {$null -ne $_.MonthlyCost} | 
            Sort-Object AnnualCost -Descending | Select-Object -First 20
        
        if ($top20Wasted.Count -gt 0) {
            $top20Wasted | Export-Excel -Path $FilePath -WorksheetName "Top 20 Wasted" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # All wasted licenses
        if ($wastedLicenses.Count -gt 0) {
            $wastedLicenses | Sort-Object AnnualCost -Descending | 
                Export-Excel -Path $FilePath -WorksheetName "All Wasted Licenses" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "True" -Range "G:G" -BackgroundColor LightCoral
                    New-ConditionalText -Text "False" -Range "D:D" -BackgroundColor LightPink
                )
        }
        
        # All license assignments
        $LicenseAssignments | Export-Excel -Path $FilePath -WorksheetName "All Assignments" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
            -ConditionalText $(
                New-ConditionalText -Text "True" -Range "G:G" -BackgroundColor LightCoral
                New-ConditionalText -Text "False" -Range "G:G" -BackgroundColor LightGreen
            )
        
        Write-Log "Excel report created successfully: $FilePath"
        return $true
    }
    catch {
        Write-Log "Error creating Excel report: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Send-EmailWithAttachment {
    param(
        [string[]]$Recipients,
        [string]$Subject,
        [string]$AttachmentPath,
        [int]$WastedCount,
        [decimal]$MonthlySavings,
        [decimal]$AnnualSavings
    )
    
    Write-Log "Preparing to send email notification..."
    
    # Create HTML body
    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        .summary { background-color: #f0f0f0; padding: 15px; border-radius: 5px; margin: 20px 0; }
        .metric { font-size: 16px; margin: 10px 0; }
        .savings { color: #5cb85c; font-weight: bold; font-size: 28px; }
        .wasted { color: #d9534f; font-weight: bold; font-size: 24px; }
        .highlight { background-color: #d4edda; border-left: 4px solid #28a745; padding: 10px; margin: 20px 0; }
        .footer { margin-top: 30px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <h2>License Usage Analysis Report</h2>
    <p>This automated report identifies license optimization opportunities and potential cost savings.</p>
    
    <div class="highlight">
        <div class="metric">
            <strong>💰 Potential Annual Savings:</strong> <span class="savings">`$$([math]::Round($AnnualSavings, 2))</span>
        </div>
        <div class="metric">
            <strong>Monthly Savings:</strong> `$$([math]::Round($MonthlySavings, 2))
        </div>
    </div>
    
    <div class="summary">
        <div class="metric">
            <strong>Potentially Wasted Licenses:</strong> <span class="wasted">$WastedCount</span>
        </div>
    </div>
    
    <p><strong>Report Details:</strong></p>
    <ul>
        <li>Inactivity Threshold: $InactivityThreshold days</li>
        <li>Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</li>
        <li>Detailed analysis is attached in Excel format</li>
    </ul>
    
    <p><strong>Optimization Recommendations:</strong></p>
    <ul>
        <li>Review and reclaim licenses from disabled accounts</li>
        <li>Identify and contact inactive users</li>
        <li>Consider downgrading underutilized premium licenses</li>
        <li>Implement license reclamation policies</li>
        <li>Set up automated license management workflows</li>
        <li>Review license allocation regularly (monthly/quarterly)</li>
    </ul>
    
    <p><strong>Next Steps:</strong></p>
    <ol>
        <li>Review the "Top 20 Wasted" sheet for quick wins</li>
        <li>Contact users in the "All Wasted Licenses" sheet</li>
        <li>Reclaim licenses from disabled or departed users</li>
        <li>Document license reclamation decisions</li>
    </ol>
    
    <div class="footer">
        <p>This is an automated report from Azure Automation. Please do not reply to this email.</p>
    </div>
</body>
</html>
"@
    
    try {
        # Read file as base64
        $attachmentBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentPath))
        
        # Prepare email message
        $emailBody = @{
            message = @{
                subject = $Subject
                body = @{
                    contentType = "HTML"
                    content = $htmlBody
                }
                toRecipients = @(
                    $Recipients | ForEach-Object {
                        @{
                            emailAddress = @{
                                address = $_
                            }
                        }
                    }
                )
                attachments = @(
                    @{
                        "@odata.type" = "#microsoft.graph.fileAttachment"
                        name = (Split-Path $AttachmentPath -Leaf)
                        contentBytes = $attachmentBase64
                    }
                )
            }
            saveToSentItems = "true"
        }
        
        $emailBodyJson = $emailBody | ConvertTo-Json -Depth 10
        
        # Send email
        $uri = "https://graph.microsoft.com/v1.0/users/$EmailFrom/sendMail"
        Invoke-MgGraphRequest -Uri $uri -Method POST -Body $emailBodyJson -ContentType "application/json"
        
        Write-Log "Email sent successfully to: $($Recipients -join ', ')"
        return $true
    }
    catch {
        Write-Log "Error sending email: $($_.Exception.Message)" -Level ERROR
        throw
    }
}
#endregion

#region Main Execution
try {
    Write-Log "========== Starting License Usage Analysis =========="
    
    # Connect to Microsoft Graph
    Connect-MgGraphWithManagedIdentity
    
    # Retrieve data
    $skus = Get-SubscribedSKUs
    $licenseAssignments = Get-UserLicenseAssignments
    
    # Calculate cost savings
    $wastedLicenses = $licenseAssignments | Where-Object {$_.IsWasted}
    $costSavings = Calculate-CostSavings -WastedLicenses $wastedLicenses
    
    # Export to Excel
    Export-ToExcelWithFormatting -SKUs $skus -LicenseAssignments $licenseAssignments `
                                  -CostSavings $costSavings -FilePath $FullReportPath
    
    # Send email notification
    Send-EmailWithAttachment -Recipients $EmailRecipients -Subject $EmailSubject -AttachmentPath $FullReportPath `
                             -WastedCount $wastedLicenses.Count `
                             -MonthlySavings $costSavings.TotalMonthlySavings `
                             -AnnualSavings $costSavings.TotalAnnualSavings
    
    # Cleanup
    if (Test-Path $FullReportPath) {
        Remove-Item $FullReportPath -Force
        Write-Log "Temporary report file cleaned up"
    }
    
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Found $($wastedLicenses.Count) wasted licenses. Potential annual savings: `$$([math]::Round($costSavings.TotalAnnualSavings, 2))"
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    throw
}
finally {
    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph | Out-Null
        Write-Log "Disconnected from Microsoft Graph"
    }
    catch {
        Write-Log "Error disconnecting from Microsoft Graph: $($_.Exception.Message)" -Level WARNING
    }
}
#endregion
