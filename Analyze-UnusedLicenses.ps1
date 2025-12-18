<#
.SYNOPSIS
    Analyzes Microsoft 365 license assignments and identifies unused licenses
.DESCRIPTION
    Identifies users with assigned licenses who haven't signed in for 90+ days
    Calculates potential cost savings from license reclamation
    Provides detailed Excel report with license usage analysis
.NOTES
    Author: Jaimin
    Requires: Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement
    Authentication: Managed Identity
    Graph API Permissions: User.Read.All, Organization.Read.All, Mail.Send
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipients = "licensing-team@something.com",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:TEMP\UnusedLicenses_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# License cost estimation (monthly USD) - Update with your actual costs
$LicenseCosts = @{
    'ENTERPRISEPACK'              = 20.00  # Office 365 E3
    'ENTERPRISEPREMIUM'           = 36.00  # Office 365 E5
    'SPE_E3'                      = 32.00  # Microsoft 365 E3
    'SPE_E5'                      = 57.00  # Microsoft 365 E5
    'POWER_BI_PRO'                = 9.99   # Power BI Pro
    'PROJECTPREMIUM'              = 55.00  # Project Plan 5
    'VISIOONLINE_PLAN2'           = 15.00  # Visio Plan 2
    'FLOW_FREE'                   = 0.00   # Power Automate Free
    'POWER_BI_STANDARD'           = 0.00   # Power BI Free
    'ENTERPRISEPACK_FACULTY'      = 0.00   # Office 365 A3 (Education)
    'M365_F1'                     = 10.00  # Microsoft 365 F1
    'EMS'                         = 8.70   # Enterprise Mobility + Security E3
    'EMSPREMIUM'                  = 14.80  # Enterprise Mobility + Security E5
}

# Connect to Microsoft Graph
try {
    Write-Output "Connecting to Microsoft Graph with Managed Identity..."
    Connect-MgGraph -Identity -NoWelcome
    Write-Output "Successfully connected to Microsoft Graph"
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    throw
}

# Calculate threshold date
$ThresholdDate = (Get-Date).AddDays(-$InactiveDays)
Write-Output "Analyzing licenses for users inactive since: $($ThresholdDate.ToString('yyyy-MM-dd'))"

# Get all subscribed SKUs
Write-Output "Retrieving license information..."
$SubscribedSkus = Get-MgSubscribedSku

# Get all users with licenses
Write-Output "Retrieving licensed users..."
$LicensedUsers = Get-MgUser -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable UserCount -All `
    -Property Id, DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, SignInActivity, CreatedDateTime, UserType

Write-Output "Analyzing $($LicensedUsers.Count) licensed users..."

$UnusedLicenses = [System.Collections.Generic.List[PSCustomObject]]::new()
$ActiveLicenses = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllLicenseDetails = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($User in $LicensedUsers) {
    # Get last sign-in date
    $LastSignIn = $null
    $DaysSinceSignIn = $null
    
    if ($User.SignInActivity.LastSignInDateTime) {
        $LastSignIn = [datetime]$User.SignInActivity.LastSignInDateTime
        $DaysSinceSignIn = (New-TimeSpan -Start $LastSignIn -End (Get-Date)).Days
    }
    elseif ($User.CreatedDateTime) {
        # User never signed in, use creation date as reference
        $LastSignIn = [datetime]$User.CreatedDateTime
        $DaysSinceSignIn = (New-TimeSpan -Start $LastSignIn -End (Get-Date)).Days
    }
    
    $IsInactive = ($DaysSinceSignIn -ge $InactiveDays) -or ($null -eq $LastSignIn)
    
    foreach ($License in $User.AssignedLicenses) {
        $Sku = $SubscribedSkus | Where-Object { $_.SkuId -eq $License.SkuId }
        
        if ($Sku) {
            $SkuPartNumber = $Sku.SkuPartNumber
            $LicenseCost = if ($LicenseCosts.ContainsKey($SkuPartNumber)) { 
                $LicenseCosts[$SkuPartNumber] 
            }
            else { 
                0.00 
            }
            
            $LicenseDetail = [PSCustomObject]@{
                UserDisplayName    = $User.DisplayName
                UserPrincipalName  = $User.UserPrincipalName
                AccountEnabled     = $User.AccountEnabled
                UserType           = $User.UserType
                LicenseName        = $Sku.SkuPartNumber
                LicenseFriendlyName = if ($Sku.DisplayName) { $Sku.DisplayName } else { $Sku.SkuPartNumber }
                LastSignInDate     = if ($LastSignIn) { $LastSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Never' }
                DaysSinceSignIn    = $DaysSinceSignIn
                IsInactive         = $IsInactive
                Status             = if ($IsInactive) { 'Unused' } elseif (-not $User.AccountEnabled) { 'Disabled' } else { 'Active' }
                MonthlyCost        = $LicenseCost
                AnnualCost         = ($LicenseCost * 12)
            }
            
            $AllLicenseDetails.Add($LicenseDetail)
            
            if ($IsInactive -or -not $User.AccountEnabled) {
                $UnusedLicenses.Add($LicenseDetail)
            }
            else {
                $ActiveLicenses.Add($LicenseDetail)
            }
        }
    }
    
    # Rate limiting
    if ($LicensedUsers.IndexOf($User) % 100 -eq 0) {
        Start-Sleep -Milliseconds 100
    }
}

# Calculate savings
$TotalPotentialMonthlySavings = ($UnusedLicenses | Measure-Object -Property MonthlyCost -Sum).Sum
$TotalPotentialAnnualSavings = ($UnusedLicenses | Measure-Object -Property AnnualCost -Sum).Sum

# License summary by SKU
$LicenseSummary = $AllLicenseDetails | Group-Object LicenseName | ForEach-Object {
    $Group = $_.Group
    $UnusedCount = ($Group | Where-Object { $_.IsInactive }).Count
    $ActiveCount = ($Group | Where-Object { -not $_.IsInactive }).Count
    $TotalCost = ($Group | Where-Object { $_.IsInactive } | Measure-Object -Property MonthlyCost -Sum).Sum
    
    [PSCustomObject]@{
        LicenseType          = $_.Name
        TotalAssigned        = $Group.Count
        ActiveCount          = $ActiveCount
        UnusedCount          = $UnusedCount
        UnusedPercentage     = [math]::Round(($UnusedCount / $Group.Count) * 100, 2)
        MonthlyWaste         = $TotalCost
        AnnualWaste          = ($TotalCost * 12)
    }
} | Sort-Object AnnualWaste -Descending

# Top 20 users with most expensive unused licenses
$TopWastefulUsers = $UnusedLicenses | 
    Group-Object UserPrincipalName | 
    ForEach-Object {
        [PSCustomObject]@{
            UserPrincipalName = $_.Name
            UserDisplayName   = $_.Group[0].UserDisplayName
            LicenseCount      = $_.Count
            Licenses          = ($_.Group.LicenseName -join ', ')
            MonthlyCost       = ($_.Group | Measure-Object -Property MonthlyCost -Sum).Sum
            AnnualCost        = ($_.Group | Measure-Object -Property AnnualCost -Sum).Sum
            LastSignInDate    = $_.Group[0].LastSignInDate
            DaysSinceSignIn   = $_.Group[0].DaysSinceSignIn
        }
    } | Sort-Object AnnualCost -Descending | Select-Object -First 20

# Generate Excel report
Write-Output "Generating Excel report..."

try {
    # Summary Dashboard
    $Summary = [PSCustomObject]@{
        ReportDate                 = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        InactiveThresholdDays      = $InactiveDays
        TotalLicensedUsers         = $LicensedUsers.Count
        TotalLicensesAssigned      = $AllLicenseDetails.Count
        UnusedLicenseCount         = $UnusedLicenses.Count
        ActiveLicenseCount         = $ActiveLicenses.Count
        PotentialMonthlySavings    = [math]::Round($TotalPotentialMonthlySavings, 2)
        PotentialAnnualSavings     = [math]::Round($TotalPotentialAnnualSavings, 2)
    }
    
    $Summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -BoldTopRow -FreezeTopRow
    
    if ($LicenseSummary.Count -gt 0) {
        $LicenseSummary | Export-Excel -Path $OutputPath -WorksheetName "LicenseSummary" -AutoSize -BoldTopRow -FreezeTopRow -Append `
            -ConditionalText $(
                New-ConditionalText -Range "E:E" -ConditionalType GreaterThan 50 -BackgroundColor Red -ConditionalTextColor White
                New-ConditionalText -Range "E:E" -ConditionalType Between 25 50 -BackgroundColor Yellow
            )
    }
    
    if ($UnusedLicenses.Count -gt 0) {
        $UnusedLicenses | Sort-Object AnnualCost -Descending | Export-Excel -Path $OutputPath -WorksheetName "UnusedLicenses" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($TopWastefulUsers.Count -gt 0) {
        $TopWastefulUsers | Export-Excel -Path $OutputPath -WorksheetName "Top20WastefulUsers" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($ActiveLicenses.Count -gt 0) {
        $ActiveLicenses | Sort-Object UserPrincipalName | Export-Excel -Path $OutputPath -WorksheetName "ActiveLicenses" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    Write-Output "Excel report generated: $OutputPath"
}
catch {
    Write-Error "Failed to generate Excel report: $_"
    throw
}

# Send email
try {
    $GraphToken = (Get-MgContext).TokenCredential.GetTokenAsync(
        (New-Object System.Threading.CancellationToken), 
        (New-Object Microsoft.Identity.Client.AuthenticationProviderOption)
    ).Result.Token
    
    if ($GraphToken -and (Test-Path $OutputPath)) {
        Write-Output "Sending email notification..."
        
        $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
        $FileBase64 = [System.Convert]::ToBase64String($FileBytes)
        $FileName = Split-Path $OutputPath -Leaf
        
        # Top licenses by waste
        $TopLicenses = $LicenseSummary | Select-Object -First 5 | ForEach-Object {
            "<li><strong>$($_.LicenseType)</strong>: $($_.UnusedCount) unused ($([math]::Round($_.AnnualWaste, 2)) annual waste)</li>"
        }
        
        $EmailBody = @"
<html>
<body>
<h2>Microsoft 365 License Usage Analysis</h2>
<p>Report generated on: <strong>$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</strong></p>

<h3>💰 Cost Savings Opportunity:</h3>
<ul>
    <li><strong style="color:red; font-size:16px;">Potential Monthly Savings: `$$($Summary.PotentialMonthlySavings)</strong></li>
    <li><strong style="color:red; font-size:16px;">Potential Annual Savings: `$$($Summary.PotentialAnnualSavings)</strong></li>
</ul>

<h3>Summary:</h3>
<ul>
    <li>Total Licensed Users: <strong>$($Summary.TotalLicensedUsers)</strong></li>
    <li>Total Licenses Assigned: <strong>$($Summary.TotalLicensesAssigned)</strong></li>
    <li>Unused Licenses: <strong style="color:red;">$($Summary.UnusedLicenseCount)</strong> (inactive for $InactiveDays+ days)</li>
    <li>Active Licenses: <strong style="color:green;">$($Summary.ActiveLicenseCount)</strong></li>
</ul>

<h3>Top License Types by Waste:</h3>
<ul>
$($TopLicenses -join "`n")
</ul>

<p><strong>Action Required:</strong> Review the attached Excel workbook to identify specific users whose licenses can be reclaimed.</p>

<p><em>This is an automated report from Azure Automation.</em></p>
</body>
</html>
"@
        
        $Headers = @{
            "Authorization" = "Bearer $GraphToken"
            "Content-Type"  = "application/json"
        }
        
        $EmailMessage = @{
            message = @{
                subject      = "💰 License Optimization Report - Potential Savings: `$$($Summary.PotentialAnnualSavings)/year"
                body         = @{
                    contentType = "HTML"
                    content     = $EmailBody
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $EmailRecipients
                        }
                    }
                )
                attachments  = @(
                    @{
                        "@odata.type"  = "#microsoft.graph.fileAttachment"
                        name           = $FileName
                        contentType    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        contentBytes   = $FileBase64
                    }
                )
            }
        }
        
        $EmailJson = $EmailMessage | ConvertTo-Json -Depth 10
        
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me/sendMail" `
            -Headers $Headers `
            -Method Post `
            -Body $EmailJson
        
        Write-Output "Email sent successfully"
    }
}
catch {
    Write-Warning "Failed to send email: $_"
}

Disconnect-MgGraph | Out-Null
Write-Output "Script completed successfully"