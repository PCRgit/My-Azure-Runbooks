#Requires -Modules @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Applications"; ModuleVersion="2.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Monitors Azure AD App Registration secrets and certificates for expiration.

.DESCRIPTION
    This runbook identifies app registrations with secrets or certificates that are expired
    or expiring soon. Generates an Excel report with conditional formatting and sends
    email notifications to prevent service disruptions.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Graph API Permissions Required:
        - Application.Read.All
        - Mail.Send (if using email notifications)
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("admin@yourdomain.com", "devops@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "App Registration Secret Expiration Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Expiration Thresholds (in days)
$CriticalThreshold = 7    # Expires in 7 days or less
$WarningThreshold = 30    # Expires in 30 days or less
$InfoThreshold = 90       # Expires in 90 days or less

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "AppRegistrationSecrets_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-AppRegistrationSecrets {
    Write-Log "Retrieving app registrations from Microsoft Graph..."
    $allApps = @()
    $secretDetails = @()
    
    try {
        # Get all app registrations
        $uri = "https://graph.microsoft.com/v1.0/applications?`$select=id,appId,displayName,passwordCredentials,keyCredentials"
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allApps += $response.value
            $uri = $response.'@odata.nextLink'
            
            if ($uri) {
                Start-Sleep -Milliseconds 100  # Rate limiting
            }
        } while ($uri)
        
        Write-Log "Retrieved $($allApps.Count) total app registrations"
        
        # Process each app registration
        foreach ($app in $allApps) {
            $appName = $app.displayName
            $appId = $app.appId
            $objectId = $app.id
            
            # Process password credentials (secrets)
            if ($app.passwordCredentials -and $app.passwordCredentials.Count -gt 0) {
                foreach ($secret in $app.passwordCredentials) {
                    $expiryDate = [DateTime]$secret.endDateTime
                    $daysUntilExpiry = [int](($expiryDate - (Get-Date)).TotalDays)
                    
                    $status = if ($daysUntilExpiry -lt 0) {
                        "EXPIRED"
                    } elseif ($daysUntilExpiry -le $CriticalThreshold) {
                        "CRITICAL"
                    } elseif ($daysUntilExpiry -le $WarningThreshold) {
                        "WARNING"
                    } elseif ($daysUntilExpiry -le $InfoThreshold) {
                        "INFO"
                    } else {
                        "OK"
                    }
                    
                    $secretDetails += [PSCustomObject]@{
                        'ApplicationName' = $appName
                        'ApplicationId' = $appId
                        'ObjectId' = $objectId
                        'CredentialType' = 'Secret'
                        'KeyId' = $secret.keyId
                        'DisplayName' = if ($secret.displayName) { $secret.displayName } else { "N/A" }
                        'StartDate' = (Get-Date $secret.startDateTime).ToString('yyyy-MM-dd')
                        'ExpiryDate' = $expiryDate.ToString('yyyy-MM-dd')
                        'DaysUntilExpiry' = $daysUntilExpiry
                        'Status' = $status
                    }
                }
            }
            
            # Process key credentials (certificates)
            if ($app.keyCredentials -and $app.keyCredentials.Count -gt 0) {
                foreach ($cert in $app.keyCredentials) {
                    $expiryDate = [DateTime]$cert.endDateTime
                    $daysUntilExpiry = [int](($expiryDate - (Get-Date)).TotalDays)
                    
                    $status = if ($daysUntilExpiry -lt 0) {
                        "EXPIRED"
                    } elseif ($daysUntilExpiry -le $CriticalThreshold) {
                        "CRITICAL"
                    } elseif ($daysUntilExpiry -le $WarningThreshold) {
                        "WARNING"
                    } elseif ($daysUntilExpiry -le $InfoThreshold) {
                        "INFO"
                    } else {
                        "OK"
                    }
                    
                    $secretDetails += [PSCustomObject]@{
                        'ApplicationName' = $appName
                        'ApplicationId' = $appId
                        'ObjectId' = $objectId
                        'CredentialType' = 'Certificate'
                        'KeyId' = $cert.keyId
                        'DisplayName' = if ($cert.displayName) { $cert.displayName } else { "N/A" }
                        'StartDate' = (Get-Date $cert.startDateTime).ToString('yyyy-MM-dd')
                        'ExpiryDate' = $expiryDate.ToString('yyyy-MM-dd')
                        'DaysUntilExpiry' = $daysUntilExpiry
                        'Status' = $status
                    }
                }
            }
        }
        
        # Sort by days until expiry (critical first)
        $secretDetails = $secretDetails | Sort-Object DaysUntilExpiry
        
        Write-Log "Found $($secretDetails.Count) total credentials across all app registrations"
        Write-Log "  - Expired: $(($secretDetails | Where-Object {$_.Status -eq 'EXPIRED'}).Count)"
        Write-Log "  - Critical (<= $CriticalThreshold days): $(($secretDetails | Where-Object {$_.Status -eq 'CRITICAL'}).Count)"
        Write-Log "  - Warning (<= $WarningThreshold days): $(($secretDetails | Where-Object {$_.Status -eq 'WARNING'}).Count)"
        Write-Log "  - Info (<= $InfoThreshold days): $(($secretDetails | Where-Object {$_.Status -eq 'INFO'}).Count)"
        
        return $secretDetails
    }
    catch {
        Write-Log "Error retrieving app registrations: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$Secrets,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Create summary
        $expired = ($Secrets | Where-Object {$_.Status -eq 'EXPIRED'}).Count
        $critical = ($Secrets | Where-Object {$_.Status -eq 'CRITICAL'}).Count
        $warning = ($Secrets | Where-Object {$_.Status -eq 'WARNING'}).Count
        $info = ($Secrets | Where-Object {$_.Status -eq 'INFO'}).Count
        
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total Credentials',
                'Expired',
                'Critical (≤ 7 days)',
                'Warning (≤ 30 days)',
                'Info (≤ 90 days)',
                'Report Generated'
            )
            'Value' = @(
                $Secrets.Count,
                $expired,
                $critical,
                $warning,
                $info,
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Get top 10 most urgent credentials
        $top10 = $Secrets | Where-Object {$_.Status -in @('EXPIRED', 'CRITICAL', 'WARNING')} | Select-Object -First 10
        
        if ($top10.Count -gt 0) {
            $top10 | Export-Excel -Path $FilePath -WorksheetName "Top 10 Urgent" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Export all credentials with conditional formatting
        if ($Secrets.Count -gt 0) {
            $Secrets | Export-Excel -Path $FilePath -WorksheetName "All Credentials" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "EXPIRED" -Range "J:J" -BackgroundColor Red -ConditionalTextColor White
                    New-ConditionalText -Text "CRITICAL" -Range "J:J" -BackgroundColor Orange
                    New-ConditionalText -Text "WARNING" -Range "J:J" -BackgroundColor Yellow
                    New-ConditionalText -Text "INFO" -Range "J:J" -BackgroundColor LightBlue
                    New-ConditionalText -Text "OK" -Range "J:J" -BackgroundColor LightGreen
                )
        }
        
        # Export by status
        $expiredSecrets = $Secrets | Where-Object {$_.Status -eq 'EXPIRED'}
        if ($expiredSecrets.Count -gt 0) {
            $expiredSecrets | Export-Excel -Path $FilePath -WorksheetName "Expired" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        $criticalSecrets = $Secrets | Where-Object {$_.Status -eq 'CRITICAL'}
        if ($criticalSecrets.Count -gt 0) {
            $criticalSecrets | Export-Excel -Path $FilePath -WorksheetName "Critical" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
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
        [int]$ExpiredCount,
        [int]$CriticalCount,
        [int]$WarningCount
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
        .expired { color: #d9534f; font-weight: bold; font-size: 24px; }
        .critical { color: #f0ad4e; font-weight: bold; font-size: 24px; }
        .warning { color: #f0ad4e; font-weight: bold; font-size: 20px; }
        .alert { background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; margin: 20px 0; }
        .footer { margin-top: 30px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <h2>App Registration Secret Expiration Report</h2>
    <p>This automated report identifies app registration secrets and certificates that require attention.</p>
    
    <div class="summary">
        <div class="metric">
            <strong>Expired Credentials:</strong> <span class="expired">$ExpiredCount</span>
        </div>
        <div class="metric">
            <strong>Critical (≤ 7 days):</strong> <span class="critical">$CriticalCount</span>
        </div>
        <div class="metric">
            <strong>Warning (≤ 30 days):</strong> <span class="warning">$WarningCount</span>
        </div>
    </div>
    
    $(if ($ExpiredCount -gt 0 -or $CriticalCount -gt 0) {
        @"
    <div class="alert">
        <strong>⚠️ IMMEDIATE ACTION REQUIRED</strong><br/>
        There are expired or critically expiring credentials that could cause service disruptions.
    </div>
"@
    })
    
    <p><strong>Report Details:</strong></p>
    <ul>
        <li>Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</li>
        <li>Detailed results are attached in Excel format</li>
    </ul>
    
    <p><strong>Recommended Actions:</strong></p>
    <ul>
        <li>Rotate expired credentials immediately</li>
        <li>Plan rotation for critical and warning credentials</li>
        <li>Update application configurations with new secrets</li>
        <li>Document credential rotation procedures</li>
        <li>Set up proactive monitoring and alerts</li>
    </ul>
    
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
    Write-Log "========== Starting App Registration Secret Expiration Monitoring =========="
    
    # Connect to Microsoft Graph
    Connect-MgGraphWithManagedIdentity
    
    # Retrieve app registration secrets
    $secrets = Get-AppRegistrationSecrets
    
    # Calculate counts for email
    $expiredCount = ($secrets | Where-Object {$_.Status -eq 'EXPIRED'}).Count
    $criticalCount = ($secrets | Where-Object {$_.Status -eq 'CRITICAL'}).Count
    $warningCount = ($secrets | Where-Object {$_.Status -eq 'WARNING'}).Count
    
    # Export to Excel
    Export-ToExcelWithFormatting -Secrets $secrets -FilePath $FullReportPath
    
    # Send email notification
    Send-EmailWithAttachment -Recipients $EmailRecipients -Subject $EmailSubject -AttachmentPath $FullReportPath `
                             -ExpiredCount $expiredCount -CriticalCount $criticalCount -WarningCount $warningCount
    
    # Cleanup
    if (Test-Path $FullReportPath) {
        Remove-Item $FullReportPath -Force
        Write-Log "Temporary report file cleaned up"
    }
    
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Found $($secrets.Count) total credentials. Expired: $expiredCount, Critical: $criticalCount, Warning: $warningCount"
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
