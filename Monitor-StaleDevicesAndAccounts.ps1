#Requires -Modules @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.DeviceManagement"; ModuleVersion="2.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Monitors and reports on stale devices and user accounts in Azure AD/Entra ID.

.DESCRIPTION
    This runbook identifies inactive devices and user accounts based on configurable thresholds.
    Generates an Excel report with conditional formatting and sends email notifications.
    Uses Managed Identity for secure authentication.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Graph API Permissions Required:
        - Device.Read.All
        - User.Read.All
        - Mail.Send (if using email notifications)
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("admin@yourdomain.com", "security@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "Stale Devices and Accounts Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Inactivity Thresholds (in days)
$StaleDeviceThreshold = 90
$StaleUserThreshold = 90

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "StaleDevicesAndAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-StaleDevices {
    param([int]$DaysInactive)
    
    Write-Log "Retrieving devices from Microsoft Graph..."
    $allDevices = @()
    $cutoffDate = (Get-Date).AddDays(-$DaysInactive)
    
    try {
        # Get all devices with required properties
        $uri = "https://graph.microsoft.com/v1.0/devices?`$select=id,displayName,operatingSystem,operatingSystemVersion,approximateLastSignInDateTime,accountEnabled,deviceId,trustType"
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allDevices += $response.value
            $uri = $response.'@odata.nextLink'
            
            if ($uri) {
                Start-Sleep -Milliseconds 100  # Rate limiting
            }
        } while ($uri)
        
        Write-Log "Retrieved $($allDevices.Count) total devices"
        
        # Filter for stale devices
        $staleDevices = $allDevices | Where-Object {
            $lastSignIn = $_.approximateLastSignInDateTime
            if ([string]::IsNullOrEmpty($lastSignIn)) {
                return $true  # No sign-in date = stale
            }
            return ([DateTime]$lastSignIn -lt $cutoffDate)
        } | Select-Object @{N='DeviceName';E={$_.displayName}},
                          @{N='OperatingSystem';E={$_.operatingSystem}},
                          @{N='OSVersion';E={$_.operatingSystemVersion}},
                          @{N='LastSignIn';E={
                              if ([string]::IsNullOrEmpty($_.approximateLastSignInDateTime)) {
                                  'Never'
                              } else {
                                  (Get-Date $_.approximateLastSignInDateTime).ToString('yyyy-MM-dd')
                              }
                          }},
                          @{N='DaysInactive';E={
                              if ([string]::IsNullOrEmpty($_.approximateLastSignInDateTime)) {
                                  'N/A'
                              } else {
                                  [int]((Get-Date) - (Get-Date $_.approximateLastSignInDateTime)).Days
                              }
                          }},
                          @{N='Enabled';E={$_.accountEnabled}},
                          @{N='TrustType';E={$_.trustType}},
                          @{N='DeviceId';E={$_.deviceId}}
        
        Write-Log "Found $($staleDevices.Count) stale devices"
        return $staleDevices
    }
    catch {
        Write-Log "Error retrieving devices: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-StaleUsers {
    param([int]$DaysInactive)
    
    Write-Log "Retrieving users from Microsoft Graph..."
    $allUsers = @()
    $cutoffDate = (Get-Date).AddDays(-$DaysInactive)
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,signInActivity,createdDateTime,userType,mail"
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allUsers += $response.value
            $uri = $response.'@odata.nextLink'
            
            if ($uri) {
                Start-Sleep -Milliseconds 100  # Rate limiting
            }
        } while ($uri)
        
        Write-Log "Retrieved $($allUsers.Count) total users"
        
        # Filter for stale users (excluding guests by default)
        $staleUsers = $allUsers | Where-Object {
            $_.userType -ne 'Guest' -and
            ($null -ne $_.signInActivity -and 
             $null -ne $_.signInActivity.lastSignInDateTime -and
             [DateTime]$_.signInActivity.lastSignInDateTime -lt $cutoffDate) -or
            ($null -eq $_.signInActivity.lastSignInDateTime)
        } | Select-Object @{N='DisplayName';E={$_.displayName}},
                          @{N='UserPrincipalName';E={$_.userPrincipalName}},
                          @{N='Email';E={$_.mail}},
                          @{N='LastSignIn';E={
                              if ($null -eq $_.signInActivity.lastSignInDateTime) {
                                  'Never'
                              } else {
                                  (Get-Date $_.signInActivity.lastSignInDateTime).ToString('yyyy-MM-dd')
                              }
                          }},
                          @{N='DaysInactive';E={
                              if ($null -eq $_.signInActivity.lastSignInDateTime) {
                                  'N/A'
                              } else {
                                  [int]((Get-Date) - (Get-Date $_.signInActivity.lastSignInDateTime)).Days
                              }
                          }},
                          @{N='AccountEnabled';E={$_.accountEnabled}},
                          @{N='CreatedDate';E={(Get-Date $_.createdDateTime).ToString('yyyy-MM-dd')}},
                          @{N='UserId';E={$_.id}}
        
        Write-Log "Found $($staleUsers.Count) stale users"
        return $staleUsers
    }
    catch {
        Write-Log "Error retrieving users: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$StaleDevices,
        [object[]]$StaleUsers,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Create summary sheet
        $summary = [PSCustomObject]@{
            'Metric' = @('Total Stale Devices', 'Enabled Stale Devices', 'Total Stale Users', 'Enabled Stale Users', 'Report Generated', 'Inactivity Threshold (Days)')
            'Value' = @(
                $StaleDevices.Count,
                ($StaleDevices | Where-Object {$_.Enabled -eq $true}).Count,
                $StaleUsers.Count,
                ($StaleUsers | Where-Object {$_.AccountEnabled -eq $true}).Count,
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),
                $StaleDeviceThreshold
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Export Stale Devices with conditional formatting
        if ($StaleDevices.Count -gt 0) {
            $StaleDevices | Export-Excel -Path $FilePath -WorksheetName "Stale Devices" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "True" -Range "F:F" -BackgroundColor LightGreen
                    New-ConditionalText -Text "False" -Range "F:F" -BackgroundColor LightCoral
                )
        }
        
        # Export Stale Users with conditional formatting
        if ($StaleUsers.Count -gt 0) {
            $StaleUsers | Export-Excel -Path $FilePath -WorksheetName "Stale Users" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "True" -Range "F:F" -BackgroundColor LightGreen
                    New-ConditionalText -Text "False" -Range "F:F" -BackgroundColor LightCoral
                )
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
        [int]$StaleDeviceCount,
        [int]$StaleUserCount
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
        .metric { font-size: 18px; margin: 10px 0; }
        .number { color: #d9534f; font-weight: bold; font-size: 24px; }
        .footer { margin-top: 30px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <h2>Stale Devices and Accounts Report</h2>
    <p>This automated report identifies inactive devices and user accounts in your organization.</p>
    
    <div class="summary">
        <div class="metric">
            <strong>Stale Devices Found:</strong> <span class="number">$StaleDeviceCount</span>
        </div>
        <div class="metric">
            <strong>Stale User Accounts Found:</strong> <span class="number">$StaleUserCount</span>
        </div>
    </div>
    
    <p><strong>Report Details:</strong></p>
    <ul>
        <li>Inactivity Threshold: $StaleDeviceThreshold days</li>
        <li>Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</li>
        <li>Detailed results are attached in Excel format</li>
    </ul>
    
    <p><strong>Recommended Actions:</strong></p>
    <ul>
        <li>Review stale devices for potential security risks</li>
        <li>Disable or remove accounts that are no longer needed</li>
        <li>Contact device owners to confirm device status</li>
        <li>Update device compliance policies if needed</li>
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
    Write-Log "========== Starting Stale Devices and Accounts Monitoring =========="
    
    # Connect to Microsoft Graph
    Connect-MgGraphWithManagedIdentity
    
    # Retrieve stale devices and users
    $staleDevices = Get-StaleDevices -DaysInactive $StaleDeviceThreshold
    $staleUsers = Get-StaleUsers -DaysInactive $StaleUserThreshold
    
    # Export to Excel
    Export-ToExcelWithFormatting -StaleDevices $staleDevices -StaleUsers $staleUsers -FilePath $FullReportPath
    
    # Send email notification
    Send-EmailWithAttachment -Recipients $EmailRecipients -Subject $EmailSubject -AttachmentPath $FullReportPath `
                             -StaleDeviceCount $staleDevices.Count -StaleUserCount $staleUsers.Count
    
    # Cleanup
    if (Test-Path $FullReportPath) {
        Remove-Item $FullReportPath -Force
        Write-Log "Temporary report file cleaned up"
    }
    
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Found $($staleDevices.Count) stale devices and $($staleUsers.Count) stale users. Report sent to $($EmailRecipients -join ', ')"
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
