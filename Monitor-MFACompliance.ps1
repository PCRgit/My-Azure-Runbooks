#Requires -Modules @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Reports"; ModuleVersion="2.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Monitors and reports on Multi-Factor Authentication (MFA) compliance across the organization.

.DESCRIPTION
    This runbook identifies users who do not have MFA enabled or registered, breaking down
    by admin vs non-admin users. Generates an Excel report with conditional formatting
    and sends email notifications to improve security posture.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Graph API Permissions Required:
        - User.Read.All
        - UserAuthenticationMethod.Read.All
        - Directory.Read.All
        - Mail.Send (if using email notifications)
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("security@yourdomain.com", "admin@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "MFA Compliance Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Exclusion Settings
$ExcludeServiceAccounts = $true  # Exclude accounts with "svc" or "service" in name
$ExcludeDisabledUsers = $true    # Exclude disabled accounts

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "MFACompliance_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-AdminRoleMembers {
    Write-Log "Retrieving admin role members..."
    
    try {
        $adminUsers = @()
        
        # Get all directory roles
        $uri = "https://graph.microsoft.com/v1.0/directoryRoles"
        $roles = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        foreach ($role in $roles.value) {
            $roleName = $role.displayName
            $roleId = $role.id
            
            # Get members of this role
            $membersUri = "https://graph.microsoft.com/v1.0/directoryRoles/$roleId/members"
            $members = Invoke-MgGraphRequest -Uri $membersUri -Method GET
            
            foreach ($member in $members.value) {
                if ($member.'@odata.type' -eq '#microsoft.graph.user') {
                    $adminUsers += [PSCustomObject]@{
                        UserId = $member.id
                        UserPrincipalName = $member.userPrincipalName
                        Role = $roleName
                    }
                }
            }
            
            Start-Sleep -Milliseconds 100  # Rate limiting
        }
        
        Write-Log "Found $($adminUsers.Count) total admin role assignments"
        return $adminUsers
    }
    catch {
        Write-Log "Error retrieving admin roles: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-UserMFAStatus {
    Write-Log "Retrieving user MFA status..."
    
    try {
        $allUsers = @()
        $mfaStatus = @()
        
        # Get all users
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime"
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allUsers += $response.value
            $uri = $response.'@odata.nextLink'
            
            if ($uri) {
                Start-Sleep -Milliseconds 100  # Rate limiting
            }
        } while ($uri)
        
        Write-Log "Retrieved $($allUsers.Count) total users"
        
        # Get admin users for comparison
        $adminUsers = Get-AdminRoleMembers
        $adminUserIds = $adminUsers.UserId | Select-Object -Unique
        
        # Process each user
        $counter = 0
        foreach ($user in $allUsers) {
            $counter++
            
            # Apply filters
            if ($ExcludeDisabledUsers -and -not $user.accountEnabled) {
                continue
            }
            
            if ($user.userType -eq 'Guest') {
                continue
            }
            
            if ($ExcludeServiceAccounts -and 
                ($user.userPrincipalName -like '*svc*' -or 
                 $user.userPrincipalName -like '*service*' -or
                 $user.displayName -like '*service*')) {
                continue
            }
            
            # Check if user is an admin
            $isAdmin = $adminUserIds -contains $user.id
            $adminRoles = if ($isAdmin) {
                ($adminUsers | Where-Object {$_.UserId -eq $user.id} | Select-Object -ExpandProperty Role) -join '; '
            } else {
                'N/A'
            }
            
            # Get user's authentication methods
            try {
                $authMethodsUri = "https://graph.microsoft.com/v1.0/users/$($user.id)/authentication/methods"
                $authMethods = Invoke-MgGraphRequest -Uri $authMethodsUri -Method GET
                
                # Check for MFA-capable methods
                $hasMFA = $false
                $mfaMethods = @()
                
                foreach ($method in $authMethods.value) {
                    $methodType = $method.'@odata.type'
                    
                    switch ($methodType) {
                        '#microsoft.graph.phoneAuthenticationMethod' {
                            $hasMFA = $true
                            $mfaMethods += 'Phone'
                        }
                        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                            $hasMFA = $true
                            $mfaMethods += 'Authenticator App'
                        }
                        '#microsoft.graph.fido2AuthenticationMethod' {
                            $hasMFA = $true
                            $mfaMethods += 'FIDO2 Security Key'
                        }
                        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' {
                            $hasMFA = $true
                            $mfaMethods += 'Windows Hello'
                        }
                        '#microsoft.graph.emailAuthenticationMethod' {
                            $mfaMethods += 'Email (Not MFA)'
                        }
                        '#microsoft.graph.passwordAuthenticationMethod' {
                            $mfaMethods += 'Password'
                        }
                    }
                }
                
                $mfaStatus += [PSCustomObject]@{
                    'DisplayName' = $user.displayName
                    'UserPrincipalName' = $user.userPrincipalName
                    'IsAdmin' = $isAdmin
                    'AdminRoles' = $adminRoles
                    'MFAEnabled' = $hasMFA
                    'MFAMethods' = if ($mfaMethods.Count -gt 0) { $mfaMethods -join ', ' } else { 'None' }
                    'AccountEnabled' = $user.accountEnabled
                    'CreatedDate' = (Get-Date $user.createdDateTime).ToString('yyyy-MM-dd')
                    'UserId' = $user.id
                }
                
                # Rate limiting
                if ($counter % 50 -eq 0) {
                    Write-Log "Processed $counter of $($allUsers.Count) users..."
                    Start-Sleep -Milliseconds 500
                }
                else {
                    Start-Sleep -Milliseconds 100
                }
            }
            catch {
                Write-Log "Error retrieving auth methods for $($user.userPrincipalName): $($_.Exception.Message)" -Level WARNING
                
                $mfaStatus += [PSCustomObject]@{
                    'DisplayName' = $user.displayName
                    'UserPrincipalName' = $user.userPrincipalName
                    'IsAdmin' = $isAdmin
                    'AdminRoles' = $adminRoles
                    'MFAEnabled' = 'Unknown'
                    'MFAMethods' = 'Error retrieving'
                    'AccountEnabled' = $user.accountEnabled
                    'CreatedDate' = (Get-Date $user.createdDateTime).ToString('yyyy-MM-dd')
                    'UserId' = $user.id
                }
                
                Start-Sleep -Milliseconds 100
            }
        }
        
        Write-Log "Processed $($mfaStatus.Count) users for MFA compliance"
        Write-Log "  - Admins without MFA: $(($mfaStatus | Where-Object {$_.IsAdmin -and -not $_.MFAEnabled}).Count)"
        Write-Log "  - Non-admins without MFA: $(($mfaStatus | Where-Object {-not $_.IsAdmin -and -not $_.MFAEnabled}).Count)"
        
        return $mfaStatus
    }
    catch {
        Write-Log "Error retrieving user MFA status: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$MFAStatus,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Calculate summary statistics
        $totalUsers = $MFAStatus.Count
        $adminsTotal = ($MFAStatus | Where-Object {$_.IsAdmin}).Count
        $adminsWithMFA = ($MFAStatus | Where-Object {$_.IsAdmin -and $_.MFAEnabled}).Count
        $adminsWithoutMFA = ($MFAStatus | Where-Object {$_.IsAdmin -and -not $_.MFAEnabled}).Count
        
        $nonAdminsTotal = ($MFAStatus | Where-Object {-not $_.IsAdmin}).Count
        $nonAdminsWithMFA = ($MFAStatus | Where-Object {-not $_.IsAdmin -and $_.MFAEnabled}).Count
        $nonAdminsWithoutMFA = ($MFAStatus | Where-Object {-not $_.IsAdmin -and -not $_.MFAEnabled}).Count
        
        $adminComplianceRate = if ($adminsTotal -gt 0) { [math]::Round(($adminsWithMFA / $adminsTotal) * 100, 2) } else { 0 }
        $nonAdminComplianceRate = if ($nonAdminsTotal -gt 0) { [math]::Round(($nonAdminsWithMFA / $nonAdminsTotal) * 100, 2) } else { 0 }
        $overallComplianceRate = if ($totalUsers -gt 0) { [math]::Round((($adminsWithMFA + $nonAdminsWithMFA) / $totalUsers) * 100, 2) } else { 0 }
        
        # Create summary
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total Users Analyzed',
                '--- Admin Users ---',
                'Total Admins',
                'Admins with MFA',
                'Admins without MFA',
                'Admin MFA Compliance Rate',
                '--- Non-Admin Users ---',
                'Total Non-Admins',
                'Non-Admins with MFA',
                'Non-Admins without MFA',
                'Non-Admin MFA Compliance Rate',
                '--- Overall ---',
                'Overall MFA Compliance Rate',
                'Report Generated'
            )
            'Value' = @(
                $totalUsers,
                '',
                $adminsTotal,
                $adminsWithMFA,
                $adminsWithoutMFA,
                "$adminComplianceRate%",
                '',
                $nonAdminsTotal,
                $nonAdminsWithMFA,
                $nonAdminsWithoutMFA,
                "$nonAdminComplianceRate%",
                '',
                "$overallComplianceRate%",
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Top 20 admins without MFA (highest priority)
        $adminsNoMFA = $MFAStatus | Where-Object {$_.IsAdmin -and -not $_.MFAEnabled} | Select-Object -First 20
        if ($adminsNoMFA.Count -gt 0) {
            $adminsNoMFA | Export-Excel -Path $FilePath -WorksheetName "Admins Without MFA" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # All users without MFA
        $usersNoMFA = $MFAStatus | Where-Object {-not $_.MFAEnabled}
        if ($usersNoMFA.Count -gt 0) {
            $usersNoMFA | Export-Excel -Path $FilePath -WorksheetName "All Without MFA" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "True" -Range "C:C" -BackgroundColor Orange
                    New-ConditionalText -Text "False" -Range "C:C" -BackgroundColor LightGray
                )
        }
        
        # All users with MFA status
        $MFAStatus | Export-Excel -Path $FilePath -WorksheetName "All Users" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
            -ConditionalText $(
                New-ConditionalText -Text "True" -Range "E:E" -BackgroundColor LightGreen
                New-ConditionalText -Text "False" -Range "E:E" -BackgroundColor LightCoral
                New-ConditionalText -Text "True" -Range "C:C" -BackgroundColor LightYellow
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
        [int]$AdminsWithoutMFA,
        [int]$NonAdminsWithoutMFA,
        [decimal]$ComplianceRate
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
        .critical { color: #d9534f; font-weight: bold; font-size: 24px; }
        .warning { color: #f0ad4e; font-weight: bold; font-size: 20px; }
        .good { color: #5cb85c; font-weight: bold; font-size: 24px; }
        .alert { background-color: #f2dede; border-left: 4px solid #d9534f; padding: 10px; margin: 20px 0; }
        .footer { margin-top: 30px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <h2>Multi-Factor Authentication Compliance Report</h2>
    <p>This automated report identifies users who do not have MFA enabled.</p>
    
    <div class="summary">
        <div class="metric">
            <strong>Overall MFA Compliance Rate:</strong> <span class="$(if ($ComplianceRate -ge 95) { 'good' } elseif ($ComplianceRate -ge 80) { 'warning' } else { 'critical' })">$ComplianceRate%</span>
        </div>
        <div class="metric">
            <strong>Admins Without MFA:</strong> <span class="critical">$AdminsWithoutMFA</span>
        </div>
        <div class="metric">
            <strong>Non-Admins Without MFA:</strong> <span class="warning">$NonAdminsWithoutMFA</span>
        </div>
    </div>
    
    $(if ($AdminsWithoutMFA -gt 0) {
        @"
    <div class="alert">
        <strong>🚨 CRITICAL SECURITY RISK</strong><br/>
        There are $AdminsWithoutMFA administrator accounts without MFA enabled. These accounts should be prioritized immediately.
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
        <li>Immediately enable MFA for all administrator accounts</li>
        <li>Enforce MFA through Conditional Access policies</li>
        <li>Provide MFA registration support to users</li>
        <li>Consider phishing-resistant methods (FIDO2, Windows Hello)</li>
        <li>Document and communicate MFA requirements</li>
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
    Write-Log "========== Starting MFA Compliance Monitoring =========="
    
    # Connect to Microsoft Graph
    Connect-MgGraphWithManagedIdentity
    
    # Retrieve MFA status
    $mfaStatus = Get-UserMFAStatus
    
    # Calculate metrics for email
    $adminsWithoutMFA = ($mfaStatus | Where-Object {$_.IsAdmin -and -not $_.MFAEnabled}).Count
    $nonAdminsWithoutMFA = ($mfaStatus | Where-Object {-not $_.IsAdmin -and -not $_.MFAEnabled}).Count
    $totalUsers = $mfaStatus.Count
    $usersWithMFA = ($mfaStatus | Where-Object {$_.MFAEnabled}).Count
    $complianceRate = if ($totalUsers -gt 0) { [math]::Round(($usersWithMFA / $totalUsers) * 100, 2) } else { 0 }
    
    # Export to Excel
    Export-ToExcelWithFormatting -MFAStatus $mfaStatus -FilePath $FullReportPath
    
    # Send email notification
    Send-EmailWithAttachment -Recipients $EmailRecipients -Subject $EmailSubject -AttachmentPath $FullReportPath `
                             -AdminsWithoutMFA $adminsWithoutMFA -NonAdminsWithoutMFA $nonAdminsWithoutMFA `
                             -ComplianceRate $complianceRate
    
    # Cleanup
    if (Test-Path $FullReportPath) {
        Remove-Item $FullReportPath -Force
        Write-Log "Temporary report file cleaned up"
    }
    
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: MFA Compliance Rate: $complianceRate%. Admins without MFA: $adminsWithoutMFA, Non-admins without MFA: $nonAdminsWithoutMFA"
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
