#Requires -Modules @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Reports"; ModuleVersion="2.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Monitors and reports on administrative role changes in Azure AD/Entra ID.

.DESCRIPTION
    This runbook tracks administrative role assignments and removals, identifying
    security-relevant changes. Generates an Excel report with detailed audit logs
    and sends email notifications for governance and compliance.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Graph API Permissions Required:
        - AuditLog.Read.All
        - Directory.Read.All
        - Mail.Send (if using email notifications)
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("security@yourdomain.com", "compliance@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "Administrative Changes Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Audit Log Settings
$DaysToQuery = 7  # Number of days to look back for changes
$HighPriorityRoles = @(
    'Global Administrator',
    'Privileged Role Administrator',
    'Security Administrator',
    'Conditional Access Administrator',
    'Exchange Administrator',
    'SharePoint Administrator'
)

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "AdminChanges_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-AdminRoleChanges {
    param([int]$Days)
    
    Write-Log "Retrieving admin role changes from audit logs..."
    
    try {
        $startDate = (Get-Date).AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $allAuditLogs = @()
        
        # Query audit logs for role assignments and removals
        $filter = "activityDateTime ge $startDate and (category eq 'RoleManagement')"
        $uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=$filter&`$top=999"
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            $allAuditLogs += $response.value
            $uri = $response.'@odata.nextLink'
            
            if ($uri) {
                Start-Sleep -Milliseconds 100  # Rate limiting
            }
        } while ($uri)
        
        Write-Log "Retrieved $($allAuditLogs.Count) audit log entries"
        
        # Process and filter relevant changes
        $roleChanges = @()
        
        foreach ($log in $allAuditLogs) {
            # Filter for role assignment/removal activities
            if ($log.activityDisplayName -notmatch 'Add member to role|Remove member from role') {
                continue
            }
            
            $changeType = if ($log.activityDisplayName -match 'Add') { 'Added' } else { 'Removed' }
            
            # Extract relevant information
            $actor = 'Unknown'
            $actorId = $null
            $targetUser = 'Unknown'
            $targetUserId = $null
            $roleName = 'Unknown'
            $roleId = $null
            
            # Parse initiatedBy
            if ($log.initiatedBy.user) {
                $actor = if ($log.initiatedBy.user.userPrincipalName) { 
                    $log.initiatedBy.user.userPrincipalName 
                } else { 
                    $log.initiatedBy.user.displayName 
                }
                $actorId = $log.initiatedBy.user.id
            }
            elseif ($log.initiatedBy.app) {
                $actor = "Application: $($log.initiatedBy.app.displayName)"
                $actorId = $log.initiatedBy.app.id
            }
            
            # Parse target resources
            foreach ($target in $log.targetResources) {
                if ($target.type -eq 'User') {
                    $targetUser = if ($target.userPrincipalName) { 
                        $target.userPrincipalName 
                    } else { 
                        $target.displayName 
                    }
                    $targetUserId = $target.id
                }
                elseif ($target.type -eq 'Role') {
                    $roleName = $target.displayName
                    $roleId = $target.id
                }
            }
            
            # Determine priority
            $priority = if ($HighPriorityRoles -contains $roleName) { 'HIGH' } else { 'NORMAL' }
            
            $roleChanges += [PSCustomObject]@{
                'Timestamp' = (Get-Date $log.activityDateTime).ToString('yyyy-MM-dd HH:mm:ss')
                'ChangeType' = $changeType
                'RoleName' = $roleName
                'TargetUser' = $targetUser
                'TargetUserId' = $targetUserId
                'PerformedBy' = $actor
                'PerformedById' = $actorId
                'Priority' = $priority
                'Result' = $log.result
                'ResultReason' = if ($log.resultReason) { $log.resultReason } else { 'N/A' }
                'CorrelationId' = $log.correlationId
                'LogId' = $log.id
            }
        }
        
        # Sort by timestamp (newest first)
        $roleChanges = $roleChanges | Sort-Object Timestamp -Descending
        
        Write-Log "Processed $($roleChanges.Count) role changes"
        Write-Log "  - High priority changes: $(($roleChanges | Where-Object {$_.Priority -eq 'HIGH'}).Count)"
        Write-Log "  - Assignments: $(($roleChanges | Where-Object {$_.ChangeType -eq 'Added'}).Count)"
        Write-Log "  - Removals: $(($roleChanges | Where-Object {$_.ChangeType -eq 'Removed'}).Count)"
        
        return $roleChanges
    }
    catch {
        Write-Log "Error retrieving admin role changes: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-CurrentAdminRoleMembers {
    Write-Log "Retrieving current admin role members..."
    
    try {
        $currentMembers = @()
        
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
                    $priority = if ($HighPriorityRoles -contains $roleName) { 'HIGH' } else { 'NORMAL' }
                    
                    $currentMembers += [PSCustomObject]@{
                        'RoleName' = $roleName
                        'DisplayName' = $member.displayName
                        'UserPrincipalName' = $member.userPrincipalName
                        'UserId' = $member.id
                        'Priority' = $priority
                    }
                }
                elseif ($member.'@odata.type' -eq '#microsoft.graph.servicePrincipal') {
                    $currentMembers += [PSCustomObject]@{
                        'RoleName' = $roleName
                        'DisplayName' = "Service Principal: $($member.displayName)"
                        'UserPrincipalName' = $member.appId
                        'UserId' = $member.id
                        'Priority' = 'SERVICE'
                    }
                }
            }
            
            Start-Sleep -Milliseconds 100  # Rate limiting
        }
        
        Write-Log "Retrieved $($currentMembers.Count) current admin role members"
        return $currentMembers
    }
    catch {
        Write-Log "Error retrieving current admin members: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$RoleChanges,
        [object[]]$CurrentMembers,
        [string]$FilePath,
        [int]$Days
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Create summary
        $totalChanges = $RoleChanges.Count
        $highPriorityChanges = ($RoleChanges | Where-Object {$_.Priority -eq 'HIGH'}).Count
        $assignments = ($RoleChanges | Where-Object {$_.ChangeType -eq 'Added'}).Count
        $removals = ($RoleChanges | Where-Object {$_.ChangeType -eq 'Removed'}).Count
        $totalCurrentAdmins = $CurrentMembers.Count
        $highPriorityAdmins = ($CurrentMembers | Where-Object {$_.Priority -eq 'HIGH'}).Count
        
        $summary = [PSCustomObject]@{
            'Metric' = @(
                '--- Recent Changes ---',
                'Time Period (Days)',
                'Total Role Changes',
                'High Priority Changes',
                'Role Assignments',
                'Role Removals',
                '--- Current State ---',
                'Total Admin Role Members',
                'High Priority Admins',
                'Report Generated'
            )
            'Value' = @(
                '',
                $Days,
                $totalChanges,
                $highPriorityChanges,
                $assignments,
                $removals,
                '',
                $totalCurrentAdmins,
                $highPriorityAdmins,
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Export high priority changes
        $highPriorityChanges = $RoleChanges | Where-Object {$_.Priority -eq 'HIGH'}
        if ($highPriorityChanges.Count -gt 0) {
            $highPriorityChanges | Export-Excel -Path $FilePath -WorksheetName "High Priority Changes" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "Added" -Range "B:B" -BackgroundColor LightGreen
                    New-ConditionalText -Text "Removed" -Range "B:B" -BackgroundColor LightCoral
                    New-ConditionalText -Text "HIGH" -Range "H:H" -BackgroundColor Orange
                )
        }
        
        # Export all role changes
        if ($RoleChanges.Count -gt 0) {
            $RoleChanges | Export-Excel -Path $FilePath -WorksheetName "All Changes" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "Added" -Range "B:B" -BackgroundColor LightGreen
                    New-ConditionalText -Text "Removed" -Range "B:B" -BackgroundColor LightCoral
                    New-ConditionalText -Text "HIGH" -Range "H:H" -BackgroundColor Orange
                    New-ConditionalText -Text "failure" -Range "I:I" -BackgroundColor Red
                )
        }
        
        # Export changes by role
        if ($RoleChanges.Count -gt 0) {
            $changesByRole = $RoleChanges | Group-Object RoleName | ForEach-Object {
                [PSCustomObject]@{
                    'RoleName' = $_.Name
                    'TotalChanges' = $_.Count
                    'Assignments' = ($_.Group | Where-Object {$_.ChangeType -eq 'Added'}).Count
                    'Removals' = ($_.Group | Where-Object {$_.ChangeType -eq 'Removed'}).Count
                    'Priority' = if ($HighPriorityRoles -contains $_.Name) { 'HIGH' } else { 'NORMAL' }
                }
            } | Sort-Object TotalChanges -Descending
            
            $changesByRole | Export-Excel -Path $FilePath -WorksheetName "Changes by Role" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "HIGH" -Range "E:E" -BackgroundColor Orange
                )
        }
        
        # Export current admin members
        if ($CurrentMembers.Count -gt 0) {
            $CurrentMembers | Sort-Object RoleName, DisplayName | 
                Export-Excel -Path $FilePath -WorksheetName "Current Admins" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "HIGH" -Range "E:E" -BackgroundColor LightYellow
                    New-ConditionalText -Text "SERVICE" -Range "E:E" -BackgroundColor LightBlue
                )
        }
        
        # Export high priority role members
        $highPriorityMembers = $CurrentMembers | Where-Object {$_.Priority -eq 'HIGH'} | Sort-Object RoleName, DisplayName
        if ($highPriorityMembers.Count -gt 0) {
            $highPriorityMembers | Export-Excel -Path $FilePath -WorksheetName "High Priority Admins" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
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
        [int]$TotalChanges,
        [int]$HighPriorityChanges,
        [int]$Days
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
        .total { color: #337ab7; font-weight: bold; font-size: 24px; }
        .high-priority { color: #f0ad4e; font-weight: bold; font-size: 24px; }
        .alert { background-color: #fcf8e3; border-left: 4px solid #f0ad4e; padding: 10px; margin: 20px 0; }
        .footer { margin-top: 30px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <h2>Administrative Changes Report</h2>
    <p>This automated report tracks administrative role assignments and removals for security and compliance.</p>
    
    <div class="summary">
        <div class="metric">
            <strong>Total Role Changes ($Days days):</strong> <span class="total">$TotalChanges</span>
        </div>
        <div class="metric">
            <strong>High Priority Role Changes:</strong> <span class="high-priority">$HighPriorityChanges</span>
        </div>
    </div>
    
    $(if ($HighPriorityChanges -gt 0) {
        @"
    <div class="alert">
        <strong>⚠️ HIGH PRIORITY CHANGES DETECTED</strong><br/>
        Changes to privileged roles (Global Admin, Security Admin, etc.) require review.
    </div>
"@
    })
    
    <p><strong>High Priority Roles Monitored:</strong></p>
    <ul>
        $(foreach ($role in $HighPriorityRoles) { "<li>$role</li>" })
    </ul>
    
    <p><strong>Report Details:</strong></p>
    <ul>
        <li>Time Period: Last $Days days</li>
        <li>Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</li>
        <li>Detailed audit logs are attached in Excel format</li>
    </ul>
    
    <p><strong>Recommended Actions:</strong></p>
    <ul>
        <li>Review all high priority role changes</li>
        <li>Verify changes were authorized and documented</li>
        <li>Ensure least privilege principles are followed</li>
        <li>Confirm removed users no longer require access</li>
        <li>Update access reviews and documentation</li>
        <li>Investigate any unexpected or unauthorized changes</li>
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
    Write-Log "========== Starting Administrative Changes Monitoring =========="
    
    # Connect to Microsoft Graph
    Connect-MgGraphWithManagedIdentity
    
    # Retrieve data
    $roleChanges = Get-AdminRoleChanges -Days $DaysToQuery
    $currentMembers = Get-CurrentAdminRoleMembers
    
    # Calculate metrics
    $highPriorityChanges = ($roleChanges | Where-Object {$_.Priority -eq 'HIGH'}).Count
    
    # Export to Excel
    Export-ToExcelWithFormatting -RoleChanges $roleChanges -CurrentMembers $currentMembers `
                                  -FilePath $FullReportPath -Days $DaysToQuery
    
    # Send email notification
    Send-EmailWithAttachment -Recipients $EmailRecipients -Subject $EmailSubject -AttachmentPath $FullReportPath `
                             -TotalChanges $roleChanges.Count -HighPriorityChanges $highPriorityChanges `
                             -Days $DaysToQuery
    
    # Cleanup
    if (Test-Path $FullReportPath) {
        Remove-Item $FullReportPath -Force
        Write-Log "Temporary report file cleaned up"
    }
    
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Found $($roleChanges.Count) role changes. High priority: $highPriorityChanges"
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
