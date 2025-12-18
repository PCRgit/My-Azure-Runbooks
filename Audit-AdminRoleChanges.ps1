<#
.SYNOPSIS
    Audits administrative role changes performed by users (excludes system changes)
.DESCRIPTION
    Tracks role assignments/removals to privileged Azure AD roles
    Filters out system-generated changes to focus on user-initiated actions
    Generates detailed audit trail with initiator information
.NOTES
    Author: Jaimin
    Requires: Microsoft.Graph.Reports, Microsoft.Graph.Identity.DirectoryManagement
    Authentication: Managed Identity
    Graph API Permissions: AuditLog.Read.All, Directory.Read.All, Mail.Send
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$AuditDays = 30,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipients = "security-team@something.com",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:TEMP\AdminRoleChanges_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Privileged roles to monitor
$PrivilegedRoles = @(
    'Global Administrator',
    'Privileged Role Administrator',
    'User Administrator',
    'Exchange Administrator',
    'SharePoint Administrator',
    'Security Administrator',
    'Compliance Administrator',
    'Application Administrator',
    'Cloud Application Administrator',
    'Authentication Administrator',
    'Privileged Authentication Administrator',
    'Billing Administrator',
    'Conditional Access Administrator'
)

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

# Calculate date filter
$StartDate = (Get-Date).AddDays(-$AuditDays).ToString('yyyy-MM-ddT00:00:00Z')
Write-Output "Retrieving audit logs from: $StartDate"

# Get directory audit logs for role assignments
Write-Output "Retrieving directory role assignment changes..."

$RoleChanges = [System.Collections.Generic.List[PSCustomObject]]::new()

# Filter for role assignment activities
$Filter = "activityDateTime ge $StartDate and (category eq 'RoleManagement')"

try {
    $AuditLogs = Get-MgAuditLogDirectoryAudit -Filter $Filter -All
    
    Write-Output "Retrieved $($AuditLogs.Count) audit log entries"
    
    foreach ($Log in $AuditLogs) {
        # Filter out system-initiated changes
        $InitiatedBy = $null
        $InitiatorType = $null
        $InitiatorName = $null
        $InitiatorUPN = $null
        
        if ($Log.InitiatedBy.User) {
            $InitiatedBy = $Log.InitiatedBy.User
            $InitiatorType = 'User'
            $InitiatorName = $InitiatedBy.DisplayName
            $InitiatorUPN = $InitiatedBy.UserPrincipalName
        }
        elseif ($Log.InitiatedBy.App) {
            # Skip system/app-initiated changes unless it's a user-consented app
            $InitiatedBy = $Log.InitiatedBy.App
            $InitiatorType = 'Application'
            $InitiatorName = $InitiatedBy.DisplayName
            
            # Skip known system apps
            $SystemApps = @(
                'Microsoft Azure AD', 
                'Azure Active Directory', 
                'Windows Azure Active Directory',
                'Microsoft.Azure.SyncFabric',
                'Azure AD Device Registration'
            )
            
            if ($InitiatorName -in $SystemApps) {
                continue  # Skip system-initiated changes
            }
        }
        else {
            continue  # Skip if no clear initiator
        }
        
        # Parse target resources
        foreach ($Target in $Log.TargetResources) {
            $RoleName = $null
            $TargetUserName = $null
            $TargetUserUPN = $null
            
            # Extract role name from modified properties
            $RoleProperty = $Target.ModifiedProperties | Where-Object { $_.DisplayName -eq 'Role.DisplayName' }
            if ($RoleProperty) {
                $RoleName = $RoleProperty.NewValue -replace '"', ''
            }
            
            # Get target user information
            if ($Target.Type -eq 'User') {
                $TargetUserName = $Target.DisplayName
                $TargetUserUPN = $Target.UserPrincipalName
            }
            
            # Filter for privileged roles only
            if ($RoleName -in $PrivilegedRoles) {
                $ActionType = switch ($Log.ActivityDisplayName) {
                    'Add member to role' { 'Role Assigned' }
                    'Remove member from role' { 'Role Removed' }
                    'Add eligible member to role' { 'Eligible Role Assigned (PIM)' }
                    'Remove eligible member from role' { 'Eligible Role Removed (PIM)' }
                    default { $Log.ActivityDisplayName }
                }
                
                $RoleChanges.Add([PSCustomObject]@{
                        Timestamp           = [datetime]$Log.ActivityDateTime
                        ActionType          = $ActionType
                        RoleName            = $RoleName
                        TargetUserName      = $TargetUserName
                        TargetUserUPN       = $TargetUserUPN
                        InitiatorType       = $InitiatorType
                        InitiatorName       = $InitiatorName
                        InitiatorUPN        = $InitiatorUPN
                        Result              = $Log.Result
                        AdditionalDetails   = ($Log.AdditionalDetails | ForEach-Object { "$($_.Key): $($_.Value)" }) -join '; '
                        CorrelationId       = $Log.CorrelationId
                    })
            }
        }
    }
}
catch {
    Write-Error "Failed to retrieve audit logs: $_"
    throw
}

Write-Output "Found $($RoleChanges.Count) user-initiated role changes"

# Generate statistics
$RoleAssignments = $RoleChanges | Where-Object { $_.ActionType -like '*Assigned*' }
$RoleRemovals = $RoleChanges | Where-Object { $_.ActionType -like '*Removed*' }

$ChangesByRole = $RoleChanges | Group-Object RoleName | ForEach-Object {
    [PSCustomObject]@{
        RoleName        = $_.Name
        TotalChanges    = $_.Count
        Assignments     = ($_.Group | Where-Object { $_.ActionType -like '*Assigned*' }).Count
        Removals        = ($_.Group | Where-Object { $_.ActionType -like '*Removed*' }).Count
    }
} | Sort-Object TotalChanges -Descending

$ChangesByInitiator = $RoleChanges | Group-Object InitiatorName | ForEach-Object {
    [PSCustomObject]@{
        InitiatorName   = $_.Name
        InitiatorType   = $_.Group[0].InitiatorType
        TotalChanges    = $_.Count
        Assignments     = ($_.Group | Where-Object { $_.ActionType -like '*Assigned*' }).Count
        Removals        = ($_.Group | Where-Object { $_.ActionType -like '*Removed*' }).Count
    }
} | Sort-Object TotalChanges -Descending

# Generate Excel report
Write-Output "Generating Excel report..."

try {
    # Summary
    $Summary = [PSCustomObject]@{
        ReportDate         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        AuditPeriodDays    = $AuditDays
        TotalChanges       = $RoleChanges.Count
        RoleAssignments    = $RoleAssignments.Count
        RoleRemovals       = $RoleRemovals.Count
        UniqueInitiators   = ($RoleChanges | Select-Object -Unique InitiatorName).Count
        UniqueTargetUsers  = ($RoleChanges | Select-Object -Unique TargetUserUPN).Count
    }
    
    $Summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -BoldTopRow -FreezeTopRow
    
    if ($RoleChanges.Count -gt 0) {
        $RoleChanges | Sort-Object Timestamp -Descending | Export-Excel -Path $OutputPath -WorksheetName "AllRoleChanges" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($ChangesByRole.Count -gt 0) {
        $ChangesByRole | Export-Excel -Path $OutputPath -WorksheetName "ChangesByRole" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($ChangesByInitiator.Count -gt 0) {
        $ChangesByInitiator | Export-Excel -Path $OutputPath -WorksheetName "ChangesByInitiator" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($RoleAssignments.Count -gt 0) {
        $RoleAssignments | Sort-Object Timestamp -Descending | Export-Excel -Path $OutputPath -WorksheetName "RoleAssignments" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($RoleRemovals.Count -gt 0) {
        $RoleRemovals | Sort-Object Timestamp -Descending | Export-Excel -Path $OutputPath -WorksheetName "RoleRemovals" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    Write-Output "Excel report generated: $OutputPath"
}
catch {
    Write-Error "Failed to generate Excel report: $_"
    throw
}

# Send email notification
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
        
        # Recent changes
        $RecentChanges = $RoleChanges | Sort-Object Timestamp -Descending | Select-Object -First 10 | ForEach-Object {
            "<li><strong>$($_.Timestamp.ToString('yyyy-MM-dd HH:mm'))</strong>: $($_.InitiatorName) - $($_.ActionType) for <strong>$($_.RoleName)</strong> to $($_.TargetUserName)</li>"
        }
        
        $EmailBody = @"
<html>
<body>
<h2>Administrative Role Changes Audit Report</h2>
<p>Report generated on: <strong>$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</strong></p>

<h3>Summary (Last $AuditDays Days):</h3>
<ul>
    <li>Total User-Initiated Changes: <strong>$($Summary.TotalChanges)</strong></li>
    <li>Role Assignments: <strong style="color:green;">$($Summary.RoleAssignments)</strong></li>
    <li>Role Removals: <strong style="color:red;">$($Summary.RoleRemovals)</strong></li>
    <li>Unique Initiators: <strong>$($Summary.UniqueInitiators)</strong></li>
    <li>Unique Target Users: <strong>$($Summary.UniqueTargetUsers)</strong></li>
</ul>

<h3>Recent Changes (Last 10):</h3>
<ul>
$($RecentChanges -join "`n")
</ul>

<p>Please review the attached Excel workbook for complete audit details.</p>

<p><em>This is an automated security audit from Azure Automation.</em></p>
</body>
</html>
"@
        
        $Headers = @{
            "Authorization" = "Bearer $GraphToken"
            "Content-Type"  = "application/json"
        }
        
        $EmailMessage = @{
            message = @{
                subject      = "🔒 Admin Role Changes Audit - $($Summary.TotalChanges) Changes Detected"
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