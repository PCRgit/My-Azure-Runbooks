<#
.SYNOPSIS
    Monitors and reports Entra ID group membership changes
.DESCRIPTION
    Tracks group membership additions/removals and compares against previous baseline
    Stores baseline in Azure Storage Table for historical tracking
.NOTES
    Author: Jaimin
    Requires: Microsoft.Graph.Groups, Microsoft.Graph.Users modules
    Authentication: Managed Identity
    Graph API Permissions: Group.Read.All, User.Read.All, Mail.Send
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$MonitoredGroupNames,  # If empty, monitors ALL groups
    
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipients = "identity-team@something.com",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:TEMP\GroupMembershipChanges_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    
    [Parameter(Mandatory = $false)]
    [string]$BaselineStoragePath = "$env:TEMP\GroupMembershipBaseline.json"
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

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

# Get Graph token for email
try {
    $GraphToken = (Get-MgContext).TokenCredential.GetTokenAsync(
        (New-Object System.Threading.CancellationToken), 
        (New-Object Microsoft.Identity.Client.AuthenticationProviderOption)
    ).Result.Token
    
    $Headers = @{
        "Authorization" = "Bearer $GraphToken"
        "Content-Type"  = "application/json"
    }
}
catch {
    Write-Warning "Could not get token for email functionality"
    $GraphToken = $null
}

# Get groups to monitor
if ($MonitoredGroupNames) {
    $Groups = @()
    foreach ($GroupName in $MonitoredGroupNames) {
        $Group = Get-MgGroup -Filter "displayName eq '$GroupName'" -All
        if ($Group) {
            $Groups += $Group
        }
    }
}
else {
    Write-Output "Retrieving all Entra ID groups..."
    $Groups = Get-MgGroup -All -Property Id, DisplayName, GroupTypes, SecurityEnabled, MailEnabled
}

Write-Output "Monitoring $($Groups.Count) group(s)..."

# Current membership snapshot
$CurrentMembership = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($Group in $Groups) {
    Write-Output "Processing group: $($Group.DisplayName)"
    
    try {
        $Members = Get-MgGroupMember -GroupId $Group.Id -All
        
        foreach ($Member in $Members) {
            # Get member details
            $MemberDetails = $null
            if ($Member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                $MemberDetails = Get-MgUser -UserId $Member.Id -Property DisplayName, UserPrincipalName, Mail
                $MemberType = 'User'
            }
            elseif ($Member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                $MemberDetails = Get-MgGroup -GroupId $Member.Id
                $MemberType = 'Group'
            }
            elseif ($Member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.servicePrincipal') {
                $MemberType = 'ServicePrincipal'
            }
            else {
                $MemberType = 'Other'
            }
            
            $CurrentMembership.Add([PSCustomObject]@{
                    GroupId          = $Group.Id
                    GroupName        = $Group.DisplayName
                    MemberId         = $Member.Id
                    MemberName       = if ($MemberDetails) { $MemberDetails.DisplayName } else { $Member.Id }
                    MemberUPN        = if ($MemberDetails.UserPrincipalName) { $MemberDetails.UserPrincipalName } else { 'N/A' }
                    MemberType       = $MemberType
                    SnapshotDate     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                })
        }
    }
    catch {
        Write-Warning "Failed to get members for group $($Group.DisplayName): $_"
    }
    
    # Rate limiting
    Start-Sleep -Milliseconds 100
}

# Load previous baseline
$PreviousMembership = @()
if (Test-Path $BaselineStoragePath) {
    Write-Output "Loading previous baseline from $BaselineStoragePath"
    $PreviousMembership = Get-Content $BaselineStoragePath | ConvertFrom-Json
}
else {
    Write-Output "No previous baseline found. This will be the first baseline."
}

# Detect changes
$AddedMembers = [System.Collections.Generic.List[PSCustomObject]]::new()
$RemovedMembers = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($PreviousMembership.Count -gt 0) {
    Write-Output "Comparing membership changes..."
    
    # Find additions
    foreach ($Current in $CurrentMembership) {
        $Key = "$($Current.GroupId)-$($Current.MemberId)"
        $PreviousKey = "$($PreviousMembership.GroupId)-$($PreviousMembership.MemberId)"
        
        if ($Key -notin $PreviousKey) {
            $AddedMembers.Add([PSCustomObject]@{
                    ChangeType   = 'Added'
                    DetectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    GroupName    = $Current.GroupName
                    MemberName   = $Current.MemberName
                    MemberUPN    = $Current.MemberUPN
                    MemberType   = $Current.MemberType
                })
        }
    }
    
    # Find removals
    foreach ($Previous in $PreviousMembership) {
        $Key = "$($Previous.GroupId)-$($Previous.MemberId)"
        $CurrentKey = "$($CurrentMembership.GroupId)-$($CurrentMembership.MemberId)"
        
        if ($Key -notin $CurrentKey) {
            $RemovedMembers.Add([PSCustomObject]@{
                    ChangeType   = 'Removed'
                    DetectedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    GroupName    = $Previous.GroupName
                    MemberName   = $Previous.MemberName
                    MemberUPN    = $Previous.MemberUPN
                    MemberType   = $Previous.MemberType
                })
        }
    }
}

# Combine all changes
$AllChanges = @()
$AllChanges += $AddedMembers
$AllChanges += $RemovedMembers

Write-Output "Detected changes: $($AddedMembers.Count) additions, $($RemovedMembers.Count) removals"

# Save current state as new baseline
Write-Output "Saving new baseline..."
$CurrentMembership | ConvertTo-Json -Depth 10 | Set-Content $BaselineStoragePath

# Generate report if changes detected
if ($AllChanges.Count -gt 0) {
    Write-Output "Generating change report..."
    
    try {
        # Summary
        $Summary = [PSCustomObject]@{
            ReportDate     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            TotalGroups    = $Groups.Count
            TotalAdditions = $AddedMembers.Count
            TotalRemovals  = $RemovedMembers.Count
            TotalChanges   = $AllChanges.Count
        }
        
        $Summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -BoldTopRow -FreezeTopRow
        
        if ($AllChanges.Count -gt 0) {
            $AllChanges | Sort-Object GroupName, ChangeType | Export-Excel -Path $OutputPath -WorksheetName "AllChanges" -AutoSize -BoldTopRow -FreezeTopRow -Append
        }
        
        if ($AddedMembers.Count -gt 0) {
            $AddedMembers | Sort-Object GroupName | Export-Excel -Path $OutputPath -WorksheetName "Additions" -AutoSize -BoldTopRow -FreezeTopRow -Append
        }
        
        if ($RemovedMembers.Count -gt 0) {
            $RemovedMembers | Sort-Object GroupName | Export-Excel -Path $OutputPath -WorksheetName "Removals" -AutoSize -BoldTopRow -FreezeTopRow -Append
        }
        
        Write-Output "Report generated: $OutputPath"
    }
    catch {
        Write-Error "Failed to generate Excel report: $_"
    }
    
    # Send email notification
    if ($GraphToken -and (Test-Path $OutputPath)) {
        try {
            Write-Output "Sending change notification email..."
            
            $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
            $FileBase64 = [System.Convert]::ToBase64String($FileBytes)
            $FileName = Split-Path $OutputPath -Leaf
            
            # Generate top changes list
            $TopAdditions = $AddedMembers | Select-Object -First 10 | ForEach-Object {
                "<li><strong>$($_.MemberName)</strong> added to <strong>$($_.GroupName)</strong></li>"
            }
            
            $TopRemovals = $RemovedMembers | Select-Object -First 10 | ForEach-Object {
                "<li><strong>$($_.MemberName)</strong> removed from <strong>$($_.GroupName)</strong></li>"
            }
            
            $EmailBody = @"
<html>
<body>
<h2>Entra ID Group Membership Changes Detected</h2>
<p>Report generated on: <strong>$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</strong></p>

<h3>Summary:</h3>
<ul>
    <li>Total Groups Monitored: <strong>$($Summary.TotalGroups)</strong></li>
    <li>Member Additions: <strong style="color:green;">$($Summary.TotalAdditions)</strong></li>
    <li>Member Removals: <strong style="color:red;">$($Summary.TotalRemovals)</strong></li>
    <li>Total Changes: <strong>$($Summary.TotalChanges)</strong></li>
</ul>

$(if ($TopAdditions) {
@"
<h3>Recent Additions (Top 10):</h3>
<ul>
$($TopAdditions -join "`n")
</ul>
"@
})

$(if ($TopRemovals) {
@"
<h3>Recent Removals (Top 10):</h3>
<ul>
$($TopRemovals -join "`n")
</ul>
"@
})

<p>Please review the attached Excel workbook for complete details.</p>

<p><em>This is an automated report from Azure Automation.</em></p>
</body>
</html>
"@
            
            $EmailMessage = @{
                message = @{
                    subject      = "⚠️ Entra ID Group Membership Changes Detected - $(Get-Date -Format 'yyyy-MM-dd')"
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
                -Body $EmailJson `
                -ContentType "application/json"
            
            Write-Output "Email sent successfully"
        }
        catch {
            Write-Warning "Failed to send email: $_"
        }
    }
}
else {
    Write-Output "No membership changes detected since last run."
}

Disconnect-MgGraph | Out-Null
Write-Output "Script completed successfully"