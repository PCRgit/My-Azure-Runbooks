<#
.SYNOPSIS
    Tracks newly provisioned Azure resources by administrators and users
.DESCRIPTION
    Monitors Azure Activity Log for resource creation events
    Identifies who created what resources and when
    Provides detailed Excel report with creator attribution
.NOTES
    Author: Jaimin
    Requires: Az.Monitor, Az.Accounts, Az.Resources
    Authentication: Managed Identity
    Graph API Permissions: Mail.Send
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$Days = 7,
    
    [Parameter(Mandatory = $false)]
    [string[]]$SubscriptionIds,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipients = "ops-team@something.com",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:TEMP\NewAzureResources_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Connect to Azure
try {
    Write-Output "Connecting to Azure with Managed Identity..."
    Connect-AzAccount -Identity | Out-Null
    Write-Output "Successfully connected to Azure"
}
catch {
    Write-Error "Failed to connect to Azure: $_"
    throw
}

# Get Graph token for email
try {
    $GraphToken = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
    $Headers = @{
        "Authorization" = "Bearer $GraphToken"
        "Content-Type"  = "application/json"
    }
}
catch {
    Write-Warning "Could not get Graph token for email"
    $GraphToken = $null
}

# Get subscriptions
if (-not $SubscriptionIds) {
    $Subscriptions = Get-AzSubscription | Where-Object { $_.State -eq 'Enabled' }
}
else {
    $Subscriptions = $SubscriptionIds | ForEach-Object { Get-AzSubscription -SubscriptionId $_ }
}

Write-Output "Analyzing $($Subscriptions.Count) subscription(s) for new resources in the last $Days days..."

# Calculate time range
$StartTime = (Get-Date).AddDays(-$Days)
$EndTime = Get-Date

# Initialize collection
$NewResources = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($Subscription in $Subscriptions) {
    Write-Output "Processing subscription: $($Subscription.Name)"
    Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
    
    try {
        # Query Activity Log for resource creation events
        Write-Output "Querying activity logs..."
        
        $ActivityLogs = Get-AzActivityLog `
            -StartTime $StartTime `
            -EndTime $EndTime `
            -Status 'Succeeded' `
            -MaxRecord 10000
        
        # Filter for resource creation events
        $CreateEvents = $ActivityLogs | Where-Object { 
            $_.OperationName.Value -match '/write$|/create$' -and
            $_.EventName.Value -eq 'EndRequest' -and
            $_.ResourceId -match '/providers/'
        }
        
        Write-Output "Found $($CreateEvents.Count) resource creation events"
        
        foreach ($Event in $CreateEvents) {
            # Extract resource details
            $ResourceId = $Event.ResourceId
            $ResourceParts = $ResourceId -split '/'
            
            # Parse resource information
            $ResourceGroup = if ($ResourceId -match '/resourceGroups/([^/]+)') { $matches[1] } else { 'N/A' }
            $ResourceProvider = if ($ResourceId -match '/providers/([^/]+)') { $matches[1] } else { 'N/A' }
            $ResourceType = if ($ResourceId -match '/providers/[^/]+/([^/]+)') { $matches[1] } else { 'N/A' }
            $ResourceName = $ResourceParts[-1]
            
            # Get creator information
            $Creator = $Event.Caller
            $CreatorType = if ($Creator -match '@') { 'User' } elseif ($Creator -match '^[0-9a-f-]{36}$') { 'Service Principal' } else { 'Unknown' }
            
            # Try to get resource tags if resource still exists
            $Tags = 'N/A'
            try {
                $Resource = Get-AzResource -ResourceId $ResourceId -ErrorAction SilentlyContinue
                if ($Resource.Tags) {
                    $Tags = ($Resource.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                }
            }
            catch {
                # Resource might have been deleted
            }
            
            $NewResources.Add([PSCustomObject]@{
                    SubscriptionName    = $Subscription.Name
                    SubscriptionId      = $Subscription.Id
                    CreatedDateTime     = $Event.EventTimestamp
                    ResourceGroup       = $ResourceGroup
                    ResourceProvider    = $ResourceProvider
                    ResourceType        = $ResourceType
                    ResourceName        = $ResourceName
                    Creator             = $Creator
                    CreatorType         = $CreatorType
                    Operation           = $Event.OperationName.LocalizedValue
                    Status              = $Event.Status.Value
                    Tags                = $Tags
                    ResourceId          = $ResourceId
                    CorrelationId       = $Event.CorrelationId
                })
        }
    }
    catch {
        Write-Warning "Failed to process subscription $($Subscription.Name): $_"
    }
}

Write-Output "Total new resources found: $($NewResources.Count)"

# Generate statistics
$ResourcesByType = $NewResources | Group-Object ResourceType | ForEach-Object {
    [PSCustomObject]@{
        ResourceType = $_.Name
        Count        = $_.Count
    }
} | Sort-Object Count -Descending

$ResourcesByCreator = $NewResources | Group-Object Creator | ForEach-Object {
    [PSCustomObject]@{
        Creator      = $_.Name
        CreatorType  = $_.Group[0].CreatorType
        ResourceCount = $_.Count
        ResourceTypes = (($_.Group | Select-Object -Unique ResourceType).ResourceType -join ', ')
    }
} | Sort-Object ResourceCount -Descending

$ResourcesBySubscription = $NewResources | Group-Object SubscriptionName | ForEach-Object {
    [PSCustomObject]@{
        SubscriptionName = $_.Name
        ResourceCount    = $_.Count
    }
} | Sort-Object ResourceCount -Descending

$DailyTrend = $NewResources | Group-Object { $_.CreatedDateTime.Date } | ForEach-Object {
    [PSCustomObject]@{
        Date          = $_.Name
        ResourceCount = $_.Count
    }
} | Sort-Object Date

# Generate Excel report
Write-Output "Generating Excel report..."

try {
    # Summary
    $Summary = [PSCustomObject]@{
        ReportDate            = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        AnalysisPeriodDays    = $Days
        TotalNewResources     = $NewResources.Count
        UniqueCreators        = ($NewResources | Select-Object -Unique Creator).Count
        UniqueResourceTypes   = ($NewResources | Select-Object -Unique ResourceType).Count
        UniqueSubscriptions   = ($NewResources | Select-Object -Unique SubscriptionName).Count
    }
    
    $Summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -BoldTopRow -FreezeTopRow
    
    if ($NewResources.Count -gt 0) {
        $NewResources | Sort-Object CreatedDateTime -Descending | Export-Excel -Path $OutputPath -WorksheetName "AllNewResources" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($ResourcesByType.Count -gt 0) {
        $ResourcesByType | Export-Excel -Path $OutputPath -WorksheetName "ByResourceType" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($ResourcesByCreator.Count -gt 0) {
        $ResourcesByCreator | Export-Excel -Path $OutputPath -WorksheetName "ByCreator" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($ResourcesBySubscription.Count -gt 0) {
        $ResourcesBySubscription | Export-Excel -Path $OutputPath -WorksheetName "BySubscription" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($DailyTrend.Count -gt 0) {
        $DailyTrend | Export-Excel -Path $OutputPath -WorksheetName "DailyTrend" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    Write-Output "Excel report generated: $OutputPath"
}
catch {
    Write-Error "Failed to generate Excel report: $_"
    throw
}

# Send email notification
if ($GraphToken -and (Test-Path $OutputPath)) {
    try {
        Write-Output "Sending email notification..."
        
        $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
        $FileBase64 = [System.Convert]::ToBase64String($FileBytes)
        $FileName = Split-Path $OutputPath -Leaf
        
        # Top resource types
        $TopResourceTypes = $ResourcesByType | Select-Object -First 5 | ForEach-Object {
            "<li><strong>$($_.ResourceType)</strong>: $($_.Count) resources</li>"
        }
        
        # Top creators
        $TopCreators = $ResourcesByCreator | Select-Object -First 5 | ForEach-Object {
            "<li><strong>$($_.Creator)</strong> ($($_.CreatorType)): $($_.ResourceCount) resources</li>"
        }
        
        $EmailBody = @"
<html>
<body>
<h2>New Azure Resources Report</h2>
<p>Report generated on: <strong>$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</strong></p>

<h3>Summary (Last $Days Days):</h3>
<ul>
    <li>Total New Resources: <strong>$($Summary.TotalNewResources)</strong></li>
    <li>Unique Creators: <strong>$($Summary.UniqueCreators)</strong></li>
    <li>Unique Resource Types: <strong>$($Summary.UniqueResourceTypes)</strong></li>
    <li>Subscriptions: <strong>$($Summary.UniqueSubscriptions)</strong></li>
</ul>

<h3>Top Resource Types:</h3>
<ul>
$($TopResourceTypes -join "`n")
</ul>

<h3>Top Creators:</h3>
<ul>
$($TopCreators -join "`n")
</ul>

<p>Please review the attached Excel workbook for complete resource provisioning details.</p>

<p><em>This is an automated report from Azure Automation.</em></p>
</body>
</html>
"@
        
        $EmailMessage = @{
            message = @{
                subject      = "📊 New Azure Resources Report - $($Summary.TotalNewResources) Resources Created"
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
    catch {
        Write-Warning "Failed to send email: $_"
    }
}

Write-Output "Script completed successfully"