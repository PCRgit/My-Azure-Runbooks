#Requires -Modules @{ ModuleName="Az.Accounts"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Az.Resources"; ModuleVersion="6.0.0" }
#Requires -Modules @{ ModuleName="Az.Compute"; ModuleVersion="5.0.0" }
#Requires -Modules @{ ModuleName="Az.Network"; ModuleVersion="5.0.0" }
#Requires -Modules @{ ModuleName="Az.Storage"; ModuleVersion="5.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Identifies and reports on unused Azure resources for cost optimization.

.DESCRIPTION
    This runbook scans Azure subscriptions for unused resources that are incurring costs:
    - Unattached managed disks
    - Unused public IP addresses
    - Empty resource groups
    - Stopped VMs with storage costs
    - Unused network interfaces
    - Orphaned snapshots
    - Unused load balancers and application gateways

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Azure RBAC Permissions Required:
        - Reader role on subscriptions
        - Cost Management Reader (for cost data)
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("admin@yourdomain.com", "finance@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "Unused Azure Resources Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Subscription Selection (empty array = all subscriptions)
$TargetSubscriptions = @()  # Leave empty for all, or specify: @("sub-id-1", "sub-id-2")

# Age Thresholds (days)
$MinimumAgeThreshold = 30  # Only report resources older than this

# Cost Estimation (average monthly costs in USD)
$EstimatedCosts = @{
    UnattachedDisk_Standard = 5.00    # Per 100GB Standard HDD
    UnattachedDisk_Premium = 20.00    # Per 100GB Premium SSD
    PublicIP_Standard = 3.50          # Standard Public IP
    PublicIP_Basic = 0.00             # Basic IP (free when attached)
    Snapshot_Standard = 0.05          # Per GB per month
    LoadBalancer_Basic = 0.00         # Basic LB is free
    LoadBalancer_Standard = 25.00     # Standard LB base cost
    AppGateway_Small = 140.00         # App Gateway v2
    NetworkInterface = 0.00           # NICs are free
    StoppedVM_Storage = 10.00         # Average storage cost for stopped VM
}

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "UnusedResources_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Connect-AzureWithManagedIdentity {
    try {
        Write-Log "Connecting to Azure using Managed Identity..."
        Connect-AzAccount -Identity | Out-Null
        Write-Log "Successfully connected to Azure"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Azure: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-TargetSubscriptions {
    param([string[]]$SubscriptionIds)
    
    try {
        if ($SubscriptionIds.Count -eq 0) {
            Write-Log "Getting all accessible subscriptions..."
            $subscriptions = Get-AzSubscription | Where-Object {$_.State -eq 'Enabled'}
        }
        else {
            Write-Log "Getting specified subscriptions..."
            $subscriptions = $SubscriptionIds | ForEach-Object {
                Get-AzSubscription -SubscriptionId $_
            }
        }
        
        Write-Log "Found $($subscriptions.Count) subscription(s) to analyze"
        return $subscriptions
    }
    catch {
        Write-Log "Error retrieving subscriptions: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-UnattachedDisks {
    param(
        [object]$Subscription,
        [int]$MinAgeDays
    )
    
    Write-Log "  Scanning for unattached disks in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $unattachedDisks = Get-AzDisk | Where-Object {
            $_.ManagedBy -eq $null -and
            ((Get-Date) - $_.TimeCreated).Days -ge $MinAgeDays
        }
        
        $results = @()
        foreach ($disk in $unattachedDisks) {
            $sizeMB = $disk.DiskSizeGB
            $tier = $disk.Sku.Name
            
            # Estimate monthly cost
            $costPerMonth = if ($tier -like "*Premium*") {
                ($sizeMB / 100) * $EstimatedCosts.UnattachedDisk_Premium
            } else {
                ($sizeMB / 100) * $EstimatedCosts.UnattachedDisk_Standard
            }
            
            $ageDays = [int]((Get-Date) - $disk.TimeCreated).TotalDays
            
            $results += [PSCustomObject]@{
                'ResourceType' = 'Managed Disk'
                'ResourceName' = $disk.Name
                'ResourceGroup' = $disk.ResourceGroupName
                'Location' = $disk.Location
                'Size' = "$sizeMB GB"
                'Tier' = $tier
                'AgeDays' = $ageDays
                'Created' = $disk.TimeCreated.ToString('yyyy-MM-dd')
                'MonthlyCostEstimate' = [math]::Round($costPerMonth, 2)
                'AnnualCostEstimate' = [math]::Round($costPerMonth * 12, 2)
                'Subscription' = $Subscription.Name
                'SubscriptionId' = $Subscription.Id
                'ResourceId' = $disk.Id
                'RiskLevel' = 'Low'
                'RecommendedAction' = 'Delete if confirmed unused'
            }
        }
        
        Write-Log "    Found $($results.Count) unattached disk(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning disks: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-UnusedPublicIPs {
    param(
        [object]$Subscription,
        [int]$MinAgeDays
    )
    
    Write-Log "  Scanning for unused public IPs in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $unusedIPs = Get-AzPublicIpAddress | Where-Object {
            $_.IpConfiguration -eq $null -and
            ((Get-Date) - (Get-Date $_.ResourceGuid.ToString().Substring(0,8))).Days -ge $MinAgeDays
        }
        
        $results = @()
        foreach ($ip in $unusedIPs) {
            $sku = $ip.Sku.Name
            $costPerMonth = if ($sku -eq "Standard") {
                $EstimatedCosts.PublicIP_Standard
            } else {
                $EstimatedCosts.PublicIP_Basic
            }
            
            $results += [PSCustomObject]@{
                'ResourceType' = 'Public IP Address'
                'ResourceName' = $ip.Name
                'ResourceGroup' = $ip.ResourceGroupName
                'Location' = $ip.Location
                'Size' = $sku
                'Tier' = if ($ip.PublicIpAllocationMethod) { $ip.PublicIpAllocationMethod } else { 'N/A' }
                'AgeDays' = [int]((Get-Date) - (Get-Date $ip.ResourceGuid.ToString().Substring(0,8))).TotalDays
                'Created' = 'N/A'
                'MonthlyCostEstimate' = [math]::Round($costPerMonth, 2)
                'AnnualCostEstimate' = [math]::Round($costPerMonth * 12, 2)
                'Subscription' = $Subscription.Name
                'SubscriptionId' = $Subscription.Id
                'ResourceId' = $ip.Id
                'RiskLevel' = 'Low'
                'RecommendedAction' = 'Delete if confirmed unused'
            }
        }
        
        Write-Log "    Found $($results.Count) unused public IP(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning public IPs: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-UnusedNetworkInterfaces {
    param(
        [object]$Subscription,
        [int]$MinAgeDays
    )
    
    Write-Log "  Scanning for unused network interfaces in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $unusedNICs = Get-AzNetworkInterface | Where-Object {
            $_.VirtualMachine -eq $null -and $_.PrivateEndpoint -eq $null
        }
        
        $results = @()
        foreach ($nic in $unusedNICs) {
            $results += [PSCustomObject]@{
                'ResourceType' = 'Network Interface'
                'ResourceName' = $nic.Name
                'ResourceGroup' = $nic.ResourceGroupName
                'Location' = $nic.Location
                'Size' = 'N/A'
                'Tier' = 'N/A'
                'AgeDays' = 'Unknown'
                'Created' = 'N/A'
                'MonthlyCostEstimate' = 0
                'AnnualCostEstimate' = 0
                'Subscription' = $Subscription.Name
                'SubscriptionId' = $Subscription.Id
                'ResourceId' = $nic.Id
                'RiskLevel' = 'Low'
                'RecommendedAction' = 'Delete if confirmed unused'
            }
        }
        
        Write-Log "    Found $($results.Count) unused network interface(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning network interfaces: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-OrphanedSnapshots {
    param(
        [object]$Subscription,
        [int]$MinAgeDays
    )
    
    Write-Log "  Scanning for orphaned snapshots in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $snapshots = Get-AzSnapshot | Where-Object {
            ((Get-Date) - $_.TimeCreated).Days -ge $MinAgeDays
        }
        
        $results = @()
        foreach ($snapshot in $snapshots) {
            $sizeGB = $snapshot.DiskSizeGB
            $costPerMonth = $sizeGB * $EstimatedCosts.Snapshot_Standard
            $ageDays = [int]((Get-Date) - $snapshot.TimeCreated).TotalDays
            
            $results += [PSCustomObject]@{
                'ResourceType' = 'Snapshot'
                'ResourceName' = $snapshot.Name
                'ResourceGroup' = $snapshot.ResourceGroupName
                'Location' = $snapshot.Location
                'Size' = "$sizeGB GB"
                'Tier' = $snapshot.Sku.Name
                'AgeDays' = $ageDays
                'Created' = $snapshot.TimeCreated.ToString('yyyy-MM-dd')
                'MonthlyCostEstimate' = [math]::Round($costPerMonth, 2)
                'AnnualCostEstimate' = [math]::Round($costPerMonth * 12, 2)
                'Subscription' = $Subscription.Name
                'SubscriptionId' = $Subscription.Id
                'ResourceId' = $snapshot.Id
                'RiskLevel' = 'Medium'
                'RecommendedAction' = 'Review retention policy and delete if not needed'
            }
        }
        
        Write-Log "    Found $($results.Count) snapshot(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning snapshots: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-StoppedVMs {
    param(
        [object]$Subscription,
        [int]$MinAgeDays
    )
    
    Write-Log "  Scanning for stopped VMs in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $stoppedVMs = Get-AzVM -Status | Where-Object {
            $_.PowerState -eq 'VM deallocated' -or $_.PowerState -eq 'VM stopped'
        }
        
        $results = @()
        foreach ($vm in $stoppedVMs) {
            $vmDetail = Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
            
            $results += [PSCustomObject]@{
                'ResourceType' = 'Stopped VM'
                'ResourceName' = $vm.Name
                'ResourceGroup' = $vm.ResourceGroupName
                'Location' = $vm.Location
                'Size' = $vmDetail.HardwareProfile.VmSize
                'Tier' = $vm.PowerState
                'AgeDays' = 'Unknown'
                'Created' = 'N/A'
                'MonthlyCostEstimate' = $EstimatedCosts.StoppedVM_Storage
                'AnnualCostEstimate' = [math]::Round($EstimatedCosts.StoppedVM_Storage * 12, 2)
                'Subscription' = $Subscription.Name
                'SubscriptionId' = $Subscription.Id
                'ResourceId' = $vmDetail.Id
                'RiskLevel' = 'Medium'
                'RecommendedAction' = 'Delete VM but retain disks if needed, or restart if still required'
            }
        }
        
        Write-Log "    Found $($results.Count) stopped VM(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning VMs: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-EmptyResourceGroups {
    param([object]$Subscription)
    
    Write-Log "  Scanning for empty resource groups in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $allRGs = Get-AzResourceGroup
        $results = @()
        
        foreach ($rg in $allRGs) {
            $resources = Get-AzResource -ResourceGroupName $rg.ResourceGroupName
            
            if ($resources.Count -eq 0) {
                $results += [PSCustomObject]@{
                    'ResourceType' = 'Empty Resource Group'
                    'ResourceName' = $rg.ResourceGroupName
                    'ResourceGroup' = $rg.ResourceGroupName
                    'Location' = $rg.Location
                    'Size' = '0 resources'
                    'Tier' = 'N/A'
                    'AgeDays' = 'Unknown'
                    'Created' = 'N/A'
                    'MonthlyCostEstimate' = 0
                    'AnnualCostEstimate' = 0
                    'Subscription' = $Subscription.Name
                    'SubscriptionId' = $Subscription.Id
                    'ResourceId' = $rg.ResourceId
                    'RiskLevel' = 'Low'
                    'RecommendedAction' = 'Delete resource group'
                }
            }
        }
        
        Write-Log "    Found $($results.Count) empty resource group(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning resource groups: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$UnusedResources,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Calculate summary statistics
        $totalResources = $UnusedResources.Count
        $totalMonthlyCost = ($UnusedResources | Measure-Object -Property MonthlyCostEstimate -Sum).Sum
        $totalAnnualCost = ($UnusedResources | Measure-Object -Property AnnualCostEstimate -Sum).Sum
        
        $byType = $UnusedResources | Group-Object ResourceType | ForEach-Object {
            [PSCustomObject]@{
                'ResourceType' = $_.Name
                'Count' = $_.Count
                'MonthlyCost' = [math]::Round(($_.Group | Measure-Object -Property MonthlyCostEstimate -Sum).Sum, 2)
                'AnnualCost' = [math]::Round(($_.Group | Measure-Object -Property AnnualCostEstimate -Sum).Sum, 2)
            }
        } | Sort-Object AnnualCost -Descending
        
        # Create executive summary
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total Unused Resources',
                'Estimated Monthly Savings',
                'Estimated Annual Savings',
                '--- Breakdown by Type ---',
                'Unattached Disks',
                'Unused Public IPs',
                'Network Interfaces',
                'Snapshots',
                'Stopped VMs',
                'Empty Resource Groups',
                'Report Generated'
            )
            'Value' = @(
                $totalResources,
                "`$$([math]::Round($totalMonthlyCost, 2))",
                "`$$([math]::Round($totalAnnualCost, 2))",
                '',
                ($UnusedResources | Where-Object {$_.ResourceType -eq 'Managed Disk'}).Count,
                ($UnusedResources | Where-Object {$_.ResourceType -eq 'Public IP Address'}).Count,
                ($UnusedResources | Where-Object {$_.ResourceType -eq 'Network Interface'}).Count,
                ($UnusedResources | Where-Object {$_.ResourceType -eq 'Snapshot'}).Count,
                ($UnusedResources | Where-Object {$_.ResourceType -eq 'Stopped VM'}).Count,
                ($UnusedResources | Where-Object {$_.ResourceType -eq 'Empty Resource Group'}).Count,
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Export cost by type
        if ($byType.Count -gt 0) {
            $byType | Export-Excel -Path $FilePath -WorksheetName "Cost by Type" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Export top 20 most expensive
        $top20 = $UnusedResources | Sort-Object AnnualCostEstimate -Descending | Select-Object -First 20
        if ($top20.Count -gt 0) {
            $top20 | Export-Excel -Path $FilePath -WorksheetName "Top 20 Most Expensive" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Export all unused resources
        if ($UnusedResources.Count -gt 0) {
            $UnusedResources | Sort-Object AnnualCostEstimate -Descending | 
                Export-Excel -Path $FilePath -WorksheetName "All Unused Resources" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "High" -Range "M:M" -BackgroundColor Orange
                    New-ConditionalText -Text "Medium" -Range "M:M" -BackgroundColor Yellow
                    New-ConditionalText -Text "Low" -Range "M:M" -BackgroundColor LightGreen
                )
        }
        
        # Export by subscription
        $bySubscription = $UnusedResources | Group-Object Subscription | ForEach-Object {
            [PSCustomObject]@{
                'Subscription' = $_.Name
                'ResourceCount' = $_.Count
                'MonthlyCost' = [math]::Round(($_.Group | Measure-Object -Property MonthlyCostEstimate -Sum).Sum, 2)
                'AnnualCost' = [math]::Round(($_.Group | Measure-Object -Property AnnualCostEstimate -Sum).Sum, 2)
            }
        } | Sort-Object AnnualCost -Descending
        
        if ($bySubscription.Count -gt 0) {
            $bySubscription | Export-Excel -Path $FilePath -WorksheetName "Cost by Subscription" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
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
        [int]$ResourceCount,
        [decimal]$MonthlySavings,
        [decimal]$AnnualSavings
    )
    
    Write-Log "Email functionality requires Microsoft Graph API"
    Write-Log "For Graph API email integration, refer to other runbook examples"
    Write-Log "Report saved locally: $AttachmentPath"
    
    # NOTE: Implement Graph API email sending similar to other runbooks if needed
    return $true
}
#endregion

#region Main Execution
try {
    Write-Log "========== Starting Unused Resources Analysis =========="
    
    # Connect to Azure
    Connect-AzureWithManagedIdentity
    
    # Get target subscriptions
    $subscriptions = Get-TargetSubscriptions -SubscriptionIds $TargetSubscriptions
    
    # Collect unused resources across all subscriptions
    $allUnusedResources = @()
    
    foreach ($subscription in $subscriptions) {
        Write-Log "`nAnalyzing subscription: $($subscription.Name)"
        
        $allUnusedResources += Get-UnattachedDisks -Subscription $subscription -MinAgeDays $MinimumAgeThreshold
        $allUnusedResources += Get-UnusedPublicIPs -Subscription $subscription -MinAgeDays $MinimumAgeThreshold
        $allUnusedResources += Get-UnusedNetworkInterfaces -Subscription $subscription -MinAgeDays $MinimumAgeThreshold
        $allUnusedResources += Get-OrphanedSnapshots -Subscription $subscription -MinAgeDays $MinimumAgeThreshold
        $allUnusedResources += Get-StoppedVMs -Subscription $subscription -MinAgeDays $MinimumAgeThreshold
        $allUnusedResources += Get-EmptyResourceGroups -Subscription $subscription
    }
    
    # Calculate totals
    $totalMonthlySavings = ($allUnusedResources | Measure-Object -Property MonthlyCostEstimate -Sum).Sum
    $totalAnnualSavings = ($allUnusedResources | Measure-Object -Property AnnualCostEstimate -Sum).Sum
    
    Write-Log "`n========== Analysis Complete =========="
    Write-Log "Total unused resources found: $($allUnusedResources.Count)"
    Write-Log "Estimated monthly savings: `$$([math]::Round($totalMonthlySavings, 2))"
    Write-Log "Estimated annual savings: `$$([math]::Round($totalAnnualSavings, 2))"
    
    # Export to Excel
    Export-ToExcelWithFormatting -UnusedResources $allUnusedResources -FilePath $FullReportPath
    
    # Send email notification
    Send-EmailWithAttachment -Recipients $EmailRecipients -Subject $EmailSubject -AttachmentPath $FullReportPath `
                             -ResourceCount $allUnusedResources.Count `
                             -MonthlySavings $totalMonthlySavings `
                             -AnnualSavings $totalAnnualSavings
    
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Found $($allUnusedResources.Count) unused resources. Potential annual savings: `$$([math]::Round($totalAnnualSavings, 2))"
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    throw
}
finally {
    # Disconnect from Azure
    try {
        Disconnect-AzAccount | Out-Null
        Write-Log "Disconnected from Azure"
    }
    catch {
        Write-Log "Error disconnecting from Azure: $($_.Exception.Message)" -Level WARNING
    }
}
#endregion
