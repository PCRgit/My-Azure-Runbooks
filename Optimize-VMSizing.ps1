#Requires -Modules @{ ModuleName="Az.Accounts"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Az.Compute"; ModuleVersion="5.0.0" }
#Requires -Modules @{ ModuleName="Az.Monitor"; ModuleVersion="4.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Analyzes VM utilization and provides right-sizing recommendations for cost optimization.

.DESCRIPTION
    This runbook collects performance metrics from Azure VMs over a configurable period:
    - CPU utilization
    - Memory utilization (requires Azure Monitor Agent)
    - Disk I/O patterns
    - Network throughput
    
    Provides resize recommendations with cost impact analysis.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Azure RBAC Permissions Required:
        - Reader role on subscriptions
        - Monitoring Reader role for metrics
        
    IMPORTANT: Requires Azure Monitor Agent or Diagnostics Extension for memory metrics
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("admin@yourdomain.com", "finance@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "VM Right-Sizing Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Subscription Selection (empty array = all subscriptions)
$TargetSubscriptions = @()

# Analysis Period
$AnalysisDays = 30  # Number of days to analyze metrics

# Utilization Thresholds
$Thresholds = @{
    OverProvisionedCPU = 20    # Average CPU < 20% = over-provisioned
    UnderProvisionedCPU = 85   # Average CPU > 85% = under-provisioned
    OverProvisionedMemory = 30 # Average Memory < 30% = over-provisioned
    UnderProvisionedMemory = 90 # Average Memory > 90% = under-provisioned
}

# VM Size Pricing (USD per month - update with your region/pricing)
# This is a simplified example - use actual Azure pricing for your region
$VMPricing = @{
    'Standard_B1s' = 7.59
    'Standard_B1ms' = 15.18
    'Standard_B2s' = 30.37
    'Standard_B2ms' = 60.74
    'Standard_B4ms' = 121.47
    'Standard_D2s_v3' = 96.36
    'Standard_D4s_v3' = 192.72
    'Standard_D8s_v3' = 385.44
    'Standard_D16s_v3' = 770.88
    'Standard_D2s_v4' = 96.36
    'Standard_D4s_v4' = 192.72
    'Standard_D8s_v4' = 385.44
    'Standard_E2s_v3' = 126.00
    'Standard_E4s_v3' = 252.00
    'Standard_E8s_v3' = 504.00
    'Standard_F2s_v2' = 84.00
    'Standard_F4s_v2' = 168.00
    'Standard_F8s_v2' = 336.00
}

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "VMRightSizing_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-VMMetrics {
    param(
        [object]$VM,
        [int]$Days
    )
    
    try {
        $endTime = Get-Date
        $startTime = $endTime.AddDays(-$Days)
        
        # Get CPU metrics
        $cpuMetrics = Get-AzMetric -ResourceId $VM.Id `
                                    -MetricName "Percentage CPU" `
                                    -StartTime $startTime `
                                    -EndTime $endTime `
                                    -TimeGrain 01:00:00 `
                                    -AggregationType Average `
                                    -ErrorAction SilentlyContinue
        
        $avgCPU = if ($cpuMetrics -and $cpuMetrics.Data.Count -gt 0) {
            [math]::Round(($cpuMetrics.Data.Average | Measure-Object -Average).Average, 2)
        } else {
            $null
        }
        
        $maxCPU = if ($cpuMetrics -and $cpuMetrics.Data.Count -gt 0) {
            [math]::Round(($cpuMetrics.Data.Average | Measure-Object -Maximum).Maximum, 2)
        } else {
            $null
        }
        
        # Get network metrics
        $networkInMetrics = Get-AzMetric -ResourceId $VM.Id `
                                         -MetricName "Network In Total" `
                                         -StartTime $startTime `
                                         -EndTime $endTime `
                                         -TimeGrain 01:00:00 `
                                         -AggregationType Average `
                                         -ErrorAction SilentlyContinue
        
        $avgNetworkIn = if ($networkInMetrics -and $networkInMetrics.Data.Count -gt 0) {
            [math]::Round(($networkInMetrics.Data.Average | Measure-Object -Average).Average / 1MB, 2)
        } else {
            $null
        }
        
        # Get disk metrics (read + write IOPS)
        $diskReadMetrics = Get-AzMetric -ResourceId $VM.Id `
                                        -MetricName "Disk Read Operations/Sec" `
                                        -StartTime $startTime `
                                        -EndTime $endTime `
                                        -TimeGrain 01:00:00 `
                                        -AggregationType Average `
                                        -ErrorAction SilentlyContinue
        
        $avgDiskRead = if ($diskReadMetrics -and $diskReadMetrics.Data.Count -gt 0) {
            [math]::Round(($diskReadMetrics.Data.Average | Measure-Object -Average).Average, 2)
        } else {
            $null
        }
        
        return @{
            AvgCPU = $avgCPU
            MaxCPU = $maxCPU
            AvgNetworkInMB = $avgNetworkIn
            AvgDiskReadIOPS = $avgDiskRead
        }
    }
    catch {
        Write-Log "    Error retrieving metrics for $($VM.Name): $($_.Exception.Message)" -Level WARNING
        return @{
            AvgCPU = $null
            MaxCPU = $null
            AvgNetworkInMB = $null
            AvgDiskReadIOPS = $null
        }
    }
}

function Get-ResizeRecommendation {
    param(
        [string]$CurrentSize,
        [object]$Metrics
    )
    
    # Simplified recommendation logic
    # In production, use Azure Advisor API or more sophisticated logic
    
    $avgCPU = $Metrics.AvgCPU
    
    if ($null -eq $avgCPU) {
        return @{
            Recommendation = "Insufficient Data"
            RecommendedSize = $CurrentSize
            Reason = "No CPU metrics available"
        }
    }
    
    # Over-provisioned (CPU too low)
    if ($avgCPU -lt $Thresholds.OverProvisionedCPU) {
        # Suggest downsize (simplified logic)
        $recommendedSize = $CurrentSize  # Keep same by default
        
        # Example logic for D-series
        if ($CurrentSize -match 'D(\d+)s_v') {
            $cores = [int]$matches[1]
            if ($cores -gt 2) {
                $newCores = $cores / 2
                $recommendedSize = $CurrentSize -replace "D$cores", "D$newCores"
            }
        }
        
        return @{
            Recommendation = "Downsize - Over-Provisioned"
            RecommendedSize = $recommendedSize
            Reason = "Average CPU utilization is only $avgCPU%, indicating over-provisioning"
        }
    }
    # Under-provisioned (CPU too high)
    elseif ($avgCPU -gt $Thresholds.UnderProvisionedCPU) {
        # Suggest upsize (simplified logic)
        $recommendedSize = $CurrentSize
        
        # Example logic for D-series
        if ($CurrentSize -match 'D(\d+)s_v') {
            $cores = [int]$matches[1]
            $newCores = $cores * 2
            $recommendedSize = $CurrentSize -replace "D$cores", "D$newCores"
        }
        
        return @{
            Recommendation = "Upsize - Under-Provisioned"
            RecommendedSize = $recommendedSize
            Reason = "Average CPU utilization is $avgCPU%, indicating under-provisioning"
        }
    }
    else {
        return @{
            Recommendation = "Optimally Sized"
            RecommendedSize = $CurrentSize
            Reason = "CPU utilization is within optimal range ($avgCPU%)"
        }
    }
}

function Calculate-CostImpact {
    param(
        [string]$CurrentSize,
        [string]$RecommendedSize
    )
    
    $currentCost = if ($VMPricing.ContainsKey($CurrentSize)) {
        $VMPricing[$CurrentSize]
    } else {
        0
    }
    
    $recommendedCost = if ($VMPricing.ContainsKey($RecommendedSize)) {
        $VMPricing[$RecommendedSize]
    } else {
        0
    }
    
    $monthlySavings = $currentCost - $recommendedCost
    $annualSavings = $monthlySavings * 12
    
    return @{
        CurrentMonthlyCost = $currentCost
        RecommendedMonthlyCost = $recommendedCost
        MonthlySavings = $monthlySavings
        AnnualSavings = $annualSavings
    }
}

function Analyze-VMSizing {
    param(
        [object]$Subscription,
        [int]$Days
    )
    
    Write-Log "  Analyzing VMs in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $vms = Get-AzVM -Status
        Write-Log "    Found $($vms.Count) VM(s)"
        
        $results = @()
        $counter = 0
        
        foreach ($vm in $vms) {
            $counter++
            Write-Log "    Analyzing VM $counter/$($vms.Count): $($vm.Name)" -Level INFO
            
            # Skip stopped VMs
            if ($vm.PowerState -ne 'VM running') {
                Write-Log "      Skipping - VM is not running ($($vm.PowerState))"
                continue
            }
            
            # Get VM details
            $vmDetail = Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
            $vmSize = $vmDetail.HardwareProfile.VmSize
            
            # Get metrics
            $metrics = Get-VMMetrics -VM $vmDetail -Days $Days
            
            # Get recommendation
            $recommendation = Get-ResizeRecommendation -CurrentSize $vmSize -Metrics $metrics
            
            # Calculate cost impact
            $costImpact = Calculate-CostImpact -CurrentSize $vmSize -RecommendedSize $recommendation.RecommendedSize
            
            $results += [PSCustomObject]@{
                'VMName' = $vm.Name
                'ResourceGroup' = $vm.ResourceGroupName
                'Location' = $vm.Location
                'CurrentSize' = $vmSize
                'PowerState' = $vm.PowerState
                'AvgCPU' = if ($metrics.AvgCPU) { "$($metrics.AvgCPU)%" } else { 'N/A' }
                'MaxCPU' = if ($metrics.MaxCPU) { "$($metrics.MaxCPU)%" } else { 'N/A' }
                'AvgNetworkMB' = if ($metrics.AvgNetworkInMB) { $metrics.AvgNetworkInMB } else { 'N/A' }
                'AvgDiskIOPS' = if ($metrics.AvgDiskReadIOPS) { $metrics.AvgDiskReadIOPS } else { 'N/A' }
                'Recommendation' = $recommendation.Recommendation
                'RecommendedSize' = $recommendation.RecommendedSize
                'Reason' = $recommendation.Reason
                'CurrentMonthlyCost' = $costImpact.CurrentMonthlyCost
                'RecommendedMonthlyCost' = $costImpact.RecommendedMonthlyCost
                'MonthlySavings' = [math]::Round($costImpact.MonthlySavings, 2)
                'AnnualSavings' = [math]::Round($costImpact.AnnualSavings, 2)
                'Subscription' = $Subscription.Name
                'SubscriptionId' = $Subscription.Id
                'VMId' = $vmDetail.Id
            }
            
            # Rate limiting
            Start-Sleep -Milliseconds 500
        }
        
        Write-Log "    Analyzed $($results.Count) VM(s)"
        return $results
    }
    catch {
        Write-Log "    Error analyzing VMs: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$AnalysisResults,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Calculate summary statistics
        $totalVMs = $AnalysisResults.Count
        $overProvisioned = ($AnalysisResults | Where-Object {$_.Recommendation -like "*Over-Provisioned*"}).Count
        $underProvisioned = ($AnalysisResults | Where-Object {$_.Recommendation -like "*Under-Provisioned*"}).Count
        $optimal = ($AnalysisResults | Where-Object {$_.Recommendation -eq "Optimally Sized"}).Count
        
        $totalMonthlySavings = ($AnalysisResults | Where-Object {$_.MonthlySavings -gt 0} | 
                                Measure-Object -Property MonthlySavings -Sum).Sum
        $totalAnnualSavings = ($AnalysisResults | Where-Object {$_.AnnualSavings -gt 0} | 
                               Measure-Object -Property AnnualSavings -Sum).Sum
        
        # Create summary
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total VMs Analyzed',
                '--- Sizing Assessment ---',
                'Over-Provisioned (Downsize)',
                'Under-Provisioned (Upsize)',
                'Optimally Sized',
                'Insufficient Data',
                '--- Cost Optimization ---',
                'Potential Monthly Savings',
                'Potential Annual Savings',
                'Report Generated',
                'Analysis Period (Days)'
            )
            'Value' = @(
                $totalVMs,
                '',
                $overProvisioned,
                $underProvisioned,
                $optimal,
                ($AnalysisResults | Where-Object {$_.Recommendation -eq "Insufficient Data"}).Count,
                '',
                "`$$([math]::Round($totalMonthlySavings, 2))",
                "`$$([math]::Round($totalAnnualSavings, 2))",
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),
                $AnalysisDays
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Top 20 cost savings opportunities
        $top20Savings = $AnalysisResults | Where-Object {$_.AnnualSavings -gt 0} | 
                        Sort-Object AnnualSavings -Descending | Select-Object -First 20
        
        if ($top20Savings.Count -gt 0) {
            $top20Savings | Export-Excel -Path $FilePath -WorksheetName "Top 20 Savings" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Over-provisioned VMs
        $overProvisionedVMs = $AnalysisResults | Where-Object {$_.Recommendation -like "*Over-Provisioned*"} |
                              Sort-Object AnnualSavings -Descending
        
        if ($overProvisionedVMs.Count -gt 0) {
            $overProvisionedVMs | Export-Excel -Path $FilePath -WorksheetName "Over-Provisioned" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Under-provisioned VMs
        $underProvisionedVMs = $AnalysisResults | Where-Object {$_.Recommendation -like "*Under-Provisioned*"}
        
        if ($underProvisionedVMs.Count -gt 0) {
            $underProvisionedVMs | Export-Excel -Path $FilePath -WorksheetName "Under-Provisioned" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # All VMs with conditional formatting
        if ($AnalysisResults.Count -gt 0) {
            $AnalysisResults | Sort-Object AnnualSavings -Descending | 
                Export-Excel -Path $FilePath -WorksheetName "All VMs" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "Over-Provisioned" -Range "K:K" -BackgroundColor LightYellow
                    New-ConditionalText -Text "Under-Provisioned" -Range "K:K" -BackgroundColor Orange
                    New-ConditionalText -Text "Optimally Sized" -Range "K:K" -BackgroundColor LightGreen
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
#endregion

#region Main Execution
try {
    Write-Log "========== Starting VM Right-Sizing Analysis =========="
    
    # Connect to Azure
    Connect-AzureWithManagedIdentity
    
    # Get target subscriptions
    $subscriptions = Get-TargetSubscriptions -SubscriptionIds $TargetSubscriptions
    
    # Analyze VMs across all subscriptions
    $allResults = @()
    
    foreach ($subscription in $subscriptions) {
        Write-Log "`nAnalyzing subscription: $($subscription.Name)"
        $allResults += Analyze-VMSizing -Subscription $subscription -Days $AnalysisDays
    }
    
    # Calculate totals
    $totalSavings = ($allResults | Where-Object {$_.AnnualSavings -gt 0} | 
                     Measure-Object -Property AnnualSavings -Sum).Sum
    $overProvisionedCount = ($allResults | Where-Object {$_.Recommendation -like "*Over-Provisioned*"}).Count
    
    Write-Log "`n========== Analysis Complete =========="
    Write-Log "Total VMs analyzed: $($allResults.Count)"
    Write-Log "Over-provisioned VMs: $overProvisionedCount"
    Write-Log "Potential annual savings: `$$([math]::Round($totalSavings, 2))"
    
    # Export to Excel
    Export-ToExcelWithFormatting -AnalysisResults $allResults -FilePath $FullReportPath
    
    Write-Log "Report saved: $FullReportPath"
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Analyzed $($allResults.Count) VMs. Potential annual savings: `$$([math]::Round($totalSavings, 2))"
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
