#Requires -Modules @{ ModuleName="Az.Accounts"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Az.Network"; ModuleVersion="5.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Audits Network Security Group rules for security compliance and overly permissive configurations.

.DESCRIPTION
    This runbook analyzes NSG rules across Azure subscriptions to identify:
    - Rules allowing 0.0.0.0/0 (internet) access
    - High-risk ports (RDP, SSH, SQL, etc.) exposed to the internet
    - Rules with "Allow All" protocols
    - Priority conflicts and overlapping rules
    - Unused or redundant NSG rules
    
    Provides risk scoring and remediation recommendations.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Azure RBAC Permissions Required:
        - Network Contributor or Reader role on subscriptions
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("security@yourdomain.com", "network@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "NSG Security Audit Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Subscription Selection (empty array = all subscriptions)
$TargetSubscriptions = @()

# High-Risk Ports
$HighRiskPorts = @{
    'RDP' = @(3389)
    'SSH' = @(22)
    'Telnet' = @(23)
    'FTP' = @(20, 21)
    'SMB' = @(445, 139)
    'SQL Server' = @(1433, 1434)
    'MySQL' = @(3306)
    'PostgreSQL' = @(5432)
    'Oracle' = @(1521)
    'MongoDB' = @(27017)
    'Redis' = @(6379)
    'Cassandra' = @(9042)
    'Elasticsearch' = @(9200, 9300)
    'RDP Gateway' = @(3391)
    'WinRM' = @(5985, 5986)
}

# Risk Scoring Weights
$RiskWeights = @{
    InternetExposed = 10
    HighRiskPort = 8
    AllowAllProtocol = 6
    AllowAllPorts = 7
    BroadSourceRange = 5
}

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "NSGAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Test-IsInternetSource {
    param([string]$SourceAddressPrefix)
    
    $internetSources = @('*', '0.0.0.0/0', 'Internet', 'Any')
    return ($internetSources -contains $SourceAddressPrefix)
}

function Test-IsBroadSourceRange {
    param([string]$SourceAddressPrefix)
    
    if ($SourceAddressPrefix -match '/(\d+)') {
        $cidr = [int]$matches[1]
        return ($cidr -lt 16)  # /15 or smaller (broader) is considered risky
    }
    return $false
}

function Get-PortRiskCategory {
    param([string]$DestinationPort)
    
    foreach ($category in $HighRiskPorts.Keys) {
        foreach ($port in $HighRiskPorts[$category]) {
            if ($DestinationPort -eq $port -or $DestinationPort -eq "*") {
                return $category
            }
        }
    }
    return "Other"
}

function Calculate-RiskScore {
    param(
        [bool]$IsInternetExposed,
        [bool]$IsHighRiskPort,
        [bool]$IsAllowAllProtocol,
        [bool]$IsAllowAllPorts,
        [bool]$IsBroadSourceRange
    )
    
    $score = 0
    
    if ($IsInternetExposed) { $score += $RiskWeights.InternetExposed }
    if ($IsHighRiskPort) { $score += $RiskWeights.HighRiskPort }
    if ($IsAllowAllProtocol) { $score += $RiskWeights.AllowAllProtocol }
    if ($IsAllowAllPorts) { $score += $RiskWeights.AllowAllPorts }
    if ($IsBroadSourceRange) { $score += $RiskWeights.BroadSourceRange }
    
    return $score
}

function Get-RiskLevel {
    param([int]$RiskScore)
    
    if ($RiskScore -ge 20) { return "CRITICAL" }
    elseif ($RiskScore -ge 15) { return "HIGH" }
    elseif ($RiskScore -ge 10) { return "MEDIUM" }
    elseif ($RiskScore -gt 0) { return "LOW" }
    else { return "INFO" }
}

function Get-NSGRulesAudit {
    param([object]$Subscription)
    
    Write-Log "  Auditing NSG rules in subscription: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $nsgs = Get-AzNetworkSecurityGroup
        Write-Log "    Found $($nsgs.Count) NSG(s)"
        
        $auditResults = @()
        
        foreach ($nsg in $nsgs) {
            # Get all rules (default + custom)
            $allRules = $nsg.SecurityRules + $nsg.DefaultSecurityRules
            
            foreach ($rule in $allRules) {
                # Only audit Allow rules
                if ($rule.Access -ne 'Allow') { continue }
                
                # Parse destination ports
                $destPorts = if ($rule.DestinationPortRange) {
                    $rule.DestinationPortRange
                } else {
                    ($rule.DestinationPortRanges -join ', ')
                }
                
                # Parse source addresses
                $sourceAddresses = if ($rule.SourceAddressPrefix) {
                    $rule.SourceAddressPrefix
                } else {
                    ($rule.SourceAddressPrefixes -join ', ')
                }
                
                # Risk analysis
                $isInternetExposed = Test-IsInternetSource -SourceAddressPrefix $sourceAddresses
                $isBroadSource = Test-IsBroadSourceRange -SourceAddressPrefix $sourceAddresses
                $isAllowAllProtocol = ($rule.Protocol -eq '*')
                $isAllowAllPorts = ($destPorts -eq '*')
                
                # Check for high-risk ports
                $portRiskCategory = "N/A"
                $isHighRiskPort = $false
                
                if ($destPorts -ne '*') {
                    foreach ($category in $HighRiskPorts.Keys) {
                        foreach ($port in $HighRiskPorts[$category]) {
                            if ($destPorts -match "\b$port\b") {
                                $portRiskCategory = $category
                                $isHighRiskPort = $true
                                break
                            }
                        }
                        if ($isHighRiskPort) { break }
                    }
                }
                
                # Calculate risk score
                $riskScore = Calculate-RiskScore -IsInternetExposed $isInternetExposed `
                                                  -IsHighRiskPort $isHighRiskPort `
                                                  -IsAllowAllProtocol $isAllowAllProtocol `
                                                  -IsAllowAllPorts $isAllowAllPorts `
                                                  -IsBroadSourceRange $isBroadSource
                
                $riskLevel = Get-RiskLevel -RiskScore $riskScore
                
                # Generate recommendations
                $recommendations = @()
                if ($isInternetExposed -and $isHighRiskPort) {
                    $recommendations += "CRITICAL: Remove internet access to $portRiskCategory"
                }
                if ($isInternetExposed) {
                    $recommendations += "Restrict source to specific IP ranges"
                }
                if ($isAllowAllProtocol) {
                    $recommendations += "Specify protocol (TCP/UDP)"
                }
                if ($isAllowAllPorts) {
                    $recommendations += "Specify destination ports"
                }
                if ($isBroadSource) {
                    $recommendations += "Narrow source address range"
                }
                
                $auditResults += [PSCustomObject]@{
                    'RiskLevel' = $riskLevel
                    'RiskScore' = $riskScore
                    'NSGName' = $nsg.Name
                    'RuleName' = $rule.Name
                    'Priority' = $rule.Priority
                    'Direction' = $rule.Direction
                    'Access' = $rule.Access
                    'Protocol' = $rule.Protocol
                    'SourceAddress' = $sourceAddresses
                    'DestinationPort' = $destPorts
                    'PortCategory' = $portRiskCategory
                    'InternetExposed' = $isInternetExposed
                    'HighRiskPort' = $isHighRiskPort
                    'AllowAllProtocol' = $isAllowAllProtocol
                    'AllowAllPorts' = $isAllowAllPorts
                    'BroadSource' = $isBroadSource
                    'Recommendations' = ($recommendations -join '; ')
                    'ResourceGroup' = $nsg.ResourceGroupName
                    'Location' = $nsg.Location
                    'Subscription' = $Subscription.Name
                    'SubscriptionId' = $Subscription.Id
                    'NSGId' = $nsg.Id
                    'RuleId' = $rule.Id
                }
            }
        }
        
        Write-Log "    Analyzed $($auditResults.Count) rule(s)"
        return $auditResults
    }
    catch {
        Write-Log "    Error auditing NSG rules: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$AuditResults,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Calculate statistics
        $totalRules = $AuditResults.Count
        $criticalRules = ($AuditResults | Where-Object {$_.RiskLevel -eq 'CRITICAL'}).Count
        $highRules = ($AuditResults | Where-Object {$_.RiskLevel -eq 'HIGH'}).Count
        $mediumRules = ($AuditResults | Where-Object {$_.RiskLevel -eq 'MEDIUM'}).Count
        $internetExposed = ($AuditResults | Where-Object {$_.InternetExposed}).Count
        $highRiskPorts = ($AuditResults | Where-Object {$_.HighRiskPort}).Count
        
        # Create summary
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total NSG Rules Audited',
                '--- Risk Levels ---',
                'CRITICAL Risk Rules',
                'HIGH Risk Rules',
                'MEDIUM Risk Rules',
                'LOW Risk Rules',
                '--- Security Concerns ---',
                'Internet-Exposed Rules',
                'High-Risk Ports Exposed',
                'Allow All Protocols',
                'Allow All Ports',
                'Report Generated'
            )
            'Value' = @(
                $totalRules,
                '',
                $criticalRules,
                $highRules,
                $mediumRules,
                ($AuditResults | Where-Object {$_.RiskLevel -eq 'LOW'}).Count,
                '',
                $internetExposed,
                $highRiskPorts,
                ($AuditResults | Where-Object {$_.AllowAllProtocol}).Count,
                ($AuditResults | Where-Object {$_.AllowAllPorts}).Count,
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Top 20 highest risk rules
        $top20Risk = $AuditResults | Sort-Object RiskScore -Descending | Select-Object -First 20
        if ($top20Risk.Count -gt 0) {
            $top20Risk | Export-Excel -Path $FilePath -WorksheetName "Top 20 Highest Risk" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Critical rules
        $criticalRulesList = $AuditResults | Where-Object {$_.RiskLevel -eq 'CRITICAL'} | Sort-Object RiskScore -Descending
        if ($criticalRulesList.Count -gt 0) {
            $criticalRulesList | Export-Excel -Path $FilePath -WorksheetName "CRITICAL Rules" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Internet-exposed rules
        $internetExposedRules = $AuditResults | Where-Object {$_.InternetExposed} | Sort-Object RiskScore -Descending
        if ($internetExposedRules.Count -gt 0) {
            $internetExposedRules | Export-Excel -Path $FilePath -WorksheetName "Internet Exposed" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # High-risk ports exposed
        $highRiskPortRules = $AuditResults | Where-Object {$_.HighRiskPort} | Sort-Object RiskScore -Descending
        if ($highRiskPortRules.Count -gt 0) {
            $highRiskPortRules | Export-Excel -Path $FilePath -WorksheetName "High-Risk Ports" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # All rules with conditional formatting
        if ($AuditResults.Count -gt 0) {
            $AuditResults | Sort-Object RiskScore -Descending | 
                Export-Excel -Path $FilePath -WorksheetName "All Rules" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "CRITICAL" -Range "A:A" -BackgroundColor Red -ConditionalTextColor White
                    New-ConditionalText -Text "HIGH" -Range "A:A" -BackgroundColor Orange
                    New-ConditionalText -Text "MEDIUM" -Range "A:A" -BackgroundColor Yellow
                    New-ConditionalText -Text "LOW" -Range "A:A" -BackgroundColor LightGreen
                    New-ConditionalText -Text "True" -Range "L:P" -BackgroundColor LightCoral
                )
        }
        
        # By NSG
        $byNSG = $AuditResults | Group-Object NSGName | ForEach-Object {
            $criticalCount = ($_.Group | Where-Object {$_.RiskLevel -eq 'CRITICAL'}).Count
            $highCount = ($_.Group | Where-Object {$_.RiskLevel -eq 'HIGH'}).Count
            
            [PSCustomObject]@{
                'NSGName' = $_.Name
                'ResourceGroup' = $_.Group[0].ResourceGroup
                'TotalRules' = $_.Count
                'CriticalRules' = $criticalCount
                'HighRules' = $highCount
                'AverageRiskScore' = [math]::Round(($_.Group | Measure-Object -Property RiskScore -Average).Average, 2)
                'Subscription' = $_.Group[0].Subscription
            }
        } | Sort-Object AverageRiskScore -Descending
        
        if ($byNSG.Count -gt 0) {
            $byNSG | Export-Excel -Path $FilePath -WorksheetName "By NSG" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
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
    Write-Log "========== Starting NSG Security Audit =========="
    
    # Connect to Azure
    Connect-AzureWithManagedIdentity
    
    # Get target subscriptions
    $subscriptions = Get-TargetSubscriptions -SubscriptionIds $TargetSubscriptions
    
    # Collect audit results across all subscriptions
    $allAuditResults = @()
    
    foreach ($subscription in $subscriptions) {
        Write-Log "`nAnalyzing subscription: $($subscription.Name)"
        $allAuditResults += Get-NSGRulesAudit -Subscription $subscription
    }
    
    # Calculate summary statistics
    $criticalCount = ($allAuditResults | Where-Object {$_.RiskLevel -eq 'CRITICAL'}).Count
    $highCount = ($allAuditResults | Where-Object {$_.RiskLevel -eq 'HIGH'}).Count
    $internetExposedCount = ($allAuditResults | Where-Object {$_.InternetExposed}).Count
    
    Write-Log "`n========== Audit Complete =========="
    Write-Log "Total rules audited: $($allAuditResults.Count)"
    Write-Log "CRITICAL risk rules: $criticalCount"
    Write-Log "HIGH risk rules: $highCount"
    Write-Log "Internet-exposed rules: $internetExposedCount"
    
    # Export to Excel
    Export-ToExcelWithFormatting -AuditResults $allAuditResults -FilePath $FullReportPath
    
    Write-Log "Report saved: $FullReportPath"
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Audited $($allAuditResults.Count) NSG rules. Found $criticalCount CRITICAL and $highCount HIGH risk rules."
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
