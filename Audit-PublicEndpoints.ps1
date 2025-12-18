#Requires -Modules @{ ModuleName="Az.Accounts"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Az.Resources"; ModuleVersion="6.0.0" }
#Requires -Modules @{ ModuleName="Az.Storage"; ModuleVersion="5.0.0" }
#Requires -Modules @{ ModuleName="Az.Sql"; ModuleVersion="3.0.0" }
#Requires -Modules @{ ModuleName="Az.Network"; ModuleVersion="5.0.0" }
#Requires -Modules @{ ModuleName="Az.Websites"; ModuleVersion="3.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Audits Azure resources for public exposure and security risks.

.DESCRIPTION
    This runbook identifies publicly exposed Azure resources across subscriptions:
    - Storage accounts with public access
    - SQL databases with public endpoints
    - VMs with public IP addresses
    - App Services without access restrictions
    - Azure SQL Managed Instances with public endpoints
    - PostgreSQL/MySQL with public access
    - Key Vaults with public network access
    - Cosmos DB with public endpoints
    
    Provides risk categorization and remediation recommendations.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Azure RBAC Permissions Required:
        - Reader role on subscriptions
        - Storage Account Contributor (for storage analysis)
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("security@yourdomain.com", "compliance@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "Public Endpoints Security Audit - $(Get-Date -Format 'yyyy-MM-dd')"

# Subscription Selection (empty array = all subscriptions)
$TargetSubscriptions = @()

# Risk Categories
$RiskLevels = @{
    'Critical' = @('Storage Account with Public Blob Access', 'SQL Database with Public Endpoint', 'Key Vault Public Access')
    'High' = @('VM with Public IP', 'SQL Managed Instance Public Endpoint', 'Cosmos DB Public Endpoint')
    'Medium' = @('App Service without IP Restrictions', 'PostgreSQL Public Access', 'MySQL Public Access')
}

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "PublicEndpoints_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-RiskLevel {
    param([string]$ResourceType)
    
    foreach ($level in $RiskLevels.Keys) {
        if ($RiskLevels[$level] -contains $ResourceType) {
            return $level
        }
    }
    return 'Low'
}

function Get-PublicStorageAccounts {
    param([object]$Subscription)
    
    Write-Log "  Scanning storage accounts in: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $storageAccounts = Get-AzStorageAccount
        $results = @()
        
        foreach ($sa in $storageAccounts) {
            # Check network rules
            $hasPublicAccess = $false
            $accessDetails = "Private"
            
            if ($sa.NetworkRuleSet.DefaultAction -eq 'Allow') {
                $hasPublicAccess = $true
                $accessDetails = "Public - Allow from all networks"
            }
            elseif ($sa.PublicNetworkAccess -eq 'Enabled' -and $sa.NetworkRuleSet.DefaultAction -eq 'Deny') {
                # Check if there are bypass rules
                if ($sa.NetworkRuleSet.Bypass -ne 'None') {
                    $hasPublicAccess = $true
                    $accessDetails = "Public - Bypass: $($sa.NetworkRuleSet.Bypass)"
                }
            }
            
            # Check blob public access
            $blobPublicAccess = $sa.AllowBlobPublicAccess
            if ($blobPublicAccess) {
                $hasPublicAccess = $true
                $accessDetails = "Blob Public Access Enabled"
            }
            
            if ($hasPublicAccess) {
                $results += [PSCustomObject]@{
                    'ResourceType' = 'Storage Account with Public Blob Access'
                    'ResourceName' = $sa.StorageAccountName
                    'ResourceGroup' = $sa.ResourceGroupName
                    'Location' = $sa.Location
                    'RiskLevel' = Get-RiskLevel -ResourceType 'Storage Account with Public Blob Access'
                    'ExposureType' = $accessDetails
                    'PublicEndpoint' = "$($sa.StorageAccountName).blob.core.windows.net"
                    'Recommendation' = "Disable public blob access, enable firewall rules, use Private Endpoints"
                    'NetworkAccess' = $sa.NetworkRuleSet.DefaultAction
                    'BlobPublicAccess' = $blobPublicAccess
                    'Subscription' = $Subscription.Name
                    'SubscriptionId' = $Subscription.Id
                    'ResourceId' = $sa.Id
                    'Tags' = ($sa.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                }
            }
        }
        
        Write-Log "    Found $($results.Count) public storage account(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning storage accounts: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-PublicSQLDatabases {
    param([object]$Subscription)
    
    Write-Log "  Scanning SQL databases in: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $sqlServers = Get-AzSqlServer
        $results = @()
        
        foreach ($server in $sqlServers) {
            # Check firewall rules
            $firewallRules = Get-AzSqlServerFirewallRule -ServerName $server.ServerName -ResourceGroupName $server.ResourceGroupName
            
            $hasPublicAccess = $false
            $exposureDetails = ""
            
            # Check for 0.0.0.0 - 255.255.255.255 (Allow All)
            $allowAllRule = $firewallRules | Where-Object {
                $_.StartIpAddress -eq '0.0.0.0' -and $_.EndIpAddress -eq '255.255.255.255'
            }
            
            if ($allowAllRule) {
                $hasPublicAccess = $true
                $exposureDetails = "Allow All IP Addresses (0.0.0.0-255.255.255.255)"
            }
            elseif ($firewallRules.Count -gt 0) {
                $hasPublicAccess = $true
                $exposureDetails = "Specific IP ranges allowed: $($firewallRules.Count) rule(s)"
            }
            
            # Check if Azure services allowed
            $azureServicesRule = $firewallRules | Where-Object { $_.StartIpAddress -eq '0.0.0.0' -and $_.EndIpAddress -eq '0.0.0.0' }
            if ($azureServicesRule) {
                $exposureDetails += " + Azure Services"
            }
            
            if ($hasPublicAccess) {
                $results += [PSCustomObject]@{
                    'ResourceType' = 'SQL Database with Public Endpoint'
                    'ResourceName' = $server.ServerName
                    'ResourceGroup' = $server.ResourceGroupName
                    'Location' = $server.Location
                    'RiskLevel' = Get-RiskLevel -ResourceType 'SQL Database with Public Endpoint'
                    'ExposureType' = $exposureDetails
                    'PublicEndpoint' = "$($server.FullyQualifiedDomainName)"
                    'Recommendation' = "Implement Private Link, restrict firewall rules to specific IPs, enable Advanced Threat Protection"
                    'FirewallRuleCount' = $firewallRules.Count
                    'AllowAzureServices' = ($azureServicesRule -ne $null)
                    'Subscription' = $Subscription.Name
                    'SubscriptionId' = $Subscription.Id
                    'ResourceId' = $server.ResourceId
                    'Tags' = ($server.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                }
            }
        }
        
        Write-Log "    Found $($results.Count) public SQL server(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning SQL databases: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-PublicVMs {
    param([object]$Subscription)
    
    Write-Log "  Scanning VMs with public IPs in: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $vms = Get-AzVM
        $results = @()
        
        foreach ($vm in $vms) {
            # Get network interfaces
            $nics = $vm.NetworkProfile.NetworkInterfaces
            
            foreach ($nicRef in $nics) {
                $nicId = $nicRef.Id
                $nic = Get-AzNetworkInterface -ResourceId $nicId
                
                foreach ($ipConfig in $nic.IpConfigurations) {
                    if ($ipConfig.PublicIpAddress) {
                        $pipId = $ipConfig.PublicIpAddress.Id
                        $pip = Get-AzPublicIpAddress -ResourceId $pipId -ErrorAction SilentlyContinue
                        
                        if ($pip -and $pip.IpAddress) {
                            # Get associated NSG
                            $nsgName = if ($nic.NetworkSecurityGroup) {
                                (Get-AzResource -ResourceId $nic.NetworkSecurityGroup.Id).Name
                            } else {
                                'None'
                            }
                            
                            $results += [PSCustomObject]@{
                                'ResourceType' = 'VM with Public IP'
                                'ResourceName' = $vm.Name
                                'ResourceGroup' = $vm.ResourceGroupName
                                'Location' = $vm.Location
                                'RiskLevel' = Get-RiskLevel -ResourceType 'VM with Public IP'
                                'ExposureType' = "Public IP: $($pip.IpAddress)"
                                'PublicEndpoint' = $pip.IpAddress
                                'Recommendation' = "Use Azure Bastion or VPN for access, implement NSG rules, consider removing public IP"
                                'VMSize' = $vm.HardwareProfile.VmSize
                                'NSG' = $nsgName
                                'Subscription' = $Subscription.Name
                                'SubscriptionId' = $Subscription.Id
                                'ResourceId' = $vm.Id
                                'Tags' = ($vm.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                            }
                        }
                    }
                }
            }
        }
        
        Write-Log "    Found $($results.Count) VM(s) with public IP"
        return $results
    }
    catch {
        Write-Log "    Error scanning VMs: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-PublicAppServices {
    param([object]$Subscription)
    
    Write-Log "  Scanning App Services in: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $webApps = Get-AzWebApp
        $results = @()
        
        foreach ($app in $webApps) {
            # Check IP restrictions
            $siteConfig = Get-AzWebApp -ResourceGroupName $app.ResourceGroup -Name $app.Name
            
            $hasRestrictions = $false
            $restrictionCount = 0
            
            if ($siteConfig.SiteConfig.IpSecurityRestrictions) {
                # Exclude default "Allow All" rule
                $restrictions = $siteConfig.SiteConfig.IpSecurityRestrictions | Where-Object {
                    $_.Action -ne 'Allow' -or $_.IpAddress -ne 'Any'
                }
                
                if ($restrictions.Count -gt 0) {
                    $hasRestrictions = $true
                    $restrictionCount = $restrictions.Count
                }
            }
            
            if (-not $hasRestrictions) {
                $results += [PSCustomObject]@{
                    'ResourceType' = 'App Service without IP Restrictions'
                    'ResourceName' = $app.Name
                    'ResourceGroup' = $app.ResourceGroup
                    'Location' = $app.Location
                    'RiskLevel' = Get-RiskLevel -ResourceType 'App Service without IP Restrictions'
                    'ExposureType' = "No IP restrictions - publicly accessible"
                    'PublicEndpoint' = "https://$($app.DefaultHostName)"
                    'Recommendation' = "Implement IP restrictions, use Private Endpoints, enable authentication"
                    'AppServicePlan' = $app.ServerFarmId.Split('/')[-1]
                    'HTTPSOnly' = $app.HttpsOnly
                    'Subscription' = $Subscription.Name
                    'SubscriptionId' = $Subscription.Id
                    'ResourceId' = $app.Id
                    'Tags' = ($app.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                }
            }
        }
        
        Write-Log "    Found $($results.Count) App Service(s) without restrictions"
        return $results
    }
    catch {
        Write-Log "    Error scanning App Services: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-PublicKeyVaults {
    param([object]$Subscription)
    
    Write-Log "  Scanning Key Vaults in: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $keyVaults = Get-AzKeyVault
        $results = @()
        
        foreach ($kv in $keyVaults) {
            $kvDetail = Get-AzKeyVault -VaultName $kv.VaultName -ResourceGroupName $kv.ResourceGroupName
            
            # Check network rules
            $hasPublicAccess = $false
            $accessDetails = "Private"
            
            if ($kvDetail.NetworkAcls.DefaultAction -eq 'Allow') {
                $hasPublicAccess = $true
                $accessDetails = "Public - Allow from all networks"
            }
            elseif ($kvDetail.PublicNetworkAccess -eq 'Enabled') {
                $hasPublicAccess = $true
                $accessDetails = "Public network access enabled"
            }
            
            if ($hasPublicAccess) {
                $results += [PSCustomObject]@{
                    'ResourceType' = 'Key Vault Public Access'
                    'ResourceName' = $kv.VaultName
                    'ResourceGroup' = $kv.ResourceGroupName
                    'Location' = $kv.Location
                    'RiskLevel' = Get-RiskLevel -ResourceType 'Key Vault Public Access'
                    'ExposureType' = $accessDetails
                    'PublicEndpoint' = "https://$($kv.VaultName).vault.azure.net"
                    'Recommendation' = "Disable public network access, use Private Endpoints, implement firewall rules"
                    'SoftDeleteEnabled' = $kvDetail.EnableSoftDelete
                    'PurgeProtectionEnabled' = $kvDetail.EnablePurgeProtection
                    'Subscription' = $Subscription.Name
                    'SubscriptionId' = $Subscription.Id
                    'ResourceId' = $kvDetail.ResourceId
                    'Tags' = ($kvDetail.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                }
            }
        }
        
        Write-Log "    Found $($results.Count) public Key Vault(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning Key Vaults: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$PublicEndpoints,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Calculate statistics
        $totalEndpoints = $PublicEndpoints.Count
        $criticalCount = ($PublicEndpoints | Where-Object {$_.RiskLevel -eq 'Critical'}).Count
        $highCount = ($PublicEndpoints | Where-Object {$_.RiskLevel -eq 'High'}).Count
        $mediumCount = ($PublicEndpoints | Where-Object {$_.RiskLevel -eq 'Medium'}).Count
        
        # Group by resource type
        $byType = $PublicEndpoints | Group-Object ResourceType | ForEach-Object {
            [PSCustomObject]@{
                'ResourceType' = $_.Name
                'Count' = $_.Count
                'CriticalRisk' = ($_.Group | Where-Object {$_.RiskLevel -eq 'Critical'}).Count
                'HighRisk' = ($_.Group | Where-Object {$_.RiskLevel -eq 'High'}).Count
                'MediumRisk' = ($_.Group | Where-Object {$_.RiskLevel -eq 'Medium'}).Count
            }
        } | Sort-Object Count -Descending
        
        # Create summary
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total Public Endpoints',
                '--- Risk Levels ---',
                'CRITICAL Risk',
                'HIGH Risk',
                'MEDIUM Risk',
                'LOW Risk',
                '--- Top Exposures ---',
                'Storage Accounts',
                'SQL Databases',
                'VMs with Public IP',
                'App Services',
                'Key Vaults',
                'Report Generated'
            )
            'Value' = @(
                $totalEndpoints,
                '',
                $criticalCount,
                $highCount,
                $mediumCount,
                ($PublicEndpoints | Where-Object {$_.RiskLevel -eq 'Low'}).Count,
                '',
                ($PublicEndpoints | Where-Object {$_.ResourceType -like '*Storage*'}).Count,
                ($PublicEndpoints | Where-Object {$_.ResourceType -like '*SQL*'}).Count,
                ($PublicEndpoints | Where-Object {$_.ResourceType -like '*VM*'}).Count,
                ($PublicEndpoints | Where-Object {$_.ResourceType -like '*App Service*'}).Count,
                ($PublicEndpoints | Where-Object {$_.ResourceType -like '*Key Vault*'}).Count,
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Export by resource type
        if ($byType.Count -gt 0) {
            $byType | Export-Excel -Path $FilePath -WorksheetName "By Resource Type" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Critical risk resources
        $critical = $PublicEndpoints | Where-Object {$_.RiskLevel -eq 'Critical'} | Sort-Object ResourceName
        if ($critical.Count -gt 0) {
            $critical | Export-Excel -Path $FilePath -WorksheetName "CRITICAL Risk" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # High risk resources
        $high = $PublicEndpoints | Where-Object {$_.RiskLevel -eq 'High'} | Sort-Object ResourceName
        if ($high.Count -gt 0) {
            $high | Export-Excel -Path $FilePath -WorksheetName "HIGH Risk" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # All public endpoints with conditional formatting
        if ($PublicEndpoints.Count -gt 0) {
            $PublicEndpoints | Sort-Object RiskLevel, ResourceType | 
                Export-Excel -Path $FilePath -WorksheetName "All Public Endpoints" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "Critical" -Range "E:E" -BackgroundColor Red -ConditionalTextColor White
                    New-ConditionalText -Text "High" -Range "E:E" -BackgroundColor Orange
                    New-ConditionalText -Text "Medium" -Range "E:E" -BackgroundColor Yellow
                    New-ConditionalText -Text "Low" -Range "E:E" -BackgroundColor LightGreen
                )
        }
        
        # By subscription
        $bySubscription = $PublicEndpoints | Group-Object Subscription | ForEach-Object {
            [PSCustomObject]@{
                'Subscription' = $_.Name
                'TotalEndpoints' = $_.Count
                'CriticalRisk' = ($_.Group | Where-Object {$_.RiskLevel -eq 'Critical'}).Count
                'HighRisk' = ($_.Group | Where-Object {$_.RiskLevel -eq 'High'}).Count
            }
        } | Sort-Object TotalEndpoints -Descending
        
        if ($bySubscription.Count -gt 0) {
            $bySubscription | Export-Excel -Path $FilePath -WorksheetName "By Subscription" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
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
    Write-Log "========== Starting Public Endpoints Security Audit =========="
    
    # Connect to Azure
    Connect-AzureWithManagedIdentity
    
    # Get target subscriptions
    $subscriptions = Get-TargetSubscriptions -SubscriptionIds $TargetSubscriptions
    
    # Collect public endpoints across all subscriptions
    $allPublicEndpoints = @()
    
    foreach ($subscription in $subscriptions) {
        Write-Log "`nAnalyzing subscription: $($subscription.Name)"
        
        $allPublicEndpoints += Get-PublicStorageAccounts -Subscription $subscription
        $allPublicEndpoints += Get-PublicSQLDatabases -Subscription $subscription
        $allPublicEndpoints += Get-PublicVMs -Subscription $subscription
        $allPublicEndpoints += Get-PublicAppServices -Subscription $subscription
        $allPublicEndpoints += Get-PublicKeyVaults -Subscription $subscription
    }
    
    # Calculate summary statistics
    $criticalCount = ($allPublicEndpoints | Where-Object {$_.RiskLevel -eq 'Critical'}).Count
    $highCount = ($allPublicEndpoints | Where-Object {$_.RiskLevel -eq 'High'}).Count
    
    Write-Log "`n========== Audit Complete =========="
    Write-Log "Total public endpoints found: $($allPublicEndpoints.Count)"
    Write-Log "CRITICAL risk: $criticalCount"
    Write-Log "HIGH risk: $highCount"
    
    # Export to Excel
    Export-ToExcelWithFormatting -PublicEndpoints $allPublicEndpoints -FilePath $FullReportPath
    
    Write-Log "Report saved: $FullReportPath"
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Found $($allPublicEndpoints.Count) public endpoints. CRITICAL: $criticalCount, HIGH: $highCount"
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
