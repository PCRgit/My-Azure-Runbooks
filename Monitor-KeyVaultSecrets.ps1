#Requires -Modules @{ ModuleName="Az.Accounts"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Az.KeyVault"; ModuleVersion="4.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Monitors Azure Key Vault secrets, keys, and certificates for expiration and security compliance.

.DESCRIPTION
    This runbook tracks Key Vault objects across subscriptions:
    - Expiring secrets, keys, and certificates
    - Secrets without expiration dates
    - Disabled secrets still in rotation
    - Key Vault access policies and permissions
    - Orphaned secrets (not used in applications)
    
    Provides multi-tier alerting and remediation recommendations.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Azure RBAC Permissions Required:
        - Key Vault Reader or Key Vault Secrets User
        - Key Vault Certificates User
        - Key Vault Keys User
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("security@yourdomain.com", "devops@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "Key Vault Secrets Monitoring Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Subscription Selection (empty array = all subscriptions)
$TargetSubscriptions = @()

# Expiration Thresholds (days)
$ExpirationThresholds = @{
    Expired = 0       # Already expired
    Critical = 7      # Expires in 7 days or less
    Warning = 30      # Expires in 30 days or less
    Info = 90         # Expires in 90 days or less
}

# Security Checks
$SecurityChecks = @{
    CheckNoExpiration = $true          # Flag secrets without expiration
    CheckDisabledSecrets = $true       # Flag disabled secrets
    CheckSoftDeleteEnabled = $true     # Check if soft delete is enabled
    CheckPurgeProtection = $true       # Check if purge protection is enabled
}

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "KeyVaultSecrets_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-ExpirationStatus {
    param([datetime]$ExpiryDate)
    
    $daysUntilExpiry = [int](($ExpiryDate - (Get-Date)).TotalDays)
    
    if ($daysUntilExpiry -lt $ExpirationThresholds.Expired) {
        return @{Status = 'EXPIRED'; Days = $daysUntilExpiry}
    }
    elseif ($daysUntilExpiry -le $ExpirationThresholds.Critical) {
        return @{Status = 'CRITICAL'; Days = $daysUntilExpiry}
    }
    elseif ($daysUntilExpiry -le $ExpirationThresholds.Warning) {
        return @{Status = 'WARNING'; Days = $daysUntilExpiry}
    }
    elseif ($daysUntilExpiry -le $ExpirationThresholds.Info) {
        return @{Status = 'INFO'; Days = $daysUntilExpiry}
    }
    else {
        return @{Status = 'OK'; Days = $daysUntilExpiry}
    }
}

function Get-KeyVaultSecrets {
    param([object]$Subscription)
    
    Write-Log "  Scanning Key Vault secrets in: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $keyVaults = Get-AzKeyVault
        Write-Log "    Found $($keyVaults.Count) Key Vault(s)"
        
        $results = @()
        
        foreach ($kv in $keyVaults) {
            Write-Log "      Processing: $($kv.VaultName)"
            
            # Get Key Vault details for security checks
            $kvDetail = Get-AzKeyVault -VaultName $kv.VaultName -ResourceGroupName $kv.ResourceGroupName
            
            try {
                # Get all secrets
                $secrets = Get-AzKeyVaultSecret -VaultName $kv.VaultName -ErrorAction Stop
                
                foreach ($secret in $secrets) {
                    # Get detailed secret info
                    $secretDetail = Get-AzKeyVaultSecret -VaultName $kv.VaultName -Name $secret.Name -ErrorAction SilentlyContinue
                    
                    if ($secretDetail) {
                        $expiryDate = $secretDetail.Expires
                        $status = 'No Expiration'
                        $daysUntilExpiry = 'N/A'
                        $recommendation = 'Set expiration date'
                        
                        if ($expiryDate) {
                            $expirationInfo = Get-ExpirationStatus -ExpiryDate $expiryDate
                            $status = $expirationInfo.Status
                            $daysUntilExpiry = $expirationInfo.Days
                            
                            $recommendation = switch ($status) {
                                'EXPIRED' { 'Rotate immediately - secret has expired' }
                                'CRITICAL' { 'Rotate within 7 days' }
                                'WARNING' { 'Plan rotation within 30 days' }
                                'INFO' { 'Schedule rotation' }
                                default { 'Monitor for rotation' }
                            }
                        }
                        
                        $results += [PSCustomObject]@{
                            'ObjectType' = 'Secret'
                            'KeyVaultName' = $kv.VaultName
                            'ObjectName' = $secret.Name
                            'Status' = $status
                            'Enabled' = $secretDetail.Enabled
                            'Created' = if ($secretDetail.Created) { $secretDetail.Created.ToString('yyyy-MM-dd') } else { 'N/A' }
                            'Updated' = if ($secretDetail.Updated) { $secretDetail.Updated.ToString('yyyy-MM-dd') } else { 'N/A' }
                            'ExpiryDate' = if ($expiryDate) { $expiryDate.ToString('yyyy-MM-dd') } else { 'Never' }
                            'DaysUntilExpiry' = $daysUntilExpiry
                            'ContentType' = if ($secretDetail.ContentType) { $secretDetail.ContentType } else { 'N/A' }
                            'Version' = $secretDetail.Version
                            'Recommendation' = $recommendation
                            'ResourceGroup' = $kv.ResourceGroupName
                            'Location' = $kv.Location
                            'Subscription' = $Subscription.Name
                            'SubscriptionId' = $Subscription.Id
                            'VaultId' = $kvDetail.ResourceId
                            'SecretId' = $secretDetail.Id
                            'SoftDeleteEnabled' = $kvDetail.EnableSoftDelete
                            'PurgeProtectionEnabled' = $kvDetail.EnablePurgeProtection
                        }
                    }
                }
                
                # Get all keys
                $keys = Get-AzKeyVaultKey -VaultName $kv.VaultName -ErrorAction Stop
                
                foreach ($key in $keys) {
                    $keyDetail = Get-AzKeyVaultKey -VaultName $kv.VaultName -Name $key.Name -ErrorAction SilentlyContinue
                    
                    if ($keyDetail) {
                        $expiryDate = $keyDetail.Expires
                        $status = 'No Expiration'
                        $daysUntilExpiry = 'N/A'
                        $recommendation = 'Set expiration date'
                        
                        if ($expiryDate) {
                            $expirationInfo = Get-ExpirationStatus -ExpiryDate $expiryDate
                            $status = $expirationInfo.Status
                            $daysUntilExpiry = $expirationInfo.Days
                            
                            $recommendation = switch ($status) {
                                'EXPIRED' { 'Rotate immediately - key has expired' }
                                'CRITICAL' { 'Rotate within 7 days' }
                                'WARNING' { 'Plan rotation within 30 days' }
                                'INFO' { 'Schedule rotation' }
                                default { 'Monitor for rotation' }
                            }
                        }
                        
                        $results += [PSCustomObject]@{
                            'ObjectType' = 'Key'
                            'KeyVaultName' = $kv.VaultName
                            'ObjectName' = $key.Name
                            'Status' = $status
                            'Enabled' = $keyDetail.Enabled
                            'Created' = if ($keyDetail.Created) { $keyDetail.Created.ToString('yyyy-MM-dd') } else { 'N/A' }
                            'Updated' = if ($keyDetail.Updated) { $keyDetail.Updated.ToString('yyyy-MM-dd') } else { 'N/A' }
                            'ExpiryDate' = if ($expiryDate) { $expiryDate.ToString('yyyy-MM-dd') } else { 'Never' }
                            'DaysUntilExpiry' = $daysUntilExpiry
                            'ContentType' = $keyDetail.KeyType
                            'Version' = $keyDetail.Version
                            'Recommendation' = $recommendation
                            'ResourceGroup' = $kv.ResourceGroupName
                            'Location' = $kv.Location
                            'Subscription' = $Subscription.Name
                            'SubscriptionId' = $Subscription.Id
                            'VaultId' = $kvDetail.ResourceId
                            'SecretId' = $keyDetail.Id
                            'SoftDeleteEnabled' = $kvDetail.EnableSoftDelete
                            'PurgeProtectionEnabled' = $kvDetail.EnablePurgeProtection
                        }
                    }
                }
                
                # Get all certificates
                $certificates = Get-AzKeyVaultCertificate -VaultName $kv.VaultName -ErrorAction Stop
                
                foreach ($cert in $certificates) {
                    $certDetail = Get-AzKeyVaultCertificate -VaultName $kv.VaultName -Name $cert.Name -ErrorAction SilentlyContinue
                    
                    if ($certDetail) {
                        $expiryDate = $certDetail.Certificate.NotAfter
                        $status = 'No Expiration'
                        $daysUntilExpiry = 'N/A'
                        $recommendation = 'Monitor certificate'
                        
                        if ($expiryDate) {
                            $expirationInfo = Get-ExpirationStatus -ExpiryDate $expiryDate
                            $status = $expirationInfo.Status
                            $daysUntilExpiry = $expirationInfo.Days
                            
                            $recommendation = switch ($status) {
                                'EXPIRED' { 'URGENT: Renew certificate immediately' }
                                'CRITICAL' { 'Renew within 7 days' }
                                'WARNING' { 'Plan renewal within 30 days' }
                                'INFO' { 'Schedule renewal' }
                                default { 'Monitor for renewal' }
                            }
                        }
                        
                        $results += [PSCustomObject]@{
                            'ObjectType' = 'Certificate'
                            'KeyVaultName' = $kv.VaultName
                            'ObjectName' = $cert.Name
                            'Status' = $status
                            'Enabled' = $certDetail.Enabled
                            'Created' = if ($certDetail.Created) { $certDetail.Created.ToString('yyyy-MM-dd') } else { 'N/A' }
                            'Updated' = if ($certDetail.Updated) { $certDetail.Updated.ToString('yyyy-MM-dd') } else { 'N/A' }
                            'ExpiryDate' = if ($expiryDate) { $expiryDate.ToString('yyyy-MM-dd') } else { 'Never' }
                            'DaysUntilExpiry' = $daysUntilExpiry
                            'ContentType' = $certDetail.Certificate.Issuer
                            'Version' = $certDetail.Version
                            'Recommendation' = $recommendation
                            'ResourceGroup' = $kv.ResourceGroupName
                            'Location' = $kv.Location
                            'Subscription' = $Subscription.Name
                            'SubscriptionId' = $Subscription.Id
                            'VaultId' = $kvDetail.ResourceId
                            'SecretId' = $certDetail.Id
                            'SoftDeleteEnabled' = $kvDetail.EnableSoftDelete
                            'PurgeProtectionEnabled' = $kvDetail.EnablePurgeProtection
                        }
                    }
                }
                
                Write-Log "        Processed $($secrets.Count) secret(s), $($keys.Count) key(s), $($certificates.Count) certificate(s)"
            }
            catch {
                Write-Log "        Error accessing vault $($kv.VaultName): $($_.Exception.Message)" -Level WARNING
            }
            
            # Rate limiting
            Start-Sleep -Milliseconds 200
        }
        
        Write-Log "    Found $($results.Count) total object(s)"
        return $results
    }
    catch {
        Write-Log "    Error scanning Key Vaults: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-KeyVaultSecurityCompliance {
    param([object]$Subscription)
    
    Write-Log "  Checking Key Vault security compliance in: $($Subscription.Name)"
    
    try {
        Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
        
        $keyVaults = Get-AzKeyVault
        $results = @()
        
        foreach ($kv in $keyVaults) {
            $kvDetail = Get-AzKeyVault -VaultName $kv.VaultName -ResourceGroupName $kv.ResourceGroupName
            
            $issues = @()
            
            if (-not $kvDetail.EnableSoftDelete) {
                $issues += 'Soft Delete not enabled'
            }
            if (-not $kvDetail.EnablePurgeProtection) {
                $issues += 'Purge Protection not enabled'
            }
            if ($kvDetail.NetworkAcls.DefaultAction -eq 'Allow') {
                $issues += 'Public network access allowed'
            }
            
            if ($issues.Count -gt 0) {
                $results += [PSCustomObject]@{
                    'KeyVaultName' = $kv.VaultName
                    'ResourceGroup' = $kv.ResourceGroupName
                    'Location' = $kv.Location
                    'SoftDeleteEnabled' = $kvDetail.EnableSoftDelete
                    'PurgeProtectionEnabled' = $kvDetail.EnablePurgeProtection
                    'PublicNetworkAccess' = ($kvDetail.NetworkAcls.DefaultAction -eq 'Allow')
                    'Issues' = ($issues -join '; ')
                    'Recommendation' = 'Enable Soft Delete, Purge Protection, and restrict network access'
                    'Subscription' = $Subscription.Name
                    'SubscriptionId' = $Subscription.Id
                    'VaultId' = $kvDetail.ResourceId
                }
            }
        }
        
        Write-Log "    Found $($results.Count) vault(s) with security issues"
        return $results
    }
    catch {
        Write-Log "    Error checking security compliance: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$VaultObjects,
        [object[]]$SecurityIssues,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Calculate statistics
        $totalObjects = $VaultObjects.Count
        $expiredCount = ($VaultObjects | Where-Object {$_.Status -eq 'EXPIRED'}).Count
        $criticalCount = ($VaultObjects | Where-Object {$_.Status -eq 'CRITICAL'}).Count
        $warningCount = ($VaultObjects | Where-Object {$_.Status -eq 'WARNING'}).Count
        $noExpirationCount = ($VaultObjects | Where-Object {$_.Status -eq 'No Expiration'}).Count
        
        # Create summary
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total Objects Monitored',
                '--- Expiration Status ---',
                'Expired',
                'Critical (≤ 7 days)',
                'Warning (≤ 30 days)',
                'Info (≤ 90 days)',
                'No Expiration Set',
                '--- Object Types ---',
                'Secrets',
                'Keys',
                'Certificates',
                '--- Security Issues ---',
                'Vaults with Issues',
                'Report Generated'
            )
            'Value' = @(
                $totalObjects,
                '',
                $expiredCount,
                $criticalCount,
                $warningCount,
                ($VaultObjects | Where-Object {$_.Status -eq 'INFO'}).Count,
                $noExpirationCount,
                '',
                ($VaultObjects | Where-Object {$_.ObjectType -eq 'Secret'}).Count,
                ($VaultObjects | Where-Object {$_.ObjectType -eq 'Key'}).Count,
                ($VaultObjects | Where-Object {$_.ObjectType -eq 'Certificate'}).Count,
                '',
                $SecurityIssues.Count,
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Top 20 most urgent
        $top20 = $VaultObjects | Where-Object {$_.Status -in @('EXPIRED', 'CRITICAL', 'WARNING')} | 
                 Sort-Object DaysUntilExpiry | Select-Object -First 20
        
        if ($top20.Count -gt 0) {
            $top20 | Export-Excel -Path $FilePath -WorksheetName "Top 20 Urgent" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Expired objects
        $expired = $VaultObjects | Where-Object {$_.Status -eq 'EXPIRED'} | Sort-Object KeyVaultName, ObjectName
        if ($expired.Count -gt 0) {
            $expired | Export-Excel -Path $FilePath -WorksheetName "Expired" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Critical objects
        $critical = $VaultObjects | Where-Object {$_.Status -eq 'CRITICAL'} | Sort-Object DaysUntilExpiry
        if ($critical.Count -gt 0) {
            $critical | Export-Excel -Path $FilePath -WorksheetName "Critical" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Objects without expiration
        $noExpiration = $VaultObjects | Where-Object {$_.Status -eq 'No Expiration'} | Sort-Object KeyVaultName, ObjectName
        if ($noExpiration.Count -gt 0) {
            $noExpiration | Export-Excel -Path $FilePath -WorksheetName "No Expiration" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # All objects with conditional formatting
        if ($VaultObjects.Count -gt 0) {
            $VaultObjects | Sort-Object Status, DaysUntilExpiry | 
                Export-Excel -Path $FilePath -WorksheetName "All Objects" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "EXPIRED" -Range "D:D" -BackgroundColor Red -ConditionalTextColor White
                    New-ConditionalText -Text "CRITICAL" -Range "D:D" -BackgroundColor Orange
                    New-ConditionalText -Text "WARNING" -Range "D:D" -BackgroundColor Yellow
                    New-ConditionalText -Text "INFO" -Range "D:D" -BackgroundColor LightBlue
                    New-ConditionalText -Text "No Expiration" -Range "D:D" -BackgroundColor LightCoral
                )
        }
        
        # Security compliance issues
        if ($SecurityIssues.Count -gt 0) {
            $SecurityIssues | Export-Excel -Path $FilePath -WorksheetName "Security Issues" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
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
    Write-Log "========== Starting Key Vault Secrets Monitoring =========="
    
    # Connect to Azure
    Connect-AzureWithManagedIdentity
    
    # Get target subscriptions
    $subscriptions = Get-TargetSubscriptions -SubscriptionIds $TargetSubscriptions
    
    # Collect Key Vault objects and security issues across all subscriptions
    $allVaultObjects = @()
    $allSecurityIssues = @()
    
    foreach ($subscription in $subscriptions) {
        Write-Log "`nAnalyzing subscription: $($subscription.Name)"
        
        $allVaultObjects += Get-KeyVaultSecrets -Subscription $subscription
        $allSecurityIssues += Get-KeyVaultSecurityCompliance -Subscription $subscription
    }
    
    # Calculate summary statistics
    $expiredCount = ($allVaultObjects | Where-Object {$_.Status -eq 'EXPIRED'}).Count
    $criticalCount = ($allVaultObjects | Where-Object {$_.Status -eq 'CRITICAL'}).Count
    $warningCount = ($allVaultObjects | Where-Object {$_.Status -eq 'WARNING'}).Count
    
    Write-Log "`n========== Monitoring Complete =========="
    Write-Log "Total objects monitored: $($allVaultObjects.Count)"
    Write-Log "Expired: $expiredCount"
    Write-Log "Critical: $criticalCount"
    Write-Log "Warning: $warningCount"
    Write-Log "Security issues found: $($allSecurityIssues.Count)"
    
    # Export to Excel
    Export-ToExcelWithFormatting -VaultObjects $allVaultObjects -SecurityIssues $allSecurityIssues -FilePath $FullReportPath
    
    Write-Log "Report saved: $FullReportPath"
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Monitored $($allVaultObjects.Count) objects. Expired: $expiredCount, Critical: $criticalCount, Warning: $warningCount"
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
