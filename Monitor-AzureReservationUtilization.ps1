#Requires -Modules @{ ModuleName="Az.Accounts"; ModuleVersion="2.0.0" }
#Requires -Modules @{ ModuleName="Az.Reservations"; ModuleVersion="1.0.0" }
#Requires -Modules @{ ModuleName="Az.Billing"; ModuleVersion="2.0.0" }
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Monitors Azure Reserved Instance utilization to maximize ROI.

.DESCRIPTION
    This runbook tracks Reserved Instance (RI) utilization across subscriptions:
    - Reservation usage percentage
    - Underutilized reservations
    - Reservation expiration dates
    - VM size distribution vs. reservations
    - Cost impact of underutilization
    - Recommendations for reservation adjustments
    
    Helps ensure maximum value from Azure reservations.

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Azure Automation account with Managed Identity
    Azure RBAC Permissions Required:
        - Reservation Reader
        - Billing Reader
        - Reader on subscriptions
#>

#region Configuration
# Email Configuration
$EmailRecipients = @("finance@yourdomain.com", "cloudops@yourdomain.com")
$EmailFrom = "azureautomation@yourdomain.com"
$EmailSubject = "Azure Reservation Utilization Report - $(Get-Date -Format 'yyyy-MM-dd')"

# Utilization Thresholds
$UtilizationThresholds = @{
    Critical = 50      # < 50% utilization is critical
    Warning = 70       # < 70% utilization is warning
    Good = 85          # >= 85% is good utilization
}

# Expiration Alert Thresholds (days)
$ExpirationAlerts = @{
    Critical = 30      # Expires in 30 days or less
    Warning = 90       # Expires in 90 days or less
    Info = 180         # Expires in 180 days or less
}

# Analysis Period (days to look back for utilization data)
$AnalysisPeriodDays = 7

# Report Settings
$ExportPath = $env:TEMP
$ReportFileName = "ReservationUtilization_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
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

function Get-UtilizationLevel {
    param([decimal]$UtilizationPercentage)
    
    if ($UtilizationPercentage -lt $UtilizationThresholds.Critical) {
        return 'CRITICAL'
    }
    elseif ($UtilizationPercentage -lt $UtilizationThresholds.Warning) {
        return 'WARNING'
    }
    elseif ($UtilizationPercentage -ge $UtilizationThresholds.Good) {
        return 'GOOD'
    }
    else {
        return 'FAIR'
    }
}

function Get-ExpirationStatus {
    param([datetime]$ExpiryDate)
    
    $daysUntilExpiry = [int](($ExpiryDate - (Get-Date)).TotalDays)
    
    if ($daysUntilExpiry -le $ExpirationAlerts.Critical) {
        return @{Status = 'CRITICAL'; Days = $daysUntilExpiry}
    }
    elseif ($daysUntilExpiry -le $ExpirationAlerts.Warning) {
        return @{Status = 'WARNING'; Days = $daysUntilExpiry}
    }
    elseif ($daysUntilExpiry -le $ExpirationAlerts.Info) {
        return @{Status = 'INFO'; Days = $daysUntilExpiry}
    }
    else {
        return @{Status = 'OK'; Days = $daysUntilExpiry}
    }
}

function Get-ReservationOrders {
    Write-Log "Retrieving reservation orders..."
    
    try {
        # Get all reservation orders
        $reservationOrders = Get-AzReservationOrder
        
        Write-Log "Found $($reservationOrders.Count) reservation order(s)"
        return $reservationOrders
    }
    catch {
        Write-Log "Error retrieving reservation orders: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Get-ReservationDetails {
    param([object]$ReservationOrder)
    
    Write-Log "  Processing reservation order: $($ReservationOrder.Name)"
    
    try {
        $results = @()
        
        # Get reservations in this order
        $reservations = Get-AzReservation -ReservationOrderId $ReservationOrder.Name
        
        foreach ($reservation in $reservations) {
            Write-Log "    Analyzing reservation: $($reservation.Name)"
            
            # Parse reservation properties
            $properties = $reservation.Properties
            
            # Calculate days until expiration
            $expiryDate = $properties.ExpiryDate
            $expirationInfo = Get-ExpirationStatus -ExpiryDate $expiryDate
            
            # Get utilization data (last 7 days average)
            # Note: Actual utilization requires Azure Consumption API
            # This is a simplified example - you'll need to use Get-AzConsumptionReservationSummary
            $utilizationPercentage = 0
            
            try {
                # Get utilization summaries
                $endDate = Get-Date
                $startDate = $endDate.AddDays(-$AnalysisPeriodDays)
                
                # This requires appropriate permissions
                $utilizationData = Get-AzConsumptionReservationSummary `
                    -ReservationOrderId $ReservationOrder.Name `
                    -ReservationId $reservation.Name `
                    -StartDate $startDate.ToString('yyyy-MM-dd') `
                    -EndDate $endDate.ToString('yyyy-MM-dd') `
                    -Grain 'Daily' `
                    -ErrorAction SilentlyContinue
                
                if ($utilizationData) {
                    $avgUtilization = ($utilizationData | Measure-Object -Property UtilizationPercentage -Average).Average
                    $utilizationPercentage = [math]::Round($avgUtilization, 2)
                }
            }
            catch {
                Write-Log "      Could not retrieve utilization data: $($_.Exception.Message)" -Level WARNING
                $utilizationPercentage = $null
            }
            
            # Determine utilization level
            $utilizationLevel = if ($null -ne $utilizationPercentage) {
                Get-UtilizationLevel -UtilizationPercentage $utilizationPercentage
            } else {
                'UNKNOWN'
            }
            
            # Calculate potential waste
            $monthlyCost = if ($properties.BillingPlan -eq 'Monthly') {
                # Simplified - actual cost from properties
                $properties.Amount
            } else {
                $properties.Amount / 12  # Approximate monthly from yearly
            }
            
            $wastedCost = if ($null -ne $utilizationPercentage -and $utilizationPercentage -lt 100) {
                [math]::Round($monthlyCost * (100 - $utilizationPercentage) / 100, 2)
            } else {
                0
            }
            
            # Generate recommendations
            $recommendations = @()
            
            if ($utilizationLevel -eq 'CRITICAL') {
                $recommendations += "Consider exchanging or canceling this reservation"
                $recommendations += "Review VM deployment and utilization patterns"
            }
            elseif ($utilizationLevel -eq 'WARNING') {
                $recommendations += "Deploy more matching resources to increase utilization"
                $recommendations += "Consider reservation scope adjustment"
            }
            
            if ($expirationInfo.Status -in @('CRITICAL', 'WARNING')) {
                $recommendations += "Plan renewal or adjustment before expiration"
            }
            
            $results += [PSCustomObject]@{
                'ReservationOrderId' = $ReservationOrder.Name
                'ReservationId' = $reservation.Name
                'DisplayName' = $properties.DisplayName
                'ReservedResourceType' = $properties.ReservedResourceType
                'Quantity' = $properties.Quantity
                'ProvisioningState' = $properties.ProvisioningState
                'InstanceFlexibility' = $properties.InstanceFlexibility
                'Location' = $properties.Location
                'SKU' = $properties.SkuDescription
                'Term' = $properties.Term
                'BillingPlan' = $properties.BillingPlan
                'PurchaseDate' = $properties.PurchaseDate.ToString('yyyy-MM-dd')
                'ExpiryDate' = $expiryDate.ToString('yyyy-MM-dd')
                'DaysUntilExpiry' = $expirationInfo.Days
                'ExpiryStatus' = $expirationInfo.Status
                'UtilizationPercentage' = if ($null -ne $utilizationPercentage) { "$utilizationPercentage%" } else { 'No Data' }
                'UtilizationLevel' = $utilizationLevel
                'EstimatedMonthlyCost' = [math]::Round($monthlyCost, 2)
                'EstimatedMonthlyWaste' = [math]::Round($wastedCost, 2)
                'EstimatedAnnualWaste' = [math]::Round($wastedCost * 12, 2)
                'Recommendations' = if ($recommendations.Count -gt 0) { $recommendations -join '; ' } else { 'None' }
                'AppliedScopes' = if ($properties.AppliedScopes) { $properties.AppliedScopes -join '; ' } else { 'Shared' }
                'Scope' = $properties.AppliedScopeType
            }
        }
        
        Write-Log "    Processed $($results.Count) reservation(s)"
        return $results
    }
    catch {
        Write-Log "    Error processing reservation order: $($_.Exception.Message)" -Level WARNING
        return @()
    }
}

function Calculate-TotalWaste {
    param([object[]]$Reservations)
    
    $totalMonthlyWaste = 0
    $totalAnnualWaste = 0
    
    foreach ($reservation in $Reservations) {
        if ($reservation.EstimatedMonthlyWaste -is [decimal] -or $reservation.EstimatedMonthlyWaste -is [double]) {
            $totalMonthlyWaste += $reservation.EstimatedMonthlyWaste
            $totalAnnualWaste += $reservation.EstimatedAnnualWaste
        }
    }
    
    return @{
        MonthlyWaste = [math]::Round($totalMonthlyWaste, 2)
        AnnualWaste = [math]::Round($totalAnnualWaste, 2)
    }
}

function Export-ToExcelWithFormatting {
    param(
        [object[]]$Reservations,
        [string]$FilePath
    )
    
    Write-Log "Creating Excel report with formatting..."
    
    try {
        # Calculate statistics
        $totalReservations = $Reservations.Count
        $criticalUtilization = ($Reservations | Where-Object {$_.UtilizationLevel -eq 'CRITICAL'}).Count
        $warningUtilization = ($Reservations | Where-Object {$_.UtilizationLevel -eq 'WARNING'}).Count
        $goodUtilization = ($Reservations | Where-Object {$_.UtilizationLevel -eq 'GOOD'}).Count
        
        $expiringCritical = ($Reservations | Where-Object {$_.ExpiryStatus -eq 'CRITICAL'}).Count
        $expiringWarning = ($Reservations | Where-Object {$_.ExpiryStatus -eq 'WARNING'}).Count
        
        $waste = Calculate-TotalWaste -Reservations $Reservations
        
        # Calculate average utilization (only for reservations with data)
        $reservationsWithData = $Reservations | Where-Object {$_.UtilizationPercentage -ne 'No Data'}
        $avgUtilization = if ($reservationsWithData.Count -gt 0) {
            $utilizationValues = $reservationsWithData | ForEach-Object {
                [decimal]($_.UtilizationPercentage -replace '%', '')
            }
            [math]::Round(($utilizationValues | Measure-Object -Average).Average, 2)
        } else {
            0
        }
        
        # Create summary
        $summary = [PSCustomObject]@{
            'Metric' = @(
                'Total Reservations',
                '--- Utilization Status ---',
                'Critical (< 50%)',
                'Warning (< 70%)',
                'Fair (70-85%)',
                'Good (>= 85%)',
                'Unknown (No Data)',
                'Average Utilization',
                '--- Expiration Status ---',
                'Expiring Critical (≤ 30 days)',
                'Expiring Warning (≤ 90 days)',
                '--- Cost Impact ---',
                'Estimated Monthly Waste',
                'Estimated Annual Waste',
                'Analysis Period',
                'Report Generated'
            )
            'Value' = @(
                $totalReservations,
                '',
                $criticalUtilization,
                $warningUtilization,
                ($Reservations | Where-Object {$_.UtilizationLevel -eq 'FAIR'}).Count,
                $goodUtilization,
                ($Reservations | Where-Object {$_.UtilizationLevel -eq 'UNKNOWN'}).Count,
                "$avgUtilization%",
                '',
                $expiringCritical,
                $expiringWarning,
                '',
                "`$$($waste.MonthlyWaste)",
                "`$$($waste.AnnualWaste)",
                "$AnalysisPeriodDays days",
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            )
        }
        
        # Export Summary
        $summary | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        
        # Underutilized reservations (critical + warning)
        $underutilized = $Reservations | Where-Object {$_.UtilizationLevel -in @('CRITICAL', 'WARNING')} | 
                         Sort-Object EstimatedAnnualWaste -Descending
        
        if ($underutilized.Count -gt 0) {
            $underutilized | Export-Excel -Path $FilePath -WorksheetName "Underutilized" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # Expiring soon
        $expiring = $Reservations | Where-Object {$_.ExpiryStatus -in @('CRITICAL', 'WARNING')} | 
                    Sort-Object DaysUntilExpiry
        
        if ($expiring.Count -gt 0) {
            $expiring | Export-Excel -Path $FilePath -WorksheetName "Expiring Soon" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # By resource type
        $byType = $Reservations | Group-Object ReservedResourceType | ForEach-Object {
            $typeReservations = $_.Group
            $avgUtil = if ($typeReservations.UtilizationPercentage -ne 'No Data') {
                $utilValues = $typeReservations | Where-Object {$_.UtilizationPercentage -ne 'No Data'} | ForEach-Object {
                    [decimal]($_.UtilizationPercentage -replace '%', '')
                }
                if ($utilValues.Count -gt 0) {
                    [math]::Round(($utilValues | Measure-Object -Average).Average, 2)
                } else {
                    0
                }
            } else {
                0
            }
            
            [PSCustomObject]@{
                'ResourceType' = $_.Name
                'Count' = $_.Count
                'AverageUtilization' = "$avgUtil%"
                'CriticalCount' = ($typeReservations | Where-Object {$_.UtilizationLevel -eq 'CRITICAL'}).Count
                'WarningCount' = ($typeReservations | Where-Object {$_.UtilizationLevel -eq 'WARNING'}).Count
            }
        } | Sort-Object Count -Descending
        
        if ($byType.Count -gt 0) {
            $byType | Export-Excel -Path $FilePath -WorksheetName "By Resource Type" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
        }
        
        # All reservations with conditional formatting
        if ($Reservations.Count -gt 0) {
            $Reservations | Sort-Object UtilizationLevel, EstimatedAnnualWaste -Descending | 
                Export-Excel -Path $FilePath -WorksheetName "All Reservations" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow `
                -ConditionalText $(
                    New-ConditionalText -Text "CRITICAL" -Range "O:O" -BackgroundColor Red -ConditionalTextColor White
                    New-ConditionalText -Text "WARNING" -Range "O:O" -BackgroundColor Orange
                    New-ConditionalText -Text "FAIR" -Range "O:O" -BackgroundColor Yellow
                    New-ConditionalText -Text "GOOD" -Range "O:O" -BackgroundColor LightGreen
                    New-ConditionalText -Text "CRITICAL" -Range "M:M" -BackgroundColor Red -ConditionalTextColor White
                    New-ConditionalText -Text "WARNING" -Range "M:M" -BackgroundColor Orange
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
    Write-Log "========== Starting Azure Reservation Utilization Monitoring =========="
    
    # Connect to Azure
    Connect-AzureWithManagedIdentity
    
    # Get all reservation orders
    $reservationOrders = Get-ReservationOrders
    
    if ($reservationOrders.Count -eq 0) {
        Write-Log "No reservation orders found" -Level WARNING
        Write-Output "WARNING: No reservation orders found in this Azure environment"
        exit 0
    }
    
    # Process each reservation order
    $allReservations = @()
    
    foreach ($order in $reservationOrders) {
        $allReservations += Get-ReservationDetails -ReservationOrder $order
    }
    
    # Calculate summary statistics
    $criticalCount = ($allReservations | Where-Object {$_.UtilizationLevel -eq 'CRITICAL'}).Count
    $warningCount = ($allReservations | Where-Object {$_.UtilizationLevel -eq 'WARNING'}).Count
    $waste = Calculate-TotalWaste -Reservations $allReservations
    
    Write-Log "`n========== Analysis Complete =========="
    Write-Log "Total reservations: $($allReservations.Count)"
    Write-Log "Critical utilization: $criticalCount"
    Write-Log "Warning utilization: $warningCount"
    Write-Log "Estimated annual waste: `$$($waste.AnnualWaste)"
    
    # Export to Excel
    Export-ToExcelWithFormatting -Reservations $allReservations -FilePath $FullReportPath
    
    Write-Log "Report saved: $FullReportPath"
    Write-Log "========== Runbook Completed Successfully =========="
    Write-Output "SUCCESS: Analyzed $($allReservations.Count) reservations. Potential annual waste: `$$($waste.AnnualWaste)"
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
