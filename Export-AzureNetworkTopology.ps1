<#
.SYNOPSIS
    Exports comprehensive Azure network topology across all subscriptions
.DESCRIPTION
    Documents VNets, subnets, peerings, NSGs, route tables, and firewall rules
    Generates Excel report with multiple worksheets for different network components
.NOTES
    Author: Jaimin
    Requires: Az.Accounts, Az.Network, ImportExcel modules
    Authentication: Managed Identity
    Graph API Permissions: Mail.Send
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$SubscriptionIds,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipients = "network-team@something.com",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:TEMP\AzureNetworkTopology_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
)

# Error handling
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Connect using Managed Identity
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
    Write-Warning "Could not get Graph token for email. Report will be generated but not emailed."
    $GraphToken = $null
}

# Get subscriptions
if (-not $SubscriptionIds) {
    $Subscriptions = Get-AzSubscription | Where-Object { $_.State -eq 'Enabled' }
}
else {
    $Subscriptions = $SubscriptionIds | ForEach-Object { Get-AzSubscription -SubscriptionId $_ }
}

Write-Output "Analyzing $($Subscriptions.Count) subscription(s)..."

# Initialize collections
$AllVNets = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllSubnets = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllPeerings = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllNSGs = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllNSGRules = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllRouteTables = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllPublicIPs = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllFirewalls = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($Subscription in $Subscriptions) {
    Write-Output "Processing subscription: $($Subscription.Name)"
    Set-AzContext -SubscriptionId $Subscription.Id | Out-Null
    
    # Get VNets
    $VNets = Get-AzVirtualNetwork
    
    foreach ($VNet in $VNets) {
        # VNet details
        $AllVNets.Add([PSCustomObject]@{
                SubscriptionName = $Subscription.Name
                SubscriptionId   = $Subscription.Id
                VNetName         = $VNet.Name
                ResourceGroup    = $VNet.ResourceGroupName
                Location         = $VNet.Location
                AddressSpace     = ($VNet.AddressSpace.AddressPrefixes -join ', ')
                DnsServers       = ($VNet.DhcpOptions.DnsServers -join ', ')
                SubnetCount      = $VNet.Subnets.Count
                PeeringCount     = $VNet.VirtualNetworkPeerings.Count
                Tags             = ($VNet.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
            })
        
        # Subnets
        foreach ($Subnet in $VNet.Subnets) {
            $NSGName = if ($Subnet.NetworkSecurityGroup) { 
                $Subnet.NetworkSecurityGroup.Id.Split('/')[-1] 
            }
            else { 
                'None' 
            }
            
            $RouteTableName = if ($Subnet.RouteTable) { 
                $Subnet.RouteTable.Id.Split('/')[-1] 
            }
            else { 
                'None' 
            }
            
            $AllSubnets.Add([PSCustomObject]@{
                    SubscriptionName = $Subscription.Name
                    VNetName         = $VNet.Name
                    SubnetName       = $Subnet.Name
                    AddressPrefix    = ($Subnet.AddressPrefix -join ', ')
                    NSG              = $NSGName
                    RouteTable       = $RouteTableName
                    ServiceEndpoints = ($Subnet.ServiceEndpoints.Service -join ', ')
                    Delegations      = ($Subnet.Delegations.ServiceName -join ', ')
                    PrivateEndpoints = $Subnet.PrivateEndpoints.Count
                })
        }
        
        # Peerings
        foreach ($Peering in $VNet.VirtualNetworkPeerings) {
            $AllPeerings.Add([PSCustomObject]@{
                    SubscriptionName          = $Subscription.Name
                    SourceVNet                = $VNet.Name
                    PeeringName               = $Peering.Name
                    RemoteVNet                = $Peering.RemoteVirtualNetwork.Id.Split('/')[-1]
                    RemoteSubscription        = $Peering.RemoteVirtualNetwork.Id.Split('/')[2]
                    PeeringState              = $Peering.PeeringState
                    AllowVirtualNetworkAccess = $Peering.AllowVirtualNetworkAccess
                    AllowForwardedTraffic     = $Peering.AllowForwardedTraffic
                    AllowGatewayTransit       = $Peering.AllowGatewayTransit
                    UseRemoteGateways         = $Peering.UseRemoteGateways
                })
        }
    }
    
    # Get NSGs
    $NSGs = Get-AzNetworkSecurityGroup
    
    foreach ($NSG in $NSGs) {
        $AssociatedSubnets = $NSG.Subnets | ForEach-Object { $_.Id.Split('/')[-1] }
        $AssociatedNICs = $NSG.NetworkInterfaces | ForEach-Object { $_.Id.Split('/')[-1] }
        
        $AllNSGs.Add([PSCustomObject]@{
                SubscriptionName   = $Subscription.Name
                NSGName            = $NSG.Name
                ResourceGroup      = $NSG.ResourceGroupName
                Location           = $NSG.Location
                AssociatedSubnets  = ($AssociatedSubnets -join ', ')
                AssociatedNICs     = ($AssociatedNICs -join ', ')
                SecurityRuleCount  = $NSG.SecurityRules.Count
                DefaultRuleCount   = $NSG.DefaultSecurityRules.Count
            })
        
        # NSG Rules
        foreach ($Rule in $NSG.SecurityRules) {
            $AllNSGRules.Add([PSCustomObject]@{
                    SubscriptionName           = $Subscription.Name
                    NSGName                    = $NSG.Name
                    RuleName                   = $Rule.Name
                    Priority                   = $Rule.Priority
                    Direction                  = $Rule.Direction
                    Access                     = $Rule.Access
                    Protocol                   = $Rule.Protocol
                    SourceAddressPrefix        = ($Rule.SourceAddressPrefix -join ', ')
                    SourcePortRange            = ($Rule.SourcePortRange -join ', ')
                    DestinationAddressPrefix   = ($Rule.DestinationAddressPrefix -join ', ')
                    DestinationPortRange       = ($Rule.DestinationPortRange -join ', ')
                    Description                = $Rule.Description
                })
        }
    }
    
    # Get Route Tables
    $RouteTables = Get-AzRouteTable
    
    foreach ($RouteTable in $RouteTables) {
        $AssociatedSubnets = $RouteTable.Subnets | ForEach-Object { $_.Id.Split('/')[-1] }
        
        $AllRouteTables.Add([PSCustomObject]@{
                SubscriptionName          = $Subscription.Name
                RouteTableName            = $RouteTable.Name
                ResourceGroup             = $RouteTable.ResourceGroupName
                Location                  = $RouteTable.Location
                AssociatedSubnets         = ($AssociatedSubnets -join ', ')
                RouteCount                = $RouteTable.Routes.Count
                DisableBgpRoutePropagation = $RouteTable.DisableBgpRoutePropagation
            })
    }
    
    # Get Public IPs
    $PublicIPs = Get-AzPublicIpAddress
    
    foreach ($PIP in $PublicIPs) {
        $AssociatedResource = if ($PIP.IpConfiguration) {
            $PIP.IpConfiguration.Id.Split('/')[-3]
        }
        else {
            'Unassociated'
        }
        
        $AllPublicIPs.Add([PSCustomObject]@{
                SubscriptionName    = $Subscription.Name
                PublicIPName        = $PIP.Name
                ResourceGroup       = $PIP.ResourceGroupName
                Location            = $PIP.Location
                IPAddress           = $PIP.IpAddress
                AllocationMethod    = $PIP.PublicIpAllocationMethod
                SKU                 = $PIP.Sku.Name
                AssociatedResource  = $AssociatedResource
                IdleTimeoutInMinutes = $PIP.IdleTimeoutInMinutes
            })
    }
    
    # Get Azure Firewalls
    $Firewalls = Get-AzFirewall
    
    foreach ($Firewall in $Firewalls) {
        $AllFirewalls.Add([PSCustomObject]@{
                SubscriptionName     = $Subscription.Name
                FirewallName         = $Firewall.Name
                ResourceGroup        = $Firewall.ResourceGroupName
                Location             = $Firewall.Location
                SKU                  = $Firewall.Sku.Name
                Tier                 = $Firewall.Sku.Tier
                ThreatIntelMode      = $Firewall.ThreatIntelMode
                ApplicationRuleCount = $Firewall.ApplicationRuleCollections.Count
                NetworkRuleCount     = $Firewall.NetworkRuleCollections.Count
                NatRuleCount         = $Firewall.NatRuleCollections.Count
            })
    }
}

# Generate Excel Report
Write-Output "Generating Excel report..."

try {
    # Summary Dashboard
    $Summary = [PSCustomObject]@{
        ReportDate        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalSubscriptions = $Subscriptions.Count
        TotalVNets        = $AllVNets.Count
        TotalSubnets      = $AllSubnets.Count
        TotalPeerings     = $AllPeerings.Count
        TotalNSGs         = $AllNSGs.Count
        TotalRouteTables  = $AllRouteTables.Count
        TotalPublicIPs    = $AllPublicIPs.Count
        TotalFirewalls    = $AllFirewalls.Count
    }
    
    # Export to Excel with multiple worksheets
    $Summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -BoldTopRow -FreezeTopRow
    
    if ($AllVNets.Count -gt 0) {
        $AllVNets | Export-Excel -Path $OutputPath -WorksheetName "VirtualNetworks" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($AllSubnets.Count -gt 0) {
        $AllSubnets | Export-Excel -Path $OutputPath -WorksheetName "Subnets" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($AllPeerings.Count -gt 0) {
        $AllPeerings | Export-Excel -Path $OutputPath -WorksheetName "VNetPeerings" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($AllNSGs.Count -gt 0) {
        $AllNSGs | Export-Excel -Path $OutputPath -WorksheetName "NetworkSecurityGroups" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($AllNSGRules.Count -gt 0) {
        $AllNSGRules | Export-Excel -Path $OutputPath -WorksheetName "NSGRules" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($AllRouteTables.Count -gt 0) {
        $AllRouteTables | Export-Excel -Path $OutputPath -WorksheetName "RouteTables" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($AllPublicIPs.Count -gt 0) {
        $AllPublicIPs | Export-Excel -Path $OutputPath -WorksheetName "PublicIPs" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    if ($AllFirewalls.Count -gt 0) {
        $AllFirewalls | Export-Excel -Path $OutputPath -WorksheetName "AzureFirewalls" -AutoSize -BoldTopRow -FreezeTopRow -Append
    }
    
    Write-Output "Excel report generated successfully: $OutputPath"
}
catch {
    Write-Error "Failed to generate Excel report: $_"
    throw
}

# Email the report
if ($GraphToken -and (Test-Path $OutputPath)) {
    try {
        Write-Output "Sending email with report..."
        
        # Read file as base64
        $FileBytes = [System.IO.File]::ReadAllBytes($OutputPath)
        $FileBase64 = [System.Convert]::ToBase64String($FileBytes)
        $FileName = Split-Path $OutputPath -Leaf
        
        # Email body
        $EmailBody = @"
<html>
<body>
<h2>Azure Network Topology Report</h2>
<p>Report generated on: <strong>$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</strong></p>

<h3>Summary:</h3>
<ul>
    <li>Total Subscriptions: <strong>$($Summary.TotalSubscriptions)</strong></li>
    <li>Total Virtual Networks: <strong>$($Summary.TotalVNets)</strong></li>
    <li>Total Subnets: <strong>$($Summary.TotalSubnets)</strong></li>
    <li>Total VNet Peerings: <strong>$($Summary.TotalPeerings)</strong></li>
    <li>Total NSGs: <strong>$($Summary.TotalNSGs)</strong></li>
    <li>Total Route Tables: <strong>$($Summary.TotalRouteTables)</strong></li>
    <li>Total Public IPs: <strong>$($Summary.TotalPublicIPs)</strong></li>
    <li>Total Azure Firewalls: <strong>$($Summary.TotalFirewalls)</strong></li>
</ul>

<p>Please review the attached Excel workbook for complete network topology details.</p>

<p><em>This is an automated report from Azure Automation.</em></p>
</body>
</html>
"@
        
        $EmailMessage = @{
            message = @{
                subject      = "Azure Network Topology Report - $(Get-Date -Format 'yyyy-MM-dd')"
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
        
        Write-Output "Email sent successfully to $EmailRecipients"
    }
    catch {
        Write-Warning "Failed to send email: $_"
    }
}

Write-Output "Script completed successfully"