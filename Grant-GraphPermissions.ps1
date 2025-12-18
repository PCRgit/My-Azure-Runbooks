<#
.SYNOPSIS
    Grants Microsoft Graph API permissions to Azure Automation Managed Identity.

.DESCRIPTION
    This script automates the process of granting required Microsoft Graph API
    permissions to an Azure Automation account's Managed Identity. It handles
    all the permissions needed for the runbook collection.

.PARAMETER ManagedIdentityObjectId
    The Object ID of the Azure Automation Managed Identity.
    Find this in: Azure Portal > Automation Account > Identity > Object ID

.PARAMETER RunbookSelection
    Optional. Specify which runbooks you plan to use to grant only required permissions.
    Valid values: All, StaleDevices, AppSecrets, MFACompliance, LicenseUsage, AdminChanges
    Default: All

.EXAMPLE
    .\Grant-GraphPermissions.ps1 -ManagedIdentityObjectId "12345678-1234-1234-1234-123456789012"

.EXAMPLE
    .\Grant-GraphPermissions.ps1 -ManagedIdentityObjectId "12345678-1234-1234-1234-123456789012" -RunbookSelection "StaleDevices","MFACompliance"

.NOTES
    Author: Your Name
    Version: 1.0
    Requires: Microsoft.Graph.Authentication module
    Requires: Appropriate Azure AD permissions (Application.ReadWrite.All or AppRoleAssignment.ReadWrite.All)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Object ID of the Managed Identity")]
    [ValidateNotNullOrEmpty()]
    [string]$ManagedIdentityObjectId,
    
    [Parameter(Mandatory=$false, HelpMessage="Which runbooks to configure permissions for")]
    [ValidateSet('All', 'StaleDevices', 'AppSecrets', 'MFACompliance', 'LicenseUsage', 'AdminChanges')]
    [string[]]$RunbookSelection = @('All')
)

#region Functions
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = 'White'
    )
    Write-Host $Message -ForegroundColor $Color
}

function Test-Prerequisites {
    Write-ColorOutput "`n=== Checking Prerequisites ===" -Color Cyan
    
    # Check if Microsoft.Graph module is installed
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-ColorOutput "ERROR: Microsoft.Graph.Authentication module not found!" -Color Red
        Write-ColorOutput "Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -Color Yellow
        return $false
    }
    
    Write-ColorOutput "✓ Microsoft.Graph.Authentication module found" -Color Green
    return $true
}

function Connect-ToGraph {
    Write-ColorOutput "`n=== Connecting to Microsoft Graph ===" -Color Cyan
    
    try {
        # Connect with required scopes
        Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" -NoWelcome
        
        $context = Get-MgContext
        Write-ColorOutput "✓ Connected to Microsoft Graph" -Color Green
        Write-ColorOutput "  Account: $($context.Account)" -Color Gray
        Write-ColorOutput "  Tenant: $($context.TenantId)" -Color Gray
        
        return $true
    }
    catch {
        Write-ColorOutput "ERROR: Failed to connect to Microsoft Graph" -Color Red
        Write-ColorOutput $_.Exception.Message -Color Red
        return $false
    }
}

function Get-RequiredPermissions {
    param([string[]]$RunbookSelection)
    
    # Define permissions per runbook
    $permissionMap = @{
        StaleDevices = @('Device.Read.All', 'User.Read.All', 'Mail.Send')
        AppSecrets = @('Application.Read.All', 'Mail.Send')
        MFACompliance = @('User.Read.All', 'UserAuthenticationMethod.Read.All', 'Directory.Read.All', 'Mail.Send')
        LicenseUsage = @('User.Read.All', 'Organization.Read.All', 'Mail.Send')
        AdminChanges = @('AuditLog.Read.All', 'Directory.Read.All', 'Mail.Send')
    }
    
    $allPermissions = @()
    
    if ($RunbookSelection -contains 'All') {
        # Get all unique permissions
        foreach ($perms in $permissionMap.Values) {
            $allPermissions += $perms
        }
    }
    else {
        # Get permissions for selected runbooks
        foreach ($runbook in $RunbookSelection) {
            if ($permissionMap.ContainsKey($runbook)) {
                $allPermissions += $permissionMap[$runbook]
            }
        }
    }
    
    # Return unique permissions
    return $allPermissions | Select-Object -Unique | Sort-Object
}

function Grant-GraphPermissions {
    param(
        [string]$ManagedIdentityObjectId,
        [string[]]$Permissions
    )
    
    Write-ColorOutput "`n=== Granting Microsoft Graph Permissions ===" -Color Cyan
    Write-ColorOutput "Managed Identity Object ID: $ManagedIdentityObjectId" -Color Gray
    Write-ColorOutput "Permissions to grant: $($Permissions.Count)" -Color Gray
    
    try {
        # Get Microsoft Graph Service Principal
        Write-ColorOutput "`nFetching Microsoft Graph service principal..." -Color Gray
        $graphSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'" -ErrorAction Stop
        
        if (-not $graphSp) {
            Write-ColorOutput "ERROR: Microsoft Graph service principal not found!" -Color Red
            return $false
        }
        
        Write-ColorOutput "✓ Found Microsoft Graph service principal" -Color Green
        
        # Get the Managed Identity Service Principal
        Write-ColorOutput "`nFetching Managed Identity service principal..." -Color Gray
        $managedIdentitySp = Get-MgServicePrincipal -ServicePrincipalId $ManagedIdentityObjectId -ErrorAction Stop
        
        if (-not $managedIdentitySp) {
            Write-ColorOutput "ERROR: Managed Identity not found with Object ID: $ManagedIdentityObjectId" -Color Red
            return $false
        }
        
        Write-ColorOutput "✓ Found Managed Identity: $($managedIdentitySp.DisplayName)" -Color Green
        
        # Get existing permissions
        Write-ColorOutput "`nChecking existing permissions..." -Color Gray
        $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityObjectId
        $existingAppRoleIds = $existingAssignments | Where-Object {$_.ResourceId -eq $graphSp.Id} | Select-Object -ExpandProperty AppRoleId
        
        Write-ColorOutput "Found $($existingAppRoleIds.Count) existing Graph API permissions" -Color Gray
        
        # Grant each permission
        Write-ColorOutput "`nGranting permissions:" -Color Yellow
        $successCount = 0
        $skippedCount = 0
        $failedCount = 0
        
        foreach ($permission in $Permissions) {
            # Find the app role
            $appRole = $graphSp.AppRoles | Where-Object {$_.Value -eq $permission}
            
            if (-not $appRole) {
                Write-ColorOutput "  ✗ $permission - Not found in Graph API" -Color Red
                $failedCount++
                continue
            }
            
            # Check if already assigned
            if ($existingAppRoleIds -contains $appRole.Id) {
                Write-ColorOutput "  ○ $permission - Already granted" -Color Gray
                $skippedCount++
                continue
            }
            
            # Grant the permission
            try {
                $body = @{
                    principalId = $ManagedIdentityObjectId
                    resourceId = $graphSp.Id
                    appRoleId = $appRole.Id
                }
                
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityObjectId -BodyParameter $body -ErrorAction Stop | Out-Null
                
                Write-ColorOutput "  ✓ $permission - Granted successfully" -Color Green
                $successCount++
            }
            catch {
                Write-ColorOutput "  ✗ $permission - Failed: $($_.Exception.Message)" -Color Red
                $failedCount++
            }
        }
        
        # Summary
        Write-ColorOutput "`n=== Summary ===" -Color Cyan
        Write-ColorOutput "Successfully granted: $successCount" -Color Green
        Write-ColorOutput "Already granted (skipped): $skippedCount" -Color Gray
        Write-ColorOutput "Failed: $failedCount" -Color $(if ($failedCount -gt 0) { 'Red' } else { 'Gray' })
        
        return ($failedCount -eq 0)
    }
    catch {
        Write-ColorOutput "ERROR: An unexpected error occurred" -Color Red
        Write-ColorOutput $_.Exception.Message -Color Red
        return $false
    }
}

function Show-Summary {
    param(
        [string]$ManagedIdentityObjectId,
        [string[]]$Permissions,
        [bool]$Success
    )
    
    Write-ColorOutput "`n`n╔════════════════════════════════════════════════════════════════╗" -Color Cyan
    Write-ColorOutput "║          GRAPH API PERMISSIONS CONFIGURATION COMPLETE          ║" -Color Cyan
    Write-ColorOutput "╚════════════════════════════════════════════════════════════════╝" -Color Cyan
    
    if ($Success) {
        Write-ColorOutput "`n✓ All permissions have been configured successfully!" -Color Green
        Write-ColorOutput "`nNext Steps:" -Color Yellow
        Write-ColorOutput "  1. Import the runbooks into your Azure Automation account" -Color White
        Write-ColorOutput "  2. Update the configuration section in each runbook" -Color White
        Write-ColorOutput "  3. Install required PowerShell modules in Automation account" -Color White
        Write-ColorOutput "  4. Test each runbook manually before scheduling" -Color White
        Write-ColorOutput "  5. Configure schedules for automated execution" -Color White
    }
    else {
        Write-ColorOutput "`n⚠ Some permissions could not be granted" -Color Yellow
        Write-ColorOutput "`nPlease review the errors above and:" -Color Yellow
        Write-ColorOutput "  1. Ensure you have sufficient Azure AD permissions" -Color White
        Write-ColorOutput "  2. Verify the Managed Identity Object ID is correct" -Color White
        Write-ColorOutput "  3. Check for any connectivity issues" -Color White
        Write-ColorOutput "  4. Re-run this script after resolving issues" -Color White
    }
    
    Write-ColorOutput "`n═══════════════════════════════════════════════════════════════════" -Color Cyan
}
#endregion

#region Main Script
try {
    Write-ColorOutput @"

╔════════════════════════════════════════════════════════════════╗
║     Azure Automation Graph API Permissions Setup Script       ║
║                                                                ║
║  This script will grant Microsoft Graph API permissions to    ║
║  your Azure Automation Managed Identity for the runbooks.     ║
╚════════════════════════════════════════════════════════════════╝

"@ -Color Cyan

    # Check prerequisites
    if (-not (Test-Prerequisites)) {
        exit 1
    }
    
    # Connect to Microsoft Graph
    if (-not (Connect-ToGraph)) {
        exit 1
    }
    
    # Get required permissions based on runbook selection
    $requiredPermissions = Get-RequiredPermissions -RunbookSelection $RunbookSelection
    
    Write-ColorOutput "`n=== Configuration Summary ===" -Color Cyan
    Write-ColorOutput "Runbook Selection: $($RunbookSelection -join ', ')" -Color White
    Write-ColorOutput "Permissions Required: $($requiredPermissions.Count)" -Color White
    Write-ColorOutput "`nPermissions to be granted:" -Color Yellow
    $requiredPermissions | ForEach-Object { Write-ColorOutput "  • $_" -Color White }
    
    # Prompt for confirmation
    Write-Host "`n"
    $confirmation = Read-Host "Do you want to proceed? (Y/N)"
    
    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
        Write-ColorOutput "`nOperation cancelled by user." -Color Yellow
        exit 0
    }
    
    # Grant permissions
    $success = Grant-GraphPermissions -ManagedIdentityObjectId $ManagedIdentityObjectId -Permissions $requiredPermissions
    
    # Show summary
    Show-Summary -ManagedIdentityObjectId $ManagedIdentityObjectId -Permissions $requiredPermissions -Success $success
    
    # Disconnect
    Disconnect-MgGraph | Out-Null
    
    exit $(if ($success) { 0 } else { 1 })
}
catch {
    Write-ColorOutput "`nFATAL ERROR: $($_.Exception.Message)" -Color Red
    Write-ColorOutput $_.ScriptStackTrace -Color Red
    exit 1
}
#endregion
