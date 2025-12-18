# Configuration Template for Azure Automation Runbooks
# Copy this file and customize for your environment
# Rename to: runbook-config.ps1

<#
.SYNOPSIS
    Central configuration file for all Azure Automation runbooks.

.DESCRIPTION
    This file contains all configurable parameters for the runbook collection.
    Update the values below to match your organization's requirements.

.NOTES
    This is a template file. Copy and customize for your environment.
#>

#region Global Email Settings
# Email addresses for notifications
$Global:EmailRecipients = @{
    # Primary administrators
    Admin     = @("admin@yourdomain.com")
    
    # Security team
    Security  = @("security@yourdomain.com", "compliance@yourdomain.com")
    
    # Finance/Cost optimization
    Finance   = @("finance@yourdomain.com", "procurement@yourdomain.com")
    
    # DevOps/Operations
    DevOps    = @("devops@yourdomain.com", "operations@yourdomain.com")
    
    # Executive leadership (for critical alerts only)
    Executive = @("cio@yourdomain.com", "ciso@yourdomain.com")
}

# Sender address (must be a valid mailbox in your tenant)
$Global:EmailFrom = "azureautomation@yourdomain.com"

# Email subject prefix (helps with filtering/routing)
$Global:EmailSubjectPrefix = "[Azure Automation]"
#endregion

#region Threshold Settings

# Stale Devices and Accounts
$Global:StaleThresholds = @{
    DeviceInactivity = 90   # Days before a device is considered stale
    UserInactivity   = 90   # Days before a user is considered stale
}

# App Registration Secrets
$Global:SecretExpirationThresholds = @{
    Critical = 7    # Expires in 7 days or less - URGENT
    Warning  = 30   # Expires in 30 days or less - Plan rotation
    Info     = 90   # Expires in 90 days or less - Advance notice
}

# License Usage
$Global:LicenseThresholds = @{
    InactivityDays = 90     # Days of inactivity before license is flagged as wasted
}

# Admin Role Changes
$Global:AdminMonitoringSettings = @{
    LookbackDays = 7        # Number of days to query in audit logs
}
#endregion

#region License Cost Configuration
# Update these values with your actual Microsoft 365 license costs (per user/month)
# This is used for cost savings calculations in the License Usage runbook

$Global:LicenseCosts = @{
    # Microsoft 365 Plans
    'Microsoft 365 E3'                  = 36.00
    'Microsoft 365 E5'                  = 57.00
    'Microsoft 365 F3'                  = 8.00
    'Microsoft 365 Business Basic'      = 6.00
    'Microsoft 365 Business Standard'   = 12.50
    'Microsoft 365 Business Premium'    = 22.00
    
    # Office 365 Plans
    'Office 365 E1'                     = 8.00
    'Office 365 E3'                     = 23.00
    'Office 365 E5'                     = 38.00
    
    # Enterprise Mobility + Security
    'Enterprise Mobility + Security E3' = 10.60
    'Enterprise Mobility + Security E5' = 16.40
    
    # Add-ons and Standalone
    'Power BI Pro'                      = 9.99
    'Power BI Premium Per User'         = 20.00
    'Project Plan 1'                    = 10.00
    'Project Plan 3'                    = 30.00
    'Project Plan 5'                    = 55.00
    'Visio Plan 1'                      = 5.00
    'Visio Plan 2'                      = 15.00
    'Microsoft Defender for Office 365' = 2.00
    'Azure AD Premium P1'               = 6.00
    'Azure AD Premium P2'               = 9.00
    'Microsoft Intune'                  = 6.00
    'Exchange Online Plan 1'            = 4.00
    'Exchange Online Plan 2'            = 8.00
    'SharePoint Online Plan 1'          = 5.00
    'SharePoint Online Plan 2'          = 10.00
    
    # Add your custom licenses here as needed
}
#endregion

#region High Priority Admin Roles
# These roles are flagged as high-priority in admin monitoring
# Changes to these roles trigger special alerts

$Global:HighPriorityRoles = @(
    'Global Administrator'
    'Privileged Role Administrator'
    'Security Administrator'
    'Conditional Access Administrator'
    'Exchange Administrator'
    'SharePoint Administrator'
    'User Administrator'
    'Billing Administrator'
    'Authentication Administrator'
    'Privileged Authentication Administrator'
    'Cloud Application Administrator'
    'Application Administrator'
    'Azure AD Joined Device Local Administrator'
    'Compliance Administrator'
    'Security Operator'
    'Password Administrator'
    'Helpdesk Administrator'
)
#endregion

#region Exclusion Filters

# Stale Accounts Exclusions
$Global:StaleAccountsExclusions = @{
    # Exclude service accounts (accounts containing these patterns)
    ServiceAccountPatterns = @('svc', 'service', 'admin-', 'sa-', 'system')
    
    # Exclude disabled accounts from stale user reports
    ExcludeDisabledUsers = $true
    
    # Exclude guest users from stale user reports
    ExcludeGuestUsers = $true
}

# MFA Compliance Exclusions
$Global:MFAComplianceExclusions = @{
    # Exclude service accounts from MFA compliance checks
    ExcludeServiceAccounts = $true
    
    # Exclude disabled accounts from MFA compliance checks
    ExcludeDisabledUsers = $true
    
    # Exclude guest users from MFA compliance checks
    ExcludeGuestUsers = $false
}
#endregion

#region Report Settings

# Report generation settings
$Global:ReportSettings = @{
    # Where to store temporary report files
    ExportPath = $env:TEMP
    
    # Date format for report filenames
    DateFormat = 'yyyyMMdd_HHmmss'
    
    # Include summary dashboard in all reports
    IncludeSummary = $true
    
    # Maximum number of "Top N" items in priority lists
    TopItemsCount = 20
}

# Excel formatting preferences
$Global:ExcelFormatting = @{
    # Apply conditional formatting (color coding)
    UseConditionalFormatting = $true
    
    # Auto-size columns for readability
    AutoSize = $true
    
    # Include auto-filter on data columns
    AutoFilter = $true
    
    # Bold top row (headers)
    BoldTopRow = $true
    
    # Freeze top row for scrolling
    FreezeTopRow = $true
}
#endregion

#region Notification Preferences

# Email notification settings per runbook
$Global:NotificationPreferences = @{
    StaleDevicesAndAccounts = @{
        Recipients = $Global:EmailRecipients.Admin + $Global:EmailRecipients.Security
        SendOnlyIfFindings = $false  # Always send, even if no stale items found
        IncludeExecutiveSummary = $true
    }
    
    AppRegistrationSecrets = @{
        Recipients = $Global:EmailRecipients.DevOps + $Global:EmailRecipients.Security
        SendOnlyIfFindings = $false  # Always send for awareness
        IncludeExecutiveSummary = $true
        # Alert executive team if critical secrets are expired
        AlertExecutiveOnCritical = $true
    }
    
    MFACompliance = @{
        Recipients = $Global:EmailRecipients.Security
        SendOnlyIfFindings = $false  # Always send for compliance tracking
        IncludeExecutiveSummary = $true
        # Alert executive team if admins without MFA detected
        AlertExecutiveOnAdminNonCompliance = $true
    }
    
    LicenseUsage = @{
        Recipients = $Global:EmailRecipients.Finance + $Global:EmailRecipients.Admin
        SendOnlyIfFindings = $true  # Only send if waste detected
        IncludeExecutiveSummary = $true
        MinimumSavingsToAlert = 1000  # Only alert if annual savings >= $1000
    }
    
    AdminRoleChanges = @{
        Recipients = $Global:EmailRecipients.Security + $Global:EmailRecipients.Executive
        SendOnlyIfFindings = $false  # Always send for audit trail
        IncludeExecutiveSummary = $true
    }
}
#endregion

#region Environment Settings

# Azure environment (AzureCloud, AzureUSGovernment, etc.)
$Global:AzureEnvironment = 'AzureCloud'

# Graph API endpoints (change if using sovereign clouds)
$Global:GraphEndpoints = @{
    Commercial = 'https://graph.microsoft.com'
    USGovernment = 'https://graph.microsoft.us'
    China = 'https://microsoftgraph.chinacloudapi.cn'
    Germany = 'https://graph.microsoft.de'
}

# Current Graph API endpoint (set based on your environment)
$Global:GraphApiEndpoint = $Global:GraphEndpoints.Commercial
#endregion

#region Advanced Settings

# Rate limiting and throttling
$Global:RateLimiting = @{
    # Delay between API calls (milliseconds)
    StandardDelay = 100
    
    # Delay between batch operations (milliseconds)
    BatchDelay = 500
    
    # Maximum retry attempts for failed API calls
    MaxRetries = 3
    
    # Exponential backoff multiplier
    BackoffMultiplier = 2
}

# Logging preferences
$Global:LoggingSettings = @{
    # Log level (INFO, WARNING, ERROR)
    LogLevel = 'INFO'
    
    # Include timestamps in logs
    IncludeTimestamp = $true
    
    # Log format
    TimestampFormat = 'yyyy-MM-dd HH:mm:ss'
}
#endregion

#region Validation Function
function Test-Configuration {
    <#
    .SYNOPSIS
        Validates the configuration settings.
    
    .DESCRIPTION
        Checks that all required settings are properly configured.
    #>
    
    $isValid = $true
    $issues = @()
    
    # Validate email recipients
    if ($Global:EmailRecipients.Admin.Count -eq 0) {
        $isValid = $false
        $issues += "No admin email recipients configured"
    }
    
    # Validate email from address
    if ([string]::IsNullOrWhiteSpace($Global:EmailFrom)) {
        $isValid = $false
        $issues += "EmailFrom address not configured"
    }
    
    # Validate license costs
    if ($Global:LicenseCosts.Count -eq 0) {
        $isValid = $false
        $issues += "No license costs configured"
    }
    
    # Validate high priority roles
    if ($Global:HighPriorityRoles.Count -eq 0) {
        $isValid = $false
        $issues += "No high priority roles configured"
    }
    
    if (-not $isValid) {
        Write-Warning "Configuration validation failed:"
        $issues | ForEach-Object { Write-Warning "  - $_" }
        return $false
    }
    
    Write-Output "Configuration validation passed"
    return $true
}
#endregion

# Export all settings for use in runbooks
Export-ModuleMember -Variable * -Function Test-Configuration
