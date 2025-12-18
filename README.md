# Azure Automation Runbook Collection

A comprehensive collection of enterprise-grade Azure Automation runbooks for monitoring, compliance, and cost optimization using Microsoft Graph API and Managed Identity authentication.

## 🌟 Features

- **Security-First Architecture**: All runbooks use Managed Identity authentication (no stored credentials)
- **Professional Reporting**: Excel reports with conditional formatting and email distribution
- **Enterprise-Grade**: Comprehensive error handling, rate limiting, and logging
- **Cost Optimization**: Zero additional cost - leverages existing Microsoft 365 capabilities
- **Government Cloud Compatible**: Works with Azure Government Cloud (GCC) environments

## 📋 Runbooks

### 1. Monitor-StaleDevicesAndAccounts.ps1
Identifies inactive devices and user accounts based on configurable thresholds.

**Key Features:**
- Monitors both devices and user accounts
- Configurable inactivity thresholds
- Conditional formatting in Excel (green/red status indicators)
- Separate sheets for devices and users

**Graph API Permissions Required:**
- `Device.Read.All`
- `User.Read.All`
- `Mail.Send`

**Configuration:**
```powershell
$StaleDeviceThreshold = 90  # days
$StaleUserThreshold = 90    # days
$EmailRecipients = @("admin@yourdomain.com")
```

---

### 2. Monitor-AppRegistrationSecrets.ps1
Tracks app registration secrets and certificates for expiration.

**Key Features:**
- Monitors both secrets and certificates
- Multi-tier alerting (Expired, Critical, Warning, Info)
- Top 10 most urgent credentials report
- Separate sheets by status for easy triage

**Graph API Permissions Required:**
- `Application.Read.All`
- `Mail.Send`

**Configuration:**
```powershell
$CriticalThreshold = 7   # days
$WarningThreshold = 30   # days
$InfoThreshold = 90      # days
$EmailRecipients = @("admin@yourdomain.com")
```

---

### 3. Monitor-MFACompliance.ps1
Analyzes Multi-Factor Authentication compliance across the organization.

**Key Features:**
- Separates admin vs non-admin users
- Identifies MFA methods in use
- Compliance rate calculations
- Prioritizes admin accounts without MFA

**Graph API Permissions Required:**
- `User.Read.All`
- `UserAuthenticationMethod.Read.All`
- `Directory.Read.All`
- `Mail.Send`

**Configuration:**
```powershell
$ExcludeServiceAccounts = $true
$ExcludeDisabledUsers = $true
$EmailRecipients = @("security@yourdomain.com")
```

---

### 4. Analyze-LicenseUsage.ps1
Analyzes Microsoft 365 license usage and identifies optimization opportunities.

**Key Features:**
- Identifies unused/underutilized licenses
- Cost savings calculations (monthly/annual)
- Breakdown by license type
- Top 20 wasted licenses report

**Graph API Permissions Required:**
- `User.Read.All`
- `Organization.Read.All`
- `Mail.Send`

**Configuration:**
```powershell
$InactivityThreshold = 90  # days

# Update with your actual license costs
$LicenseCosts = @{
    'Microsoft 365 E3' = 36.00
    'Microsoft 365 E5' = 57.00
    # Add more as needed...
}
```

---

### 5. Monitor-AdminRoleChanges.ps1
Tracks administrative role assignments and removals for governance.

**Key Features:**
- Audit log analysis for role changes
- High-priority role identification
- Current admin snapshot
- Changes by role breakdown

**Graph API Permissions Required:**
- `AuditLog.Read.All`
- `Directory.Read.All`
- `Mail.Send`

**Configuration:**
```powershell
$DaysToQuery = 7  # Look back period

$HighPriorityRoles = @(
    'Global Administrator',
    'Privileged Role Administrator',
    'Security Administrator'
    # Add more as needed...
)
```

---

## 🚀 Setup Instructions

### Prerequisites

1. **Azure Automation Account** with System-Assigned Managed Identity enabled
2. **Microsoft Graph PowerShell Modules** installed in Azure Automation:
   - `Microsoft.Graph.Authentication` (v2.0.0+)
   - `Microsoft.Graph.Users` (v2.0.0+)
   - `Microsoft.Graph.Applications` (v2.0.0+)
   - `Microsoft.Graph.Reports` (v2.0.0+)
   - `Microsoft.Graph.DeviceManagement` (v2.0.0+)
   - `ImportExcel`

### Step 1: Enable Managed Identity

1. Navigate to your Azure Automation Account
2. Go to **Identity** → **System assigned**
3. Set Status to **On**
4. Save and note the Object ID

### Step 2: Grant Graph API Permissions

Use Azure CLI or PowerShell to grant the required permissions:

```powershell
# Connect to Azure AD
Connect-MgGraph -Scopes "Application.ReadWrite.All"

# Get your Managed Identity
$managedIdentityObjectId = "YOUR-MANAGED-IDENTITY-OBJECT-ID"

# Get Microsoft Graph Service Principal
$graphApp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'"

# Define required permissions based on runbooks you'll use
$permissions = @(
    "Device.Read.All",
    "User.Read.All",
    "Application.Read.All",
    "AuditLog.Read.All",
    "Directory.Read.All",
    "UserAuthenticationMethod.Read.All",
    "Organization.Read.All",
    "Mail.Send"
)

# Grant each permission
foreach ($permission in $permissions) {
    $appRole = $graphApp.AppRoles | Where-Object {$_.Value -eq $permission}
    
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $managedIdentityObjectId `
        -PrincipalId $managedIdentityObjectId `
        -ResourceId $graphApp.Id `
        -AppRoleId $appRole.Id
}
```

### Step 3: Import Runbooks

1. Navigate to **Automation Accounts** → Your Account → **Runbooks**
2. Click **+ Create a runbook** or **Import a runbook**
3. Upload each `.ps1` file
4. Publish each runbook after importing

### Step 4: Install Required Modules

1. Navigate to **Modules** → **Add a module**
2. Browse Gallery and install:
   - `Microsoft.Graph.Authentication`
   - `Microsoft.Graph.Users`
   - `Microsoft.Graph.Applications`
   - `Microsoft.Graph.Reports`
   - `Microsoft.Graph.DeviceManagement`
   - `ImportExcel`

### Step 5: Configure Email Settings

Update each runbook's configuration section:

```powershell
$EmailRecipients = @("your-email@domain.com")
$EmailFrom = "automation@yourdomain.com"
```

**Note**: The `$EmailFrom` address must be a valid mailbox in your tenant that the Managed Identity has `Mail.Send` permission for.

### Step 6: Schedule Runbooks

1. Navigate to each runbook
2. Click **Schedules** → **+ Add a schedule**
3. Recommended schedules:
   - Stale Devices/Accounts: Weekly (Monday morning)
   - App Registration Secrets: Daily
   - MFA Compliance: Weekly
   - License Usage: Monthly
   - Admin Role Changes: Daily

---

## 📊 Report Features

All runbooks generate Excel reports with:

- **Executive Summary Sheet**: Key metrics at a glance
- **Conditional Formatting**: Color-coded status indicators
- **Multiple Sheets**: Organized data for different audiences
- **Professional Styling**: Bold headers, frozen panes, auto-sizing
- **Sortable/Filterable**: Built-in Excel AutoFilter

### Example Report Structure

```
📊 Excel Report
├─ Summary (Executive metrics)
├─ High Priority Items (Top concerns)
├─ Detailed Data (Complete dataset)
└─ Analysis by Category (Groupings/trends)
```

---

## 🔒 Security Best Practices

1. **Managed Identity Only**: Never store credentials in variables or Automation Account
2. **Least Privilege**: Grant only the minimum required Graph API permissions
3. **Rate Limiting**: All runbooks include exponential backoff for API throttling
4. **Error Handling**: Comprehensive try-catch blocks with detailed logging
5. **Audit Trails**: All actions logged with timestamps and severity levels

---

## 🛠️ Customization

### Email Templates

Each runbook includes HTML email templates. Customize the styling in the `Send-EmailWithAttachment` function:

```powershell
$htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        /* Your custom CSS here */
    </style>
</head>
<body>
    <!-- Your custom content here -->
</body>
</html>
"@
```

### Thresholds and Filters

Adjust monitoring thresholds in the configuration section:

```powershell
#region Configuration
$YourThreshold = 90
$YourFilter = $true
# ... more settings
#endregion
```

### Report Formatting

Modify Excel conditional formatting in `Export-ToExcelWithFormatting`:

```powershell
-ConditionalText $(
    New-ConditionalText -Text "YOUR_VALUE" -Range "A:A" -BackgroundColor Red
)
```

---

## 📝 Logging

All runbooks use consistent logging:

```powershell
Write-Log "Message" -Level INFO    # Informational
Write-Log "Message" -Level WARNING # Warnings
Write-Log "Message" -Level ERROR   # Errors
```

Logs are output to Azure Automation job streams and can be viewed in:
- **Jobs** → Select Job → **All Logs**

---

## 🐛 Troubleshooting

### Common Issues

**Issue**: "Access Denied" or "Insufficient privileges"
- **Solution**: Verify Managed Identity has required Graph API permissions
- Check permission consent in Azure AD

**Issue**: "Module not found"
- **Solution**: Verify all required modules are installed in Azure Automation
- Wait for module import to complete (can take 5-10 minutes)

**Issue**: Email not sending
- **Solution**: Verify `$EmailFrom` is a valid mailbox
- Confirm Managed Identity has `Mail.Send` permission
- Check mailbox is not disabled or on hold

**Issue**: Rate limiting errors
- **Solution**: Increase `Start-Sleep` values in API call loops
- Consider spreading out schedule times if running multiple runbooks

### Debug Mode

Add verbose logging temporarily:

```powershell
$VerbosePreference = 'Continue'
Write-Verbose "Debug information here"
```

---

## 📈 Performance Considerations

- **Large Tenants** (10,000+ users): Consider increasing sleep intervals between API calls
- **Government Cloud**: May have lower API throttling limits - adjust accordingly
- **Concurrent Runbooks**: Stagger schedules to avoid hitting tenant-wide limits
- **Report Size**: For very large datasets, consider splitting into multiple worksheets

---

## 🤝 Contributing

Feel free to:
- Submit issues for bugs or feature requests
- Create pull requests with improvements
- Share customizations that others might find useful

---

## 📄 License

This collection is provided as-is for use in your organization. Modify and adapt as needed for your requirements.

---

## 🔗 Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/graph/)
- [Azure Automation Documentation](https://docs.microsoft.com/azure/automation/)
- [ImportExcel Module](https://github.com/dfinke/ImportExcel)
- [Azure Managed Identities](https://docs.microsoft.com/azure/active-directory/managed-identities-azure-resources/)

---

## ⚠️ Disclaimer

These runbooks are provided as examples and should be tested in a non-production environment before deployment. Always review and customize based on your organization's specific requirements and compliance needs.

---

**Version**: 4.0  
**Last Updated**: December 2025  
**Maintained By**: Jaimin
