# Quick Start Guide

Get up and running with Azure Automation runbooks in 15 minutes.

## Prerequisites Checklist

- [ ] Azure subscription with Automation Account
- [ ] Global Administrator or privileged admin access
- [ ] PowerShell 7+ installed locally (for setup script)
- [ ] Microsoft.Graph PowerShell module installed

## 5-Minute Setup

### 1. Enable Managed Identity (2 minutes)

```powershell
# Via Azure Portal
1. Go to your Automation Account
2. Click "Identity" in left menu
3. Turn on "System assigned" identity
4. Save and copy the Object ID
```

### 2. Grant Permissions (3 minutes)

Download and run the permissions setup script:

```powershell
# Install Graph module if needed
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

# Run the setup script
.\Grant-GraphPermissions.ps1 -ManagedIdentityObjectId "YOUR-OBJECT-ID"
```

**Pro Tip**: If you only need specific runbooks, use the `-RunbookSelection` parameter:
```powershell
.\Grant-GraphPermissions.ps1 -ManagedIdentityObjectId "YOUR-ID" -RunbookSelection "StaleDevices","MFACompliance"
```

### 3. Install Modules (5 minutes)

In your Automation Account, go to **Modules** → **Add a module** → **Browse from gallery**

Install these modules in order:
1. `Microsoft.Graph.Authentication` (v2.0.0+)
2. `Microsoft.Graph.Users`
3. `Microsoft.Graph.Applications`
4. `Microsoft.Graph.Reports`
5. `Microsoft.Graph.DeviceManagement`
6. `ImportExcel`

Wait for each to finish installing before starting the next.

### 4. Import Runbooks (3 minutes)

1. Go to **Runbooks** → **Import a runbook**
2. Upload each `.ps1` file
3. Click **Publish** after upload

### 5. Configure Email Settings (2 minutes)

Edit each runbook and update the configuration section:

```powershell
$EmailRecipients = @("your-email@domain.com")
$EmailFrom = "automation@yourdomain.com"
```

## Test Your First Runbook

Let's test the MFA Compliance runbook:

1. Open `Monitor-MFACompliance` runbook
2. Click **Start**
3. Wait 2-5 minutes for completion
4. Check **Output** for results
5. Check your email for the report

## What You Should See

### Successful Run Output:
```
[2024-12-17 10:30:00] [INFO] ========== Starting MFA Compliance Monitoring ==========
[2024-12-17 10:30:01] [INFO] Connecting to Microsoft Graph using Managed Identity...
[2024-12-17 10:30:02] [INFO] Successfully connected to Microsoft Graph
[2024-12-17 10:30:02] [INFO] Retrieving admin role members...
[2024-12-17 10:35:45] [INFO] Excel report created successfully
[2024-12-17 10:35:50] [INFO] Email sent successfully
[2024-12-17 10:35:51] [INFO] ========== Runbook Completed Successfully ==========
SUCCESS: MFA Compliance Rate: 87.5%. Admins without MFA: 3, Non-admins without MFA: 45
```

### Email You'll Receive:
- Subject: "MFA Compliance Report - 2024-12-17"
- Excel attachment with multiple sheets
- Summary dashboard with key metrics
- Detailed user lists

## Schedule Automation

Once tested, set up schedules:

1. Click **Schedules** → **Add a schedule**
2. Choose frequency:
   - **Daily**: App Secrets, Admin Changes
   - **Weekly**: Stale Devices, MFA Compliance  
   - **Monthly**: License Usage

**Recommended Schedule Times**:
- 6:00 AM local time (before business hours)
- Stagger different runbooks by 30 minutes to avoid throttling

## Troubleshooting Quick Fixes

### "Access Denied" Error
```powershell
# Re-run permissions script
.\Grant-GraphPermissions.ps1 -ManagedIdentityObjectId "YOUR-ID"
```

### "Module Not Found" Error
```
Wait 10 minutes after module installation, then retry
```

### Email Not Sending
```powershell
# Verify email address in Azure AD:
Get-MgUser -UserId "automation@yourdomain.com"

# Address must exist and be enabled
```

### Rate Limiting
```powershell
# In runbook configuration, increase sleep intervals:
Start-Sleep -Milliseconds 200  # Increase from 100
```

## Quick Customization

### Change Thresholds

**Stale devices** (default 90 days):
```powershell
$StaleDeviceThreshold = 60  # Change to 60 days
```

**Secret expiration** (default 7 days critical):
```powershell
$CriticalThreshold = 14  # Change to 14 days
```

### Add More Recipients

```powershell
$EmailRecipients = @(
    "admin@domain.com",
    "security@domain.com",
    "team@domain.com"
)
```

### Update License Costs

In `Analyze-LicenseUsage.ps1`:
```powershell
$LicenseCosts = @{
    'Microsoft 365 E3' = 32.00  # Your negotiated price
    'Microsoft 365 E5' = 50.00  # Your negotiated price
}
```

## Next Steps

Now that everything is working:

1. **Review Reports**: Check the Excel files for data accuracy
2. **Adjust Thresholds**: Fine-tune based on your needs  
3. **Add Recipients**: Include relevant stakeholders
4. **Set Schedules**: Automate with appropriate frequencies
5. **Document Processes**: Create internal runbook documentation

## Getting Help

**Module Issues**: 
- Check versions match requirements (v2.0.0+)
- Modules must finish installing completely

**Permission Issues**:
- Verify Managed Identity is enabled
- Confirm permissions with: `Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId "YOUR-ID"`

**Graph API Throttling**:
- Increase sleep intervals in runbook
- Stagger schedule times
- Process fewer items per run

## Advanced: Configuration File

For centralized configuration, use the template:

```powershell
# Copy template
Copy-Item runbook-config-template.ps1 runbook-config.ps1

# Edit runbook-config.ps1 with your settings

# In each runbook, dot-source the config at top:
. .\runbook-config.ps1
```

## Success Metrics

After 1 week, you should see:

- ✅ All runbooks executing successfully on schedule
- ✅ Email reports arriving consistently  
- ✅ Zero permission errors in logs
- ✅ Stakeholders reviewing and acting on reports
- ✅ Measurable improvements (e.g., fewer stale accounts)

## Maintenance Checklist

**Monthly**:
- [ ] Review and update email recipients
- [ ] Adjust thresholds based on results
- [ ] Update license costs if pricing changes

**Quarterly**:
- [ ] Review Graph API permissions (remove unused)
- [ ] Update PowerShell modules to latest versions
- [ ] Audit runbook effectiveness with stakeholders

**Annually**:
- [ ] Review all configurations
- [ ] Update documentation
- [ ] Reassess which runbooks are still needed

---

## You're Ready! 🚀

Your Azure Automation runbooks are now set up and ready to provide continuous monitoring, compliance tracking, and cost optimization for your Microsoft 365 environment.

**Questions?** Check the main [README.md](README.md) for detailed documentation.
