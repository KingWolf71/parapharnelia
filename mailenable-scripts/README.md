# MailEnable User Management Scripts

A collection of PowerShell scripts for managing MailEnable mailboxes, quotas, passwords, and user accounts.

## Prerequisites

- Windows Server with MailEnable installed
- PowerShell 5.1 or higher
- Administrative privileges on the MailEnable server
- MailEnable API/COM objects properly configured

## Scripts Overview

### 1. Add-MEMailbox.ps1
Creates new mailboxes in MailEnable.

**Usage:**
```powershell
# Basic usage
.\Add-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "john.doe" -Password "P@ssw0rd123"

# With full details
.\Add-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "jane.smith" `
    -Password "SecureP@ss" -FirstName "Jane" -LastName "Smith" -Quota 500
```

**Parameters:**
- `PostOffice` (required): Domain name
- `Mailbox` (required): Username
- `Password` (required): Mailbox password
- `FirstName` (optional): User's first name
- `LastName` (optional): User's last name
- `Quota` (optional): Mailbox quota in MB (default: 100)

---

### 2. Remove-MEMailbox.ps1
Removes mailboxes from MailEnable with optional backup.

**Usage:**
```powershell
# Basic removal (with confirmation)
.\Remove-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "old.user"

# With backup before deletion
.\Remove-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "archive.user" `
    -BackupFirst -BackupPath "C:\Backups\MailEnable"

# Force removal without confirmation
.\Remove-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "spam.account" -Force
```

**Parameters:**
- `PostOffice` (required): Domain name
- `Mailbox` (required): Username to remove
- `BackupFirst` (optional): Backup mailbox before deletion
- `BackupPath` (optional): Backup destination path
- `Force` (optional): Skip confirmation prompt

---

### 3. Set-MEMailboxQuota.ps1
Updates mailbox storage quotas.

**Usage:**
```powershell
# Set quota to 500MB
.\Set-MEMailboxQuota.ps1 -PostOffice "example.com" -Mailbox "john.doe" -QuotaMB 500

# Set unlimited quota
.\Set-MEMailboxQuota.ps1 -PostOffice "example.com" -Mailbox "vip.user" -QuotaMB 0

# Show current quota and usage before updating
.\Set-MEMailboxQuota.ps1 -PostOffice "example.com" -Mailbox "john.doe" `
    -QuotaMB 1000 -ShowCurrent
```

**Parameters:**
- `PostOffice` (required): Domain name
- `Mailbox` (required): Username
- `QuotaMB` (required): Quota size in MB (0 = unlimited)
- `ShowCurrent` (optional): Display current quota before changing

---

### 4. Set-MEMailboxPassword.ps1
Changes or generates mailbox passwords.

**Usage:**
```powershell
# Set specific password
.\Set-MEMailboxPassword.ps1 -PostOffice "example.com" -Mailbox "john.doe" `
    -NewPassword "NewP@ssw0rd123"

# Generate random password
.\Set-MEMailboxPassword.ps1 -PostOffice "example.com" -Mailbox "john.doe" `
    -GenerateRandom -ShowPassword

# Generate longer random password
.\Set-MEMailboxPassword.ps1 -PostOffice "example.com" -Mailbox "jane.smith" `
    -GenerateRandom -PasswordLength 20 -ShowPassword
```

**Parameters:**
- `PostOffice` (required): Domain name
- `Mailbox` (required): Username
- `NewPassword` (optional): New password to set
- `GenerateRandom` (optional): Generate random secure password
- `PasswordLength` (optional): Length for generated password (default: 16)
- `ShowPassword` (optional): Display the new password

---

### 5. Get-MEMailboxes.ps1
Lists and reports on mailboxes.

**Usage:**
```powershell
# List all mailboxes (simple view)
.\Get-MEMailboxes.ps1

# List mailboxes for specific domain with details
.\Get-MEMailboxes.ps1 -PostOffice "example.com" -Detailed

# Export to CSV
.\Get-MEMailboxes.ps1 -PostOffice "example.com" -ExportCSV `
    -CSVPath "C:\Reports\mailboxes.csv"

# Include disabled mailboxes
.\Get-MEMailboxes.ps1 -PostOffice "example.com" -ShowInactive -Detailed
```

**Parameters:**
- `PostOffice` (optional): Domain name (if omitted, scans all domains)
- `Detailed` (optional): Show quota, status, and size information
- `ExportCSV` (optional): Export results to CSV
- `CSVPath` (optional): CSV file path (default: mailboxes.csv)
- `ShowInactive` (optional): Include disabled mailboxes

---

## Common Scenarios

### Bulk User Creation

Create a CSV file with user details:

**users.csv:**
```csv
PostOffice,Mailbox,Password,FirstName,LastName,Quota
example.com,john.doe,P@ssw0rd1,John,Doe,250
example.com,jane.smith,P@ssw0rd2,Jane,Smith,500
example.com,bob.jones,P@ssw0rd3,Bob,Jones,100
```

Then import and create:
```powershell
Import-Csv .\users.csv | ForEach-Object {
    .\Add-MEMailbox.ps1 -PostOffice $_.PostOffice -Mailbox $_.Mailbox `
        -Password $_.Password -FirstName $_.FirstName -LastName $_.LastName `
        -Quota $_.Quota
}
```

### Reset Multiple Passwords

```powershell
$users = @("john.doe", "jane.smith", "bob.jones")
foreach ($user in $users) {
    .\Set-MEMailboxPassword.ps1 -PostOffice "example.com" -Mailbox $user `
        -GenerateRandom -ShowPassword
}
```

### Increase Quota for All Users

```powershell
# Get all mailboxes and increase quota by 100MB
.\Get-MEMailboxes.ps1 -PostOffice "example.com" -ExportCSV -CSVPath "temp.csv"
Import-Csv .\temp.csv | ForEach-Object {
    $newQuota = $_.QuotaMB + 100
    .\Set-MEMailboxQuota.ps1 -PostOffice $_.PostOffice -Mailbox $_.Mailbox `
        -QuotaMB $newQuota
}
```

### Generate Mailbox Report

```powershell
# Create detailed report of all mailboxes
.\Get-MEMailboxes.ps1 -Detailed -ExportCSV -CSVPath "C:\Reports\mailbox-report-$(Get-Date -Format 'yyyyMMdd').csv"
```

## Security Considerations

1. **Passwords**: Always use strong passwords. The password generator creates secure random passwords.
2. **Backups**: Always backup mailboxes before deletion using the `-BackupFirst` parameter.
3. **Permissions**: Run scripts with appropriate administrative privileges.
4. **Logging**: Consider adding logging to track administrative actions.
5. **Secure Storage**: Store scripts and CSV files containing credentials in secure locations.

## Troubleshooting

### COM Object Issues
If you get COM object errors, ensure:
- MailEnable is properly installed
- You're running as Administrator
- MailEnable services are running

### Path Issues
Default MailEnable installation path is:
```
C:\Program Files (x86)\Mail Enable\
```

If your installation is elsewhere, modify the path in the scripts.

### Permission Errors
Ensure you have:
- Administrator rights on the Windows server
- Permissions to access MailEnable directories
- Appropriate MailEnable admin credentials

## MailEnable API Reference

These scripts use the MailEnable COM API:
- `MEAIPO.Mailbox` - Mailbox management
- Methods: `AddMailbox()`, `RemoveMailbox()`, `GetMailbox()`, `UpdateMailbox()`

## License

These scripts are provided as-is for use with MailEnable servers.

## Support

For MailEnable-specific issues, consult:
- MailEnable documentation
- MailEnable support forums
- Your MailEnable administrator

## Version History

- **v1.0** - Initial release with core user management functionality
  - Add mailboxes
  - Remove mailboxes
  - Manage quotas
  - Password management
  - List and report mailboxes
