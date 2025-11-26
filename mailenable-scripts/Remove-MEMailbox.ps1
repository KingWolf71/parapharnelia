<#
.SYNOPSIS
    Removes a mailbox from MailEnable.

.DESCRIPTION
    This script deletes a mailbox from MailEnable. Use with caution as this action cannot be undone.
    Optionally backup the mailbox before deletion.

.PARAMETER PostOffice
    The MailEnable post office (domain) where the mailbox exists.

.PARAMETER Mailbox
    The mailbox name to remove.

.PARAMETER BackupFirst
    If specified, backs up the mailbox before deletion.

.PARAMETER BackupPath
    Path where the backup will be stored (only used if BackupFirst is specified).

.PARAMETER Force
    Skip confirmation prompt.

.EXAMPLE
    .\Remove-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "john.doe"

.EXAMPLE
    .\Remove-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "old.user" -BackupFirst -BackupPath "C:\Backups\MailEnable"

.EXAMPLE
    .\Remove-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "spam.account" -Force
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [Parameter(Mandatory=$true)]
    [string]$PostOffice,

    [Parameter(Mandatory=$true)]
    [string]$Mailbox,

    [Parameter(Mandatory=$false)]
    [switch]$BackupFirst,

    [Parameter(Mandatory=$false)]
    [string]$BackupPath = "C:\MailEnable\Backups",

    [Parameter(Mandatory=$false)]
    [switch]$Force
)

# MailEnable COM object
$MailEnableApp = "MEAIPO.Mailbox"

try {
    Write-Host "Processing mailbox removal: $Mailbox@$PostOffice" -ForegroundColor Cyan

    # Backup if requested
    if ($BackupFirst) {
        Write-Host "Backing up mailbox before deletion..." -ForegroundColor Yellow

        if (-not (Test-Path $BackupPath)) {
            New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
        }

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFile = Join-Path $BackupPath "$Mailbox@$PostOffice-$timestamp.zip"

        # Get mailbox path (typical location)
        $mailboxPath = "C:\Program Files (x86)\Mail Enable\Postoffices\$PostOffice\MAILROOT\$Mailbox"

        if (Test-Path $mailboxPath) {
            Compress-Archive -Path $mailboxPath -DestinationPath $backupFile -Force
            Write-Host "✓ Backup created: $backupFile" -ForegroundColor Green
        } else {
            Write-Host "⚠ Mailbox path not found: $mailboxPath" -ForegroundColor Yellow
        }
    }

    # Confirmation
    if (-not $Force -and -not $PSCmdlet.ShouldProcess("$Mailbox@$PostOffice", "Delete mailbox")) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }

    # Create COM object
    $MEMailbox = New-Object -ComObject $MailEnableApp

    # Set mailbox properties
    $MEMailbox.PostOffice = $PostOffice
    $MEMailbox.MailboxName = $Mailbox

    # Remove the mailbox
    $result = $MEMailbox.RemoveMailbox()

    if ($result -eq 0) {
        Write-Host "✓ Mailbox removed successfully: $Mailbox@$PostOffice" -ForegroundColor Green
    } else {
        Write-Host "✗ Failed to remove mailbox. Error code: $result" -ForegroundColor Red
    }

    # Clean up COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null

} catch {
    Write-Host "✗ Error removing mailbox: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Alternative method using MailEnable API
<#
function Remove-MEMailboxAPI {
    param(
        [string]$PostOffice,
        [string]$Mailbox
    )

    $url = "http://localhost:8080/api/mailbox/remove"
    $body = @{
        postoffice = $PostOffice
        mailbox = $Mailbox
    } | ConvertTo-Json

    Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json"
}
#>
