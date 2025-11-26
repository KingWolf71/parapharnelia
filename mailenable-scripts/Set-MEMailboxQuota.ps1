<#
.SYNOPSIS
    Sets or updates mailbox quota in MailEnable.

.DESCRIPTION
    This script modifies the quota (storage limit) for an existing mailbox in MailEnable.
    Can also disable quota by setting to 0 (unlimited).

.PARAMETER PostOffice
    The MailEnable post office (domain) where the mailbox exists.

.PARAMETER Mailbox
    The mailbox name to modify.

.PARAMETER QuotaMB
    The quota size in megabytes. Set to 0 for unlimited.

.PARAMETER ShowCurrent
    Display current quota before making changes.

.EXAMPLE
    .\Set-MEMailboxQuota.ps1 -PostOffice "example.com" -Mailbox "john.doe" -QuotaMB 500

.EXAMPLE
    .\Set-MEMailboxQuota.ps1 -PostOffice "example.com" -Mailbox "vip.user" -QuotaMB 0

.EXAMPLE
    .\Set-MEMailboxQuota.ps1 -PostOffice "example.com" -Mailbox "john.doe" -QuotaMB 1000 -ShowCurrent
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PostOffice,

    [Parameter(Mandatory=$true)]
    [string]$Mailbox,

    [Parameter(Mandatory=$true)]
    [int]$QuotaMB,

    [Parameter(Mandatory=$false)]
    [switch]$ShowCurrent
)

# MailEnable COM object
$MailEnableApp = "MEAIPO.Mailbox"

try {
    Write-Host "Updating quota for: $Mailbox@$PostOffice" -ForegroundColor Cyan

    # Create COM object
    $MEMailbox = New-Object -ComObject $MailEnableApp

    # Set mailbox properties
    $MEMailbox.PostOffice = $PostOffice
    $MEMailbox.MailboxName = $Mailbox

    # Get mailbox details first
    $result = $MEMailbox.GetMailbox()

    if ($result -ne 0) {
        Write-Host "✗ Mailbox not found or error retrieving details. Error code: $result" -ForegroundColor Red
        exit 1
    }

    # Show current quota if requested
    if ($ShowCurrent) {
        $currentQuota = $MEMailbox.QuotaMB
        $quotaDisplay = if ($currentQuota -eq 0) { "Unlimited" } else { "$currentQuota MB" }
        Write-Host "Current quota: $quotaDisplay" -ForegroundColor Yellow
    }

    # Update quota
    $MEMailbox.QuotaMB = $QuotaMB
    $result = $MEMailbox.UpdateMailbox()

    if ($result -eq 0) {
        $newQuotaDisplay = if ($QuotaMB -eq 0) { "Unlimited" } else { "$QuotaMB MB" }
        Write-Host "✓ Quota updated successfully: $newQuotaDisplay" -ForegroundColor Green
    } else {
        Write-Host "✗ Failed to update quota. Error code: $result" -ForegroundColor Red
    }

    # Clean up COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null

} catch {
    Write-Host "✗ Error updating quota: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Function to get quota usage information
function Get-MEMailboxQuotaUsage {
    param(
        [string]$PostOffice,
        [string]$Mailbox
    )

    try {
        $mailboxPath = "C:\Program Files (x86)\Mail Enable\Postoffices\$PostOffice\MAILROOT\$Mailbox"

        if (Test-Path $mailboxPath) {
            $size = (Get-ChildItem -Path $mailboxPath -Recurse -File | Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 2)

            Write-Host "`nCurrent Usage:" -ForegroundColor Cyan
            Write-Host "  Path: $mailboxPath" -ForegroundColor Gray
            Write-Host "  Size: $sizeMB MB" -ForegroundColor Gray

            return $sizeMB
        } else {
            Write-Host "⚠ Mailbox path not found: $mailboxPath" -ForegroundColor Yellow
            return 0
        }
    } catch {
        Write-Host "⚠ Could not calculate mailbox size: $($_.Exception.Message)" -ForegroundColor Yellow
        return 0
    }
}

# Optionally show usage
if ($ShowCurrent) {
    Get-MEMailboxQuotaUsage -PostOffice $PostOffice -Mailbox $Mailbox
}
