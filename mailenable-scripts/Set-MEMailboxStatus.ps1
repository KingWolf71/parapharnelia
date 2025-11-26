<#
.SYNOPSIS
    Enables or disables a MailEnable mailbox.

.DESCRIPTION
    This script changes the status of a mailbox (active/disabled) without deleting it.
    Useful for temporarily suspending accounts.

.PARAMETER PostOffice
    The MailEnable post office (domain) where the mailbox exists.

.PARAMETER Mailbox
    The mailbox name to modify.

.PARAMETER Status
    The status to set: Enable or Disable

.PARAMETER ShowCurrent
    Display current status before making changes.

.EXAMPLE
    .\Set-MEMailboxStatus.ps1 -PostOffice "example.com" -Mailbox "john.doe" -Status Enable

.EXAMPLE
    .\Set-MEMailboxStatus.ps1 -PostOffice "example.com" -Mailbox "suspended.user" -Status Disable -ShowCurrent
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PostOffice,

    [Parameter(Mandatory=$true)]
    [string]$Mailbox,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Enable", "Disable")]
    [string]$Status,

    [Parameter(Mandatory=$false)]
    [switch]$ShowCurrent
)

# MailEnable COM object
$MailEnableApp = "MEAIPO.Mailbox"

try {
    Write-Host "Updating status for: $Mailbox@$PostOffice" -ForegroundColor Cyan

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

    # Show current status if requested
    if ($ShowCurrent) {
        $currentStatus = if ($MEMailbox.Status -eq 1) { "Enabled" } else { "Disabled" }
        Write-Host "Current status: $currentStatus" -ForegroundColor Yellow
    }

    # Update status
    $statusValue = if ($Status -eq "Enable") { 1 } else { 0 }
    $MEMailbox.Status = $statusValue

    $result = $MEMailbox.UpdateMailbox()

    if ($result -eq 0) {
        Write-Host "✓ Mailbox $($Status.ToLower())d successfully: $Mailbox@$PostOffice" -ForegroundColor Green

        if ($Status -eq "Disable") {
            Write-Host "`n⚠ Note: The mailbox is disabled but not deleted." -ForegroundColor Yellow
            Write-Host "  - Mail delivery is stopped" -ForegroundColor Gray
            Write-Host "  - User cannot log in" -ForegroundColor Gray
            Write-Host "  - Mailbox data is preserved" -ForegroundColor Gray
            Write-Host "  - Re-enable anytime with -Status Enable" -ForegroundColor Gray
        } else {
            Write-Host "`n✓ Mailbox is now active" -ForegroundColor Green
            Write-Host "  - User can log in" -ForegroundColor Gray
            Write-Host "  - Mail delivery is active" -ForegroundColor Gray
        }
    } else {
        Write-Host "✗ Failed to update status. Error code: $result" -ForegroundColor Red
    }

    # Clean up COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null

} catch {
    Write-Host "✗ Error updating mailbox status: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Bulk enable/disable function
<#
function Set-MEMailboxStatusBulk {
    param(
        [string]$PostOffice,
        [string[]]$Mailboxes,
        [ValidateSet("Enable", "Disable")]
        [string]$Status
    )

    $statusValue = if ($Status -eq "Enable") { 1 } else { 0 }

    foreach ($mailbox in $Mailboxes) {
        try {
            $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
            $MEMailbox.PostOffice = $PostOffice
            $MEMailbox.MailboxName = $mailbox

            $result = $MEMailbox.GetMailbox()
            if ($result -eq 0) {
                $MEMailbox.Status = $statusValue
                $MEMailbox.UpdateMailbox()
                Write-Host "✓ $Status`: $mailbox@$PostOffice" -ForegroundColor Green
            } else {
                Write-Host "✗ Not found: $mailbox@$PostOffice" -ForegroundColor Red
            }

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
        } catch {
            Write-Host "✗ Error: $mailbox@$PostOffice - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Example usage:
# Set-MEMailboxStatusBulk -PostOffice "example.com" -Mailboxes @("user1", "user2", "user3") -Status Disable
#>
