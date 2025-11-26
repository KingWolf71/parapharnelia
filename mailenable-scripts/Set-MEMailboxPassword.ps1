<#
.SYNOPSIS
    Changes the password for a MailEnable mailbox.

.DESCRIPTION
    This script updates the password for an existing mailbox in MailEnable.
    Can generate a random secure password if not provided.

.PARAMETER PostOffice
    The MailEnable post office (domain) where the mailbox exists.

.PARAMETER Mailbox
    The mailbox name to update.

.PARAMETER NewPassword
    The new password to set. If not provided, a random password will be generated.

.PARAMETER GenerateRandom
    Generate a random secure password.

.PARAMETER PasswordLength
    Length of the generated password (default: 16).

.PARAMETER ShowPassword
    Display the new password (useful when auto-generating).

.EXAMPLE
    .\Set-MEMailboxPassword.ps1 -PostOffice "example.com" -Mailbox "john.doe" -NewPassword "NewP@ssw0rd123"

.EXAMPLE
    .\Set-MEMailboxPassword.ps1 -PostOffice "example.com" -Mailbox "john.doe" -GenerateRandom -ShowPassword

.EXAMPLE
    .\Set-MEMailboxPassword.ps1 -PostOffice "example.com" -Mailbox "jane.smith" -GenerateRandom -PasswordLength 20 -ShowPassword
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PostOffice,

    [Parameter(Mandatory=$true)]
    [string]$Mailbox,

    [Parameter(Mandatory=$false)]
    [string]$NewPassword,

    [Parameter(Mandatory=$false)]
    [switch]$GenerateRandom,

    [Parameter(Mandatory=$false)]
    [int]$PasswordLength = 16,

    [Parameter(Mandatory=$false)]
    [switch]$ShowPassword
)

# Generate random password function
function New-RandomPassword {
    param([int]$Length = 16)

    $uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $lowercase = "abcdefghijklmnopqrstuvwxyz"
    $numbers = "0123456789"
    $symbols = "!@#$%^&*()_+-=[]{}|"

    $allChars = $uppercase + $lowercase + $numbers + $symbols

    # Ensure at least one of each type
    $password = @()
    $password += $uppercase[(Get-Random -Maximum $uppercase.Length)]
    $password += $lowercase[(Get-Random -Maximum $lowercase.Length)]
    $password += $numbers[(Get-Random -Maximum $numbers.Length)]
    $password += $symbols[(Get-Random -Maximum $symbols.Length)]

    # Fill the rest randomly
    for ($i = 4; $i -lt $Length; $i++) {
        $password += $allChars[(Get-Random -Maximum $allChars.Length)]
    }

    # Shuffle the password
    $password = $password | Get-Random -Count $password.Count

    return -join $password
}

# MailEnable COM object
$MailEnableApp = "MEAIPO.Mailbox"

try {
    Write-Host "Updating password for: $Mailbox@$PostOffice" -ForegroundColor Cyan

    # Generate password if requested or not provided
    if ($GenerateRandom -or [string]::IsNullOrEmpty($NewPassword)) {
        $NewPassword = New-RandomPassword -Length $PasswordLength
        Write-Host "Generated new password" -ForegroundColor Yellow
    }

    # Validate password strength
    if ($NewPassword.Length -lt 8) {
        Write-Host "⚠ Warning: Password is less than 8 characters. Consider using a stronger password." -ForegroundColor Yellow
    }

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

    # Update password
    $MEMailbox.Password = $NewPassword
    $result = $MEMailbox.UpdateMailbox()

    if ($result -eq 0) {
        Write-Host "✓ Password updated successfully" -ForegroundColor Green

        if ($ShowPassword) {
            Write-Host "`nNew Password: $NewPassword" -ForegroundColor Cyan
            Write-Host "(Make sure to save this password securely)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "✗ Failed to update password. Error code: $result" -ForegroundColor Red
    }

    # Clean up COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null

} catch {
    Write-Host "✗ Error updating password: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Alternative method for bulk password resets
<#
function Reset-MEMailboxPasswordBulk {
    param(
        [string]$CSVPath
    )

    # CSV should have columns: PostOffice, Mailbox, NewPassword
    $users = Import-Csv -Path $CSVPath

    foreach ($user in $users) {
        try {
            $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
            $MEMailbox.PostOffice = $user.PostOffice
            $MEMailbox.MailboxName = $user.Mailbox

            $result = $MEMailbox.GetMailbox()
            if ($result -eq 0) {
                $MEMailbox.Password = $user.NewPassword
                $MEMailbox.UpdateMailbox()
                Write-Host "✓ Updated: $($user.Mailbox)@$($user.PostOffice)" -ForegroundColor Green
            }

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
        } catch {
            Write-Host "✗ Failed: $($user.Mailbox)@$($user.PostOffice)" -ForegroundColor Red
        }
    }
}
#>
