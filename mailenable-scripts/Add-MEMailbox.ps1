<#
.SYNOPSIS
    Creates a new mailbox in MailEnable.

.DESCRIPTION
    This script creates a new mailbox for a specified user in MailEnable.
    It can create mailboxes in both the Standard and Professional editions.

.PARAMETER PostOffice
    The MailEnable post office (domain) where the mailbox will be created.

.PARAMETER Mailbox
    The mailbox name (email username).

.PARAMETER Password
    The password for the mailbox.

.PARAMETER FirstName
    The user's first name (optional).

.PARAMETER LastName
    The user's last name (optional).

.PARAMETER Quota
    Mailbox quota in MB (default: 100MB).

.EXAMPLE
    .\Add-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "john.doe" -Password "P@ssw0rd123"

.EXAMPLE
    .\Add-MEMailbox.ps1 -PostOffice "example.com" -Mailbox "jane.smith" -Password "SecureP@ss" -FirstName "Jane" -LastName "Smith" -Quota 500
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PostOffice,

    [Parameter(Mandatory=$true)]
    [string]$Mailbox,

    [Parameter(Mandatory=$true)]
    [string]$Password,

    [Parameter(Mandatory=$false)]
    [string]$FirstName = "",

    [Parameter(Mandatory=$false)]
    [string]$LastName = "",

    [Parameter(Mandatory=$false)]
    [int]$Quota = 100
)

# MailEnable COM object for mailbox management
$MailEnableApp = "MEAIPO.Mailbox"

try {
    Write-Host "Creating mailbox: $Mailbox@$PostOffice" -ForegroundColor Cyan

    # Create COM object
    $MEMailbox = New-Object -ComObject $MailEnableApp

    # Set mailbox properties
    $MEMailbox.PostOffice = $PostOffice
    $MEMailbox.MailboxName = $Mailbox
    $MEMailbox.Password = $Password
    $MEMailbox.FirstName = $FirstName
    $MEMailbox.LastName = $LastName
    $MEMailbox.QuotaMB = $Quota
    $MEMailbox.Status = 1  # 1 = Active, 0 = Disabled

    # Create the mailbox
    $result = $MEMailbox.AddMailbox()

    if ($result -eq 0) {
        Write-Host "✓ Mailbox created successfully: $Mailbox@$PostOffice" -ForegroundColor Green
        Write-Host "  First Name: $FirstName" -ForegroundColor Gray
        Write-Host "  Last Name: $LastName" -ForegroundColor Gray
        Write-Host "  Quota: $Quota MB" -ForegroundColor Gray
    } else {
        Write-Host "✗ Failed to create mailbox. Error code: $result" -ForegroundColor Red
    }

    # Clean up COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null

} catch {
    Write-Host "✗ Error creating mailbox: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Alternative method using direct database access (SQL Server)
<#
function Add-MEMailboxSQL {
    param(
        [string]$ServerInstance = "localhost\MAILENABLE",
        [string]$Database = "MEPO",
        [string]$PostOffice,
        [string]$Mailbox,
        [string]$Password,
        [int]$Quota
    )

    $Query = @"
INSERT INTO [Mailbox] ([PostOffice], [MailboxName], [Password], [Status], [QuotaMB])
VALUES ('$PostOffice', '$Mailbox', '$Password', 1, $Quota)
"@

    Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $Query
}
#>
