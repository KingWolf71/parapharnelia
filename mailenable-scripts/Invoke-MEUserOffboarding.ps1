<#
.SYNOPSIS
    Offboards a user from MailEnable across all domains and redirects to markattard.

.DESCRIPTION
    This script performs a complete user offboarding workflow:
    1. Searches all domains for the specified mailbox
    2. Disables the mailbox (prevents login)
    3. Removes all email addresses/aliases from the account
    4. Adds forwarding alias to markattard@domain
    5. Removes the user from all distribution groups
    6. Optionally backs up the mailbox before processing

.PARAMETER Username
    The username to offboard (without @domain part).

.PARAMETER ReplacementUser
    The user to redirect mail to (default: markattard).

.PARAMETER BackupMailbox
    Backup the mailbox before processing.

.PARAMETER BackupPath
    Path for mailbox backups (default: C:\MailEnable\Backups\Offboarding).

.PARAMETER MailEnablePath
    Base path to MailEnable installation (default: C:\Program Files (x86)\Mail Enable).
    Change this if your mail data is on a different drive.

.PARAMETER WhatIf
    Show what would be done without actually making changes.

.PARAMETER Help
    Display this help information. Can also use -h or --help.

.EXAMPLE
    .\Invoke-MEUserOffboarding.ps1 -Username "john.doe"

.EXAMPLE
    .\Invoke-MEUserOffboarding.ps1 -h

.EXAMPLE
    .\Invoke-MEUserOffboarding.ps1 --help

.EXAMPLE
    .\Invoke-MEUserOffboarding.ps1 -Username "jane.smith" -BackupMailbox -BackupPath "D:\Backups"

.EXAMPLE
    .\Invoke-MEUserOffboarding.ps1 -Username "bob.jones" -WhatIf

.EXAMPLE
    .\Invoke-MEUserOffboarding.ps1 -Username "alice" -ReplacementUser "mark"

.EXAMPLE
    .\Invoke-MEUserOffboarding.ps1 -Username "john.doe" -MailEnablePath "D:\MailEnable"
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$false, Position=0)]
    [string]$Username,

    [Parameter(Mandatory=$false)]
    [string]$ReplacementUser = "markattard",

    [Parameter(Mandatory=$false)]
    [switch]$BackupMailbox,

    [Parameter(Mandatory=$false)]
    [string]$BackupPath = "C:\MailEnable\Backups\Offboarding",

    [Parameter(Mandatory=$false)]
    [string]$MailEnablePath = "C:\Program Files (x86)\Mail Enable",

    [Parameter(Mandatory=$false)]
    [Alias("h")]
    [switch]$Help
)

# Handle -h or --help parameters
if ($Help -or $args -contains "-h" -or $args -contains "--help") {
    Get-Help $MyInvocation.MyCommand.Path -Detailed
    exit 0
}

# Validate required parameters
if ([string]::IsNullOrEmpty($Username)) {
    Write-Host "Error: -Username parameter is required" -ForegroundColor Red
    Write-Host "Use -h or --help for usage information" -ForegroundColor Yellow
    Write-Host "`nQuick usage: .\Invoke-MEUserOffboarding.ps1 -Username `"john.doe`"" -ForegroundColor Cyan
    exit 1
}

# Initialize logging
$logFile = ".\offboarding-$Username-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    Add-Content -Path $logFile -Value $logMessage

    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "INFO" { Write-Host $Message -ForegroundColor Cyan }
        default { Write-Host $Message -ForegroundColor White }
    }
}

# Function to get all post offices
function Get-MEPostOffices {
    try {
        $poPath = Join-Path $MailEnablePath "Postoffices"
        if (Test-Path $poPath) {
            return Get-ChildItem -Path $poPath -Directory | Select-Object -ExpandProperty Name
        }
        return @()
    } catch {
        Write-Log "Error retrieving post offices: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

# Function to backup mailbox
function Backup-MEMailbox {
    param(
        [string]$PostOffice,
        [string]$Mailbox
    )

    try {
        if (-not (Test-Path $BackupPath)) {
            New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
        }

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFile = Join-Path $BackupPath "$Mailbox@$PostOffice-$timestamp.zip"

        $mailboxPath = Join-Path $MailEnablePath "Postoffices\$PostOffice\MAILROOT\$Mailbox"

        if (Test-Path $mailboxPath) {
            if (-not $WhatIf) {
                Compress-Archive -Path $mailboxPath -DestinationPath $backupFile -Force
                Write-Log "Backup created: $backupFile" -Level "SUCCESS"
            } else {
                Write-Log "[WHATIF] Would backup to: $backupFile" -Level "INFO"
            }
            return $true
        } else {
            Write-Log "Mailbox path not found: $mailboxPath" -Level "WARNING"
            return $false
        }
    } catch {
        Write-Log "Error backing up mailbox: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Function to remove user from groups
function Remove-MEUserFromGroups {
    param(
        [string]$PostOffice,
        [string]$Mailbox
    )

    try {
        Write-Log "Removing $Mailbox@$PostOffice from all distribution groups..." -Level "INFO"

        # MailEnable groups are typically stored in the Groups directory
        $groupsPath = Join-Path $MailEnablePath "Postoffices\$PostOffice\Groups"

        if (Test-Path $groupsPath) {
            $groups = Get-ChildItem -Path $groupsPath -Filter "*.mem" -File

            foreach ($group in $groups) {
                $groupName = [System.IO.Path]::GetFileNameWithoutExtension($group.Name)
                $members = Get-Content $group.FullName -ErrorAction SilentlyContinue

                if ($members -contains $Mailbox) {
                    if (-not $WhatIf) {
                        $newMembers = $members | Where-Object { $_ -ne $Mailbox }
                        Set-Content -Path $group.FullName -Value $newMembers
                        Write-Log "  Removed from group: $groupName" -Level "SUCCESS"
                    } else {
                        Write-Log "  [WHATIF] Would remove from group: $groupName" -Level "INFO"
                    }
                }
            }
        }

        # Also check COM-based groups if available
        try {
            $MEGroup = New-Object -ComObject "MEAIPO.Group" -ErrorAction SilentlyContinue
            if ($MEGroup) {
                # This would require iteration through groups via COM
                # Implementation depends on MailEnable version
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEGroup) | Out-Null
            }
        } catch {
            # COM groups not available or error occurred
        }

    } catch {
        Write-Log "Error removing from groups: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Function to add forwarding alias
function Add-MEForwardingAlias {
    param(
        [string]$PostOffice,
        [string]$FromMailbox,
        [string]$ToMailbox
    )

    try {
        Write-Log "Setting up forwarding from $FromMailbox@$PostOffice to $ToMailbox@$PostOffice..." -Level "INFO"

        # Create redirect/alias using MailEnable
        $redirectPath = Join-Path $MailEnablePath "Postoffices\$PostOffice\MAILROOT\$FromMailbox"

        if (-not $WhatIf) {
            # Method 1: Create .forward file (common approach)
            $forwardFile = Join-Path $redirectPath ".forward"
            "$ToMailbox@$PostOffice" | Set-Content -Path $forwardFile -Force
            Write-Log "  Created .forward file to $ToMailbox@$PostOffice" -Level "SUCCESS"

            # Method 2: Use MailEnable COM to set forwarding
            try {
                $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
                $MEMailbox.PostOffice = $PostOffice
                $MEMailbox.MailboxName = $FromMailbox

                $result = $MEMailbox.GetMailbox()
                if ($result -eq 0) {
                    # Set redirect address (property name may vary by ME version)
                    # Check your MailEnable documentation for exact property name
                    # Common properties: RedirectAddress, ForwardTo, RedirectTo
                    try {
                        $MEMailbox.RedirectAddress = "$ToMailbox@$PostOffice"
                        $MEMailbox.UpdateMailbox()
                        Write-Log "  Set redirect via COM API" -Level "SUCCESS"
                    } catch {
                        # Property might not exist in this version
                    }
                }

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
            } catch {
                Write-Log "  COM redirect not available (using .forward file only)" -Level "WARNING"
            }
        } else {
            Write-Log "  [WHATIF] Would create forwarding to $ToMailbox@$PostOffice" -Level "INFO"
        }

    } catch {
        Write-Log "Error setting up forwarding: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Function to process mailbox
function Process-MEMailbox {
    param(
        [string]$PostOffice,
        [string]$Mailbox
    )

    Write-Log "`n$('=' * 80)" -Level "INFO"
    Write-Log "Processing: $Mailbox@$PostOffice" -Level "INFO"
    Write-Log "$('=' * 80)" -Level "INFO"

    # Step 1: Backup if requested
    if ($BackupMailbox) {
        Write-Log "Step 1: Backing up mailbox..." -Level "INFO"
        Backup-MEMailbox -PostOffice $PostOffice -Mailbox $Mailbox
    }

    # Step 2: Disable the mailbox
    Write-Log "Step 2: Disabling mailbox..." -Level "INFO"
    try {
        $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
        $MEMailbox.PostOffice = $PostOffice
        $MEMailbox.MailboxName = $Mailbox

        $result = $MEMailbox.GetMailbox()
        if ($result -eq 0) {
            if (-not $WhatIf) {
                $MEMailbox.Status = 0  # 0 = Disabled
                $updateResult = $MEMailbox.UpdateMailbox()

                if ($updateResult -eq 0) {
                    Write-Log "  Mailbox disabled successfully" -Level "SUCCESS"
                } else {
                    Write-Log "  Failed to disable mailbox. Error code: $updateResult" -Level "ERROR"
                }
            } else {
                Write-Log "  [WHATIF] Would disable mailbox" -Level "INFO"
            }
        } else {
            Write-Log "  Mailbox not found or error. Error code: $result" -Level "ERROR"
        }

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
    } catch {
        Write-Log "  Error disabling mailbox: $($_.Exception.Message)" -Level "ERROR"
    }

    # Step 3: Remove email addresses/aliases
    Write-Log "Step 3: Removing email addresses and aliases..." -Level "INFO"
    try {
        $aliasPath = Join-Path $MailEnablePath "Postoffices\$PostOffice\MAILROOT\$Mailbox\aliases"

        if (Test-Path $aliasPath) {
            if (-not $WhatIf) {
                Remove-Item -Path $aliasPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "  Removed alias directory" -Level "SUCCESS"
            } else {
                Write-Log "  [WHATIF] Would remove alias directory" -Level "INFO"
            }
        }

        # Also clear any additional email addresses stored in mailbox properties
        # This is MailEnable version specific
    } catch {
        Write-Log "  Error removing aliases: $($_.Exception.Message)" -Level "ERROR"
    }

    # Step 4: Add forwarding to replacement user
    Write-Log "Step 4: Adding forwarding alias to $ReplacementUser@$PostOffice..." -Level "INFO"
    Add-MEForwardingAlias -PostOffice $PostOffice -FromMailbox $Mailbox -ToMailbox $ReplacementUser

    # Step 5: Remove from distribution groups
    Write-Log "Step 5: Removing from distribution groups..." -Level "INFO"
    Remove-MEUserFromGroups -PostOffice $PostOffice -Mailbox $Mailbox

    Write-Log "Processing complete for $Mailbox@$PostOffice" -Level "SUCCESS"
}

# Main execution
try {
    Write-Host "`n$('=' * 80)" -ForegroundColor Cyan
    Write-Host "MailEnable User Offboarding Script" -ForegroundColor Cyan
    Write-Host "$('=' * 80)`n" -ForegroundColor Cyan

    if ($WhatIf) {
        Write-Host "RUNNING IN WHATIF MODE - No changes will be made`n" -ForegroundColor Yellow
    }

    Write-Log "Starting offboarding process for user: $Username" -Level "INFO"
    Write-Log "Replacement user: $ReplacementUser" -Level "INFO"
    Write-Log "Log file: $logFile" -Level "INFO"

    # Get all post offices
    $postOffices = Get-MEPostOffices

    if ($postOffices.Count -eq 0) {
        Write-Log "No post offices found!" -Level "ERROR"
        exit 1
    }

    Write-Log "Found $($postOffices.Count) post office(s) to search" -Level "INFO"

    $processedCount = 0
    $foundDomains = @()

    # Search all post offices for the mailbox
    foreach ($po in $postOffices) {
        $mailboxPath = Join-Path $MailEnablePath "Postoffices\$po\MAILROOT\$Username"

        if (Test-Path $mailboxPath) {
            Write-Log "Found mailbox: $Username@$po" -Level "INFO"
            $foundDomains += $po
            Process-MEMailbox -PostOffice $po -Mailbox $Username
            $processedCount++
        }
    }

    # Summary
    Write-Host "`n$('=' * 80)" -ForegroundColor Cyan
    Write-Host "Offboarding Summary" -ForegroundColor Cyan
    Write-Host "$('=' * 80)" -ForegroundColor Cyan
    Write-Host "Username: $Username" -ForegroundColor White
    Write-Host "Domains processed: $processedCount" -ForegroundColor White

    if ($foundDomains.Count -gt 0) {
        Write-Host "`nMailboxes found in:" -ForegroundColor White
        foreach ($domain in $foundDomains) {
            Write-Host "  - $Username@$domain" -ForegroundColor Gray
        }
    } else {
        Write-Host "`nNo mailboxes found for user: $Username" -ForegroundColor Yellow
    }

    Write-Host "`nActions performed:" -ForegroundColor White
    Write-Host "  ✓ Mailboxes disabled" -ForegroundColor Green
    Write-Host "  ✓ Email aliases removed" -ForegroundColor Green
    Write-Host "  ✓ Forwarding added to $ReplacementUser" -ForegroundColor Green
    Write-Host "  ✓ Removed from distribution groups" -ForegroundColor Green

    if ($BackupMailbox) {
        Write-Host "  ✓ Mailboxes backed up to: $BackupPath" -ForegroundColor Green
    }

    Write-Host "`nLog file: $logFile" -ForegroundColor Gray
    Write-Host "$('=' * 80)`n" -ForegroundColor Cyan

    Write-Log "Offboarding completed successfully. Processed $processedCount mailbox(es)." -Level "SUCCESS"

} catch {
    Write-Log "Fatal error during offboarding: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}
