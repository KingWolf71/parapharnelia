<#
.SYNOPSIS
    Lists mailboxes from MailEnable.

.DESCRIPTION
    This script retrieves and displays mailboxes from MailEnable.
    Can filter by post office, show detailed information, and export to CSV.

.PARAMETER PostOffice
    The MailEnable post office (domain) to query. If not specified, lists from all post offices.

.PARAMETER Detailed
    Show detailed information including quota, status, and size.

.PARAMETER ExportCSV
    Export results to a CSV file.

.PARAMETER CSVPath
    Path for the CSV export (default: mailboxes.csv in current directory).

.PARAMETER ShowInactive
    Include disabled/inactive mailboxes in the results.

.EXAMPLE
    .\Get-MEMailboxes.ps1

.EXAMPLE
    .\Get-MEMailboxes.ps1 -PostOffice "example.com" -Detailed

.EXAMPLE
    .\Get-MEMailboxes.ps1 -PostOffice "example.com" -ExportCSV -CSVPath "C:\Reports\mailboxes.csv"

.EXAMPLE
    .\Get-MEMailboxes.ps1 -PostOffice "example.com" -ShowInactive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$PostOffice,

    [Parameter(Mandatory=$false)]
    [switch]$Detailed,

    [Parameter(Mandatory=$false)]
    [switch]$ExportCSV,

    [Parameter(Mandatory=$false)]
    [string]$CSVPath = ".\mailboxes.csv",

    [Parameter(Mandatory=$false)]
    [switch]$ShowInactive
)

# Function to get mailbox size
function Get-MailboxSize {
    param(
        [string]$PostOffice,
        [string]$Mailbox
    )

    try {
        $mailboxPath = "C:\Program Files (x86)\Mail Enable\Postoffices\$PostOffice\MAILROOT\$Mailbox"

        if (Test-Path $mailboxPath) {
            $size = (Get-ChildItem -Path $mailboxPath -Recurse -File -ErrorAction SilentlyContinue |
                    Measure-Object -Property Length -Sum).Sum
            return [math]::Round($size / 1MB, 2)
        }
        return 0
    } catch {
        return 0
    }
}

# Function to get all post offices
function Get-PostOffices {
    try {
        $poPath = "C:\Program Files (x86)\Mail Enable\Postoffices"
        if (Test-Path $poPath) {
            return Get-ChildItem -Path $poPath -Directory | Select-Object -ExpandProperty Name
        }
        return @()
    } catch {
        return @()
    }
}

# Main script
try {
    Write-Host "Retrieving MailEnable mailboxes..." -ForegroundColor Cyan

    $mailboxList = @()
    $postOffices = @()

    # Determine which post offices to query
    if ([string]::IsNullOrEmpty($PostOffice)) {
        $postOffices = Get-PostOffices
        Write-Host "Scanning all post offices: $($postOffices.Count) found" -ForegroundColor Yellow
    } else {
        $postOffices = @($PostOffice)
    }

    # Query each post office
    foreach ($po in $postOffices) {
        try {
            # Get mailboxes from this post office
            $mailboxPath = "C:\Program Files (x86)\Mail Enable\Postoffices\$po\MAILROOT"

            if (Test-Path $mailboxPath) {
                $mailboxes = Get-ChildItem -Path $mailboxPath -Directory

                foreach ($mb in $mailboxes) {
                    $mailboxInfo = [PSCustomObject]@{
                        PostOffice = $po
                        Mailbox = $mb.Name
                        Email = "$($mb.Name)@$po"
                        Status = "Active"
                        QuotaMB = 0
                        SizeMB = 0
                        LastModified = $mb.LastWriteTime
                    }

                    # Get detailed info if requested
                    if ($Detailed) {
                        try {
                            $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
                            $MEMailbox.PostOffice = $po
                            $MEMailbox.MailboxName = $mb.Name

                            $result = $MEMailbox.GetMailbox()
                            if ($result -eq 0) {
                                $mailboxInfo.QuotaMB = $MEMailbox.QuotaMB
                                $mailboxInfo.Status = if ($MEMailbox.Status -eq 1) { "Active" } else { "Disabled" }
                            }

                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
                        } catch {
                            # Continue if COM fails
                        }

                        # Get actual size
                        $mailboxInfo.SizeMB = Get-MailboxSize -PostOffice $po -Mailbox $mb.Name
                    }

                    # Filter inactive if not requested
                    if ($ShowInactive -or $mailboxInfo.Status -eq "Active") {
                        $mailboxList += $mailboxInfo
                    }
                }
            }
        } catch {
            Write-Host "⚠ Error querying post office: $po - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # Display results
    Write-Host "`nFound $($mailboxList.Count) mailbox(es)" -ForegroundColor Green
    Write-Host ("=" * 80) -ForegroundColor Gray

    if ($Detailed) {
        $mailboxList | Format-Table -AutoSize Email, Status, QuotaMB, SizeMB, LastModified
    } else {
        $mailboxList | Format-Table -AutoSize Email, LastModified
    }

    # Export to CSV if requested
    if ($ExportCSV) {
        $mailboxList | Export-Csv -Path $CSVPath -NoTypeInformation -Force
        Write-Host "`n✓ Exported to: $CSVPath" -ForegroundColor Green
    }

    # Summary statistics
    if ($Detailed -and $mailboxList.Count -gt 0) {
        $totalSize = ($mailboxList | Measure-Object -Property SizeMB -Sum).Sum
        $avgSize = ($mailboxList | Measure-Object -Property SizeMB -Average).Average
        $totalQuota = ($mailboxList | Measure-Object -Property QuotaMB -Sum).Sum

        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "  Total Mailboxes: $($mailboxList.Count)" -ForegroundColor Gray
        Write-Host "  Active: $(($mailboxList | Where-Object {$_.Status -eq 'Active'}).Count)" -ForegroundColor Gray
        Write-Host "  Disabled: $(($mailboxList | Where-Object {$_.Status -eq 'Disabled'}).Count)" -ForegroundColor Gray
        Write-Host "  Total Size: $([math]::Round($totalSize, 2)) MB" -ForegroundColor Gray
        Write-Host "  Average Size: $([math]::Round($avgSize, 2)) MB" -ForegroundColor Gray
        Write-Host "  Total Quota Allocated: $totalQuota MB" -ForegroundColor Gray
    }

} catch {
    Write-Host "✗ Error retrieving mailboxes: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
