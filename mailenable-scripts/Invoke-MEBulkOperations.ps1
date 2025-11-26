<#
.SYNOPSIS
    Performs bulk operations on MailEnable mailboxes.

.DESCRIPTION
    This script allows batch processing of multiple mailbox operations using a CSV file.
    Supported operations: Add, Remove, SetQuota, ResetPassword

.PARAMETER Operation
    The operation to perform: Add, Remove, SetQuota, ResetPassword

.PARAMETER CSVPath
    Path to the CSV file containing mailbox data.

.PARAMETER LogPath
    Path for the operation log file (default: bulk-operations.log).

.EXAMPLE
    .\Invoke-MEBulkOperations.ps1 -Operation Add -CSVPath "new-users.csv"

.EXAMPLE
    .\Invoke-MEBulkOperations.ps1 -Operation SetQuota -CSVPath "quota-updates.csv"

.NOTES
    CSV Format for Add operation:
    PostOffice,Mailbox,Password,FirstName,LastName,Quota

    CSV Format for Remove operation:
    PostOffice,Mailbox

    CSV Format for SetQuota operation:
    PostOffice,Mailbox,QuotaMB

    CSV Format for ResetPassword operation:
    PostOffice,Mailbox,NewPassword
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Add", "Remove", "SetQuota", "ResetPassword")]
    [string]$Operation,

    [Parameter(Mandatory=$true)]
    [string]$CSVPath,

    [Parameter(Mandatory=$false)]
    [string]$LogPath = ".\bulk-operations.log"
)

# Initialize counters
$successCount = 0
$failCount = 0
$totalCount = 0

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    Add-Content -Path $LogPath -Value $logMessage

    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        default { Write-Host $Message -ForegroundColor White }
    }
}

# Validate CSV exists
if (-not (Test-Path $CSVPath)) {
    Write-Log "CSV file not found: $CSVPath" -Level "ERROR"
    exit 1
}

# Import CSV
try {
    $data = Import-Csv -Path $CSVPath
    $totalCount = $data.Count
    Write-Log "Loaded $totalCount records from CSV" -Level "INFO"
} catch {
    Write-Log "Error reading CSV: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

Write-Log "Starting bulk operation: $Operation" -Level "INFO"
Write-Host "`n$('=' * 80)" -ForegroundColor Cyan
Write-Host "Bulk Operation: $Operation" -ForegroundColor Cyan
Write-Host "Records to process: $totalCount" -ForegroundColor Cyan
Write-Host "$('=' * 80)`n" -ForegroundColor Cyan

# Process each record
$currentRecord = 0
foreach ($record in $data) {
    $currentRecord++
    $percentComplete = [math]::Round(($currentRecord / $totalCount) * 100, 2)

    Write-Progress -Activity "Processing $Operation" -Status "Record $currentRecord of $totalCount ($percentComplete%)" -PercentComplete $percentComplete

    try {
        switch ($Operation) {
            "Add" {
                Write-Log "Adding mailbox: $($record.Mailbox)@$($record.PostOffice)" -Level "INFO"

                $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
                $MEMailbox.PostOffice = $record.PostOffice
                $MEMailbox.MailboxName = $record.Mailbox
                $MEMailbox.Password = $record.Password
                $MEMailbox.FirstName = $record.FirstName
                $MEMailbox.LastName = $record.LastName
                $MEMailbox.QuotaMB = [int]$record.Quota
                $MEMailbox.Status = 1

                $result = $MEMailbox.AddMailbox()

                if ($result -eq 0) {
                    Write-Log "✓ Successfully added: $($record.Mailbox)@$($record.PostOffice)" -Level "SUCCESS"
                    $successCount++
                } else {
                    Write-Log "✗ Failed to add: $($record.Mailbox)@$($record.PostOffice) - Error code: $result" -Level "ERROR"
                    $failCount++
                }

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
            }

            "Remove" {
                Write-Log "Removing mailbox: $($record.Mailbox)@$($record.PostOffice)" -Level "INFO"

                $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
                $MEMailbox.PostOffice = $record.PostOffice
                $MEMailbox.MailboxName = $record.Mailbox

                $result = $MEMailbox.RemoveMailbox()

                if ($result -eq 0) {
                    Write-Log "✓ Successfully removed: $($record.Mailbox)@$($record.PostOffice)" -Level "SUCCESS"
                    $successCount++
                } else {
                    Write-Log "✗ Failed to remove: $($record.Mailbox)@$($record.PostOffice) - Error code: $result" -Level "ERROR"
                    $failCount++
                }

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
            }

            "SetQuota" {
                Write-Log "Setting quota for: $($record.Mailbox)@$($record.PostOffice) to $($record.QuotaMB) MB" -Level "INFO"

                $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
                $MEMailbox.PostOffice = $record.PostOffice
                $MEMailbox.MailboxName = $record.Mailbox

                $getResult = $MEMailbox.GetMailbox()
                if ($getResult -eq 0) {
                    $MEMailbox.QuotaMB = [int]$record.QuotaMB
                    $updateResult = $MEMailbox.UpdateMailbox()

                    if ($updateResult -eq 0) {
                        Write-Log "✓ Successfully updated quota: $($record.Mailbox)@$($record.PostOffice)" -Level "SUCCESS"
                        $successCount++
                    } else {
                        Write-Log "✗ Failed to update quota: $($record.Mailbox)@$($record.PostOffice) - Error code: $updateResult" -Level "ERROR"
                        $failCount++
                    }
                } else {
                    Write-Log "✗ Mailbox not found: $($record.Mailbox)@$($record.PostOffice)" -Level "ERROR"
                    $failCount++
                }

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
            }

            "ResetPassword" {
                Write-Log "Resetting password for: $($record.Mailbox)@$($record.PostOffice)" -Level "INFO"

                $MEMailbox = New-Object -ComObject "MEAIPO.Mailbox"
                $MEMailbox.PostOffice = $record.PostOffice
                $MEMailbox.MailboxName = $record.Mailbox

                $getResult = $MEMailbox.GetMailbox()
                if ($getResult -eq 0) {
                    $MEMailbox.Password = $record.NewPassword
                    $updateResult = $MEMailbox.UpdateMailbox()

                    if ($updateResult -eq 0) {
                        Write-Log "✓ Successfully reset password: $($record.Mailbox)@$($record.PostOffice)" -Level "SUCCESS"
                        $successCount++
                    } else {
                        Write-Log "✗ Failed to reset password: $($record.Mailbox)@$($record.PostOffice) - Error code: $updateResult" -Level "ERROR"
                        $failCount++
                    }
                } else {
                    Write-Log "✗ Mailbox not found: $($record.Mailbox)@$($record.PostOffice)" -Level "ERROR"
                    $failCount++
                }

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MEMailbox) | Out-Null
            }
        }
    } catch {
        Write-Log "✗ Exception processing $($record.Mailbox)@$($record.PostOffice): $($_.Exception.Message)" -Level "ERROR"
        $failCount++
    }

    Start-Sleep -Milliseconds 100  # Small delay to avoid overwhelming the server
}

Write-Progress -Activity "Processing $Operation" -Completed

# Summary
Write-Host "`n$('=' * 80)" -ForegroundColor Cyan
Write-Host "Bulk Operation Summary" -ForegroundColor Cyan
Write-Host "$('=' * 80)" -ForegroundColor Cyan
Write-Host "Operation: $Operation" -ForegroundColor White
Write-Host "Total Records: $totalCount" -ForegroundColor White
Write-Host "Successful: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red
Write-Host "Success Rate: $([math]::Round(($successCount / $totalCount) * 100, 2))%" -ForegroundColor Yellow
Write-Host "$('=' * 80)`n" -ForegroundColor Cyan

Write-Log "Bulk operation completed. Success: $successCount, Failed: $failCount" -Level "INFO"
Write-Host "Log file: $LogPath" -ForegroundColor Gray
