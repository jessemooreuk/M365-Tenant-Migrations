<#
.SYNOPSIS
    Set 150 MB Send & Receive Size on ALL current mailboxes in the tenant.

.DESCRIPTION
    Queries all mailboxes live using Get-Mailbox (no input CSV required).
    Applies MaxSendSize and MaxReceiveSize of 150 MB to every mailbox.
    Creates a timestamped CSV report with results.

.EXAMPLE
    .\Set-TenantWideMailboxSizeLimits.ps1 -WhatIf     # Preview only

.EXAMPLE
    .\Set-TenantWideMailboxSizeLimits.ps1             # Apply changes
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

$ErrorActionPreference = 'Stop'
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = ".\TenantWide_150MB_Limits_$Timestamp.csv"

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "  Tenant-Wide 150 MB Mailbox Size Limits (Live Query)" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

# Connect to Exchange Online if not already connected
if (-not (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}

$Results = @()
$Counter = 0

# === Live lookup of current mailboxes (no CSV input) ===
Get-Mailbox -ResultSize Unlimited | ForEach-Object {

    $Counter++
    Write-Progress -Activity "Updating mailbox size limits to 150 MB" `
                   -Status "$Counter - $($_.DisplayName)" `
                   -PercentComplete (($Counter / (Get-Mailbox -ResultSize Unlimited | Measure-Object).Count) * 100)

    $Result = [PSCustomObject]@{
        UPN            = $_.UserPrincipalName
        DisplayName    = $_.DisplayName
        Status         = ""
        Error          = ""
        Timestamp      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    try {
        if ($PSCmdlet.ShouldProcess($_.UserPrincipalName, "Set MaxSendSize & MaxReceiveSize to 150MB")) {
            Set-Mailbox -Identity $_.UserPrincipalName `
                        -MaxSendSize 150MB `
                        -MaxReceiveSize 150MB `
                        -ErrorAction Stop

            $Result.Status = "Success"
        }
    } catch {
        $Result.Status = "Failed"
        $Result.Error  = $_.Exception.Message
    }

    $Results += $Result
}

Write-Progress -Activity "Updating mailbox size limits to 150 MB" -Completed

# Export report
$Results | Export-Csv $ReportFile -NoTypeInformation
Write-Host "`nReport saved to: $ReportFile" -ForegroundColor Green

# Summary
$Success = ($Results | Where-Object Status -eq 'Success').Count
$Failed  = ($Results | Where-Object Status -eq 'Failed').Count

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Total mailboxes processed : $($Results.Count)"
Write-Host "Successful                : $Success" -ForegroundColor Green
Write-Host "Failed                    : $Failed" -ForegroundColor $(if ($Failed -gt 0) { "Red" } else { "Green" })
Write-Host "===========================" -ForegroundColor Cyan
