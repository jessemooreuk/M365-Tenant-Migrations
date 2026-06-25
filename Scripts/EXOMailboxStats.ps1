<#
.SYNOPSIS
    Exchange Online Mailbox Usage Report - Counts up every byte your users are hoarding.
    Now supports CSV, HTML, or Both export formats.
.DESCRIPTION
    Uses modern EXO V3 cmdlets. Gets User + Shared mailboxes (optionally more),
    pulls primary + archive statistics, converts sizes properly with ToBytes(),
    and exports a detailed report in your chosen format (CSV, HTML, or Both).
    Prints a brutal console summary of total storage used.
    Ideal for pre-migration storage assessments and identifying mailbox bloat.
#>
param(
    [Parameter(Mandatory=$false, HelpMessage="Admin UPN (optional - leave blank for interactive MFA)")]
    [string]$AdminUPN,
    [Parameter(Mandatory=$false)]
    [switch]$ExcludeShared, # Add -ExcludeShared to skip shared mailboxes
    [Parameter(Mandatory=$false)]
    [switch]$NoArchive, # Add -NoArchive to skip archive stats (faster on huge tenants)
    [Parameter(Mandatory=$false)]
    [ValidateSet('CSV','HTML','Both')]
    [string]$ExportFormat = 'CSV',
    [Parameter(Mandatory=$false, HelpMessage="Base output path/filename without extension (timestamp appended automatically)")]
    [string]$OutputBase = "$env:USERPROFILE\Desktop\EXO_Mailbox_Usage_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
)
# === MODULE CHECK & INSTALL ===
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module (CurrentUser)..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop
# === CONNECT ===
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
if ($AdminUPN) {
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
} else {
    Connect-ExchangeOnline -ShowBanner:$false
}
# === GET MAILBOXES ===
Write-Host "Fetching mailboxes... this can take a minute (or 30) on big tenants." -ForegroundColor Yellow
$recipientTypes = @('UserMailbox')
if (-not $ExcludeShared) { $recipientTypes += 'SharedMailbox' }
$Mailboxes = Get-EXOMailbox -ResultSize Unlimited `
    -RecipientTypeDetails $recipientTypes `
    -Properties ArchiveDatabase, ProhibitSendReceiveQuota, IssueWarningQuota |
    Sort-Object DisplayName
$totalMailboxes = $Mailboxes.Count
Write-Host "Found $totalMailboxes mailboxes. Let's fucking go." -ForegroundColor Green
# === PROCESS ===
$report = [System.Collections.Generic.List[object]]::new()
$progress = 0
$totalPrimaryGB = 0
$totalArchiveGB = 0
$totalItemCount = 0
foreach ($mbx in $Mailboxes) {
    $progress++
    $percent = [math]::Round(($progress / $totalMailboxes) * 100, 0)
    Write-Progress -Activity "Ripping mailbox stats" -Status "$progress of $totalMailboxes - $($mbx.DisplayName)" -PercentComplete $percent
    try {
        # PRIMARY STATS (modern EXO cmdlet)
        $primary = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction Stop
        $primarySizeGB = if ($primary.TotalItemSize.Value) {
            [math]::Round(($primary.TotalItemSize.Value.ToBytes() / 1GB), 2)
        } else { 0 }
        $primaryDeletedGB = if ($primary.TotalDeletedItemSize.Value) {
            [math]::Round(($primary.TotalDeletedItemSize.Value.ToBytes() / 1GB), 2)
        } else { 0 }
        $totalPrimaryGB += $primarySizeGB
        $totalItemCount += $primary.ItemCount
        # ARCHIVE
        $archiveSizeGB = 0
        $archiveItems = 0
        $archiveDeleted = 0
        if (-not $NoArchive -and $mbx.ArchiveDatabase) {
            try {
                $archive = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -Archive -ErrorAction Stop
                if ($archive) {
                    $archiveSizeGB = if ($archive.TotalItemSize.Value) {
                        [math]::Round(($archive.TotalItemSize.Value.ToBytes() / 1GB), 2)
                    } else { 0 }
                    $archiveItems = $archive.ItemCount
                    $archiveDeleted = $archive.DeletedItemCount
                    $totalArchiveGB += $archiveSizeGB
                }
            } catch {
                # Archive not ready or inaccessible - skip silently
            }
        }
        # QUOTAS (handle Unlimited properly)
        $maxQuotaGB = "Unlimited"
        if ($mbx.ProhibitSendReceiveQuota -and $mbx.ProhibitSendReceiveQuota.Value) {
            $maxQuotaGB = [math]::Round(($mbx.ProhibitSendReceiveQuota.Value.ToBytes() / 1GB), 2)
        }
        $warningGB = if ($mbx.IssueWarningQuota -and $mbx.IssueWarningQuota.Value) {
            [math]::Round(($mbx.IssueWarningQuota.Value.ToBytes() / 1GB), 2)
        } else { "N/A" }
        $freeSpace = if ($maxQuotaGB -ne "Unlimited") {
            [math]::Round(($maxQuotaGB - $primarySizeGB), 2)
        } else { "Unlimited" }
        $usagePct = if ($maxQuotaGB -ne "Unlimited" -and $maxQuotaGB -gt 0) {
            [math]::Round(($primarySizeGB / $maxQuotaGB * 100), 1)
        } else { "N/A" }
        $totalForMailbox = [math]::Round(($primarySizeGB + $archiveSizeGB), 2)
        $report.Add([PSCustomObject]@{
            DisplayName = $mbx.DisplayName
            UPN = $mbx.UserPrincipalName
            PrimarySmtp = $mbx.PrimarySmtpAddress
            Type = $mbx.RecipientTypeDetails
            PrimarySizeGB = $primarySizeGB
            PrimaryItems = $primary.ItemCount
            PrimaryDeletedGB = $primaryDeletedGB
            PrimaryDeletedItems = $primary.DeletedItemCount
            LastActivity = $primary.LastUserActionTime
            ArchiveSizeGB = $archiveSizeGB
            ArchiveItems = $archiveItems
            ArchiveDeletedItems = $archiveDeleted
            TotalSizeGB = $totalForMailbox
            MaxQuotaGB = $maxQuotaGB
            WarningQuotaGB = $warningGB
            FreeSpaceGB = $freeSpace
            UsagePercent = $usagePct
        })
    } catch {
        Write-Warning "Failed on $($mbx.DisplayName): $_"
    }
}
Write-Progress -Activity "Ripping mailbox stats" -Completed
# === EXPORT ===
$csvPath = "$OutputBase.csv"
$htmlPath = "$OutputBase.html"
$sortedReport = $report | Sort-Object TotalSizeGB -Descending
switch ($ExportFormat) {
    'CSV' {
        $sortedReport | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nCSV report saved to:`n$csvPath" -ForegroundColor Green
    }
    'HTML' {
        $htmlContent = $sortedReport | ConvertTo-Html -Title "EXO Mailbox Usage Report" -PreContent @"
<h1 style="color:#1f4e79; font-family:Segoe UI;">Exchange Online Mailbox Usage Report</h1>
<p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<p><strong>Mailboxes Processed:</strong> $totalMailboxes &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Primary:</strong> $totalPrimaryGB GB &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Archive:</strong> $totalArchiveGB GB &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Grand Total:</strong> <span style="color:#c00000; font-weight:bold;">$grandTotal GB</span></p>
"@ -PostContent "<hr><p style='color:#666; font-size:0.9em; font-family:Segoe UI;'>Generated by M365 Data Migration Scripts | EXOMailboxStats.ps1</p>"
        # Basic professional styling
        $htmlContent = $htmlContent -replace '<table>', '<table style="border-collapse:collapse; width:100%; font-family:Segoe UI, Arial, sans-serif; font-size:0.95em; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">'
        $htmlContent = $htmlContent -replace '<th>', '<th style="background-color:#1f4e79; color:white; padding:10px 8px; border:1px solid #ccc; text-align:left; position:sticky; top:0; z-index:10;">'
        $htmlContent = $htmlContent -replace '<td>', '<td style="padding:6px 8px; border:1px solid #ddd; vertical-align:top;">'
        $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8 -Force
        Write-Host "`nHTML report saved to:`n$htmlPath" -ForegroundColor Green
    }
    'Both' {
        $sortedReport | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        $htmlContent = $sortedReport | ConvertTo-Html -Title "EXO Mailbox Usage Report" -PreContent @"
<h1 style="color:#1f4e79; font-family:Segoe UI;">Exchange Online Mailbox Usage Report</h1>
<p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<p><strong>Mailboxes Processed:</strong> $totalMailboxes &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Primary:</strong> $totalPrimaryGB GB &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Archive:</strong> $totalArchiveGB GB &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Grand Total:</strong> <span style="color:#c00000; font-weight:bold;">$grandTotal GB</span></p>
"@ -PostContent "<hr><p style='color:#666; font-size:0.9em; font-family:Segoe UI;'>Generated by M365 Data Migration Scripts | EXOMailboxStats.ps1</p>"
        $htmlContent = $htmlContent -replace '<table>', '<table style="border-collapse:collapse; width:100%; font-family:Segoe UI, Arial, sans-serif; font-size:0.95em; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">'
        $htmlContent = $htmlContent -replace '<th>', '<th style="background-color:#1f4e79; color:white; padding:10px 8px; border:1px solid #ccc; text-align:left; position:sticky; top:0; z-index:10;">'
        $htmlContent = $htmlContent -replace '<td>', '<td style="padding:6px 8px; border:1px solid #ddd; vertical-align:top;">'
        $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8 -Force
        Write-Host "`nCSV report saved to:`n$csvPath" -ForegroundColor Green
        Write-Host "HTML report saved to:`n$htmlPath" -ForegroundColor Green
    }
}
# === BRUTAL SUMMARY (this is the "count up" part) ===
$grandTotal = [math]::Round(($totalPrimaryGB + $totalArchiveGB), 2)
$avgSize = if ($totalMailboxes -gt 0) { [math]::Round(($grandTotal / $totalMailboxes), 2) } else { 0 }
Write-Host "`n==============================================================" -ForegroundColor Magenta
Write-Host " EXCHANGE ONLINE MAILBOX USAGE SUMMARY" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Magenta
Write-Host "Mailboxes processed : $totalMailboxes" -ForegroundColor White
Write-Host "Primary storage used : $totalPrimaryGB GB" -ForegroundColor Green
Write-Host "Archive storage used : $totalArchiveGB GB" -ForegroundColor Yellow
Write-Host "GRAND TOTAL USED : $grandTotal GB" -ForegroundColor Red
Write-Host "Average per mailbox : $avgSize GB" -ForegroundColor White
Write-Host "==============================================================" -ForegroundColor Magenta
Write-Host "`nTop 10 biggest storage whores:" -ForegroundColor Red
$sortedReport | Select-Object -First 10 DisplayName, TotalSizeGB, Type, UsagePercent | Format-Table -AutoSize
# Cleanup
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "`nDisconnected. Go forth and delete some shit." -ForegroundColor DarkGray