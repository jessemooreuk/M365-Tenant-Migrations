<# 
.SYNOPSIS
    M365 Data Migration Script: Exchange Online Mailbox Size Comparison Report (Cross-Tenant)

.DESCRIPTION
    Performs a size diff/comparison between source and destination Exchange Online mailboxes 
    for a list of specific users being migrated.
    
    - Input: CSV containing Destination UPNs (one per row under "DestUPN" header)
    - Lookup logic: Extracts the local part (before @) from each Dest UPN and constructs 
      Source UPN as <localpart>@<SourceDomain> (you provide the source domain).
    - Connects SEQUENTIALLY to Source tenant first, then Destination tenant (avoids 
      multi-tenant session conflicts common with ExchangeOnlineManagement module).
    - Retrieves Primary + Archive mailbox statistics (Total Size in MB, Item Count, Last Logon).
    - Calculates differences (Dest - Source). Positive diff = more data on destination 
      (expected after migration due to new mail or sync delta).
    - Outputs a detailed CSV report ready for Excel analysis + console summary.

    Ideal for:
    - Pre-migration baselining
    - Mid-migration progress checks (staged/batched migrations)
    - Post-cutover verification that content arrived intact
    - Supplementing 3rd-party migration tool reports (BitTitan, Quest, AvePoint, etc.)

.PARAMETER DestUPNsCsvPath
    [Mandatory] Full or relative path to input CSV. Must contain a column named "DestUPN".

.PARAMETER SourceDomain
    [Mandatory] The domain suffix used in the SOURCE tenant for these users 
    (e.g. "contoso.onmicrosoft.com" or "olddomain.com"). 
    Script builds SourceUPN = (Dest local part) + "@" + SourceDomain

.PARAMETER SourceAdminUPN
    Optional. UPN of an admin account in the SOURCE tenant. If omitted, script prompts.

.PARAMETER DestAdminUPN
    Optional. UPN of an admin account in the DESTINATION tenant. If omitted, script prompts.

.PARAMETER ReportOutputPath
    Optional. Full path for the output CSV report. 
    Default: .\Migration_Size_Comparison_YYYYMMDD_HHmmss.csv in current directory.

.EXAMPLE
    # Basic interactive run (recommended for ad-hoc specific user migrations)
    .\ExchangeOnline_MailboxSize_Comparison_Report.ps1 `
        -DestUPNsCsvPath "C:\Migration\Phase1-Users.csv" `
        -SourceDomain "contoso.onmicrosoft.com"

.EXAMPLE
    # Fully parameterized (good for scheduled or documented runs)
    .\ExchangeOnline_MailboxSize_Comparison_Report.ps1 `
        -DestUPNsCsvPath ".\users.csv" `
        -SourceDomain "fabrikam.com" `
        -SourceAdminUPN "migadmin@contoso.onmicrosoft.com" `
        -DestAdminUPN "migadmin@fabrikam.com" `
        -ReportOutputPath "C:\Reports\Phase2_SizeDiff_2026-06-26.csv"

.NOTES
    Author          : Grok (xAI) - M365 Data Migration Scripts series
    Version         : 1.0 | 2026-06-26
    Requires        : PowerShell 5.1+ (7.x preferred) + ExchangeOnlineManagement module
    Permissions     : Source & Dest tenants - accounts need Get-Mailbox + Get-MailboxStatistics 
                      (View-Only Organization Management or custom RBAC role is sufficient)
    Multi-tenant    : Script deliberately connects/disconnects sequentially. Do NOT run 
                      Connect-ExchangeOnline to both tenants at the same time in same session.
    Error handling  : Missing mailboxes reported as "N/A" with reason in Notes column.
    Throttling      : For < 500-1000 specific users this is fast. Larger lists may need 
                      added Start-Sleep or batching.
    Archive handling: Archive stats only collected when ArchiveStatus = Active or ArchiveDatabase present.
    Size accuracy   : Uses robust MB conversion (handles ByteQuantifiedSize + string fallback).

    DISCLAIMER: This is a custom reporting utility for migration verification. 
    It does not perform migration. Always validate with official Microsoft reports / 
    your migration platform. Test thoroughly in a pilot before production use.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to CSV file with DestUPN column")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$DestUPNsCsvPath,

    [Parameter(Mandatory = $true, HelpMessage = "Source tenant domain suffix (e.g. contoso.onmicrosoft.com)")]
    [string]$SourceDomain,

    [Parameter(Mandatory = $false)]
    [string]$SourceAdminUPN,

    [Parameter(Mandatory = $false)]
    [string]$DestAdminUPN,

    [Parameter(Mandatory = $false)]
    [string]$ReportOutputPath = ".\Migration_Size_Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#==============================================================================================
# PRE-REQUISITES & MODULE HANDLING
#==============================================================================================
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "  M365 DATA MIGRATION SCRIPT - Exchange Online Size Comparison" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray

# Ensure ExchangeOnlineManagement module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "`n[INFO] ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[OK] Module installed successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install ExchangeOnlineManagement module. Please install manually: Install-Module ExchangeOnlineManagement"
        exit 1
    }
}

Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
Write-Host "[OK] ExchangeOnlineManagement v$((Get-Module ExchangeOnlineManagement).Version) loaded." -ForegroundColor Green

#==============================================================================================
# HELPER FUNCTIONS
#==============================================================================================
function Get-MailboxSizeInMB {
    <#
    .SYNOPSIS
        Robustly converts Exchange TotalItemSize (ByteQuantifiedSize or string) to MB (rounded 2 decimals).
    #>
    param(
        [Parameter(Mandatory = $false)]
        $TotalItemSize
    )
    if ($null -eq $TotalItemSize) { return 0 }

    # Try native ByteQuantifiedSize methods first
    try {
        if ($TotalItemSize -is [Microsoft.Exchange.Data.ByteQuantifiedSize]) {
            return [math]::Round($TotalItemSize.ToMB(), 2)
        }
        if ($TotalItemSize.PSObject.Methods.Name -contains 'ToMB') {
            return [math]::Round($TotalItemSize.ToMB(), 2)
        }
    } catch {
        # Fall through to string parsing
    }

    # Reliable string fallback used across migration community scripts
    $sizeStr = $TotalItemSize.ToString()
    if ($sizeStr -match '\(([\d,]+)\s*bytes\)') {
        try {
            $bytes = [long]($Matches[1] -replace ',', '')
            return [math]::Round($bytes / 1MB, 2)
        } catch { return 0 }
    }
    if ($sizeStr -match '([\d.]+)\s*GB') {
        return [math]::Round([double]$Matches[1] * 1024, 2)
    }
    if ($sizeStr -match '([\d.]+)\s*MB') {
        return [math]::Round([double]$Matches[1], 2)
    }
    return 0
}

function Get-MailboxStatsSafe {
    <#
    .SYNOPSIS
        Safely retrieves mailbox statistics (primary or archive) with error handling.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        [switch]$Archive
    )
    try {
        $params = @{
            Identity      = $Identity
            ErrorAction   = 'Stop'
        }
        if ($Archive) { $params['Archive'] = $true }

        $stats = Get-MailboxStatistics @params
        if (-not $stats) { return $null }

        return [PSCustomObject]@{
            SizeMB        = Get-MailboxSizeInMB -TotalItemSize $stats.TotalItemSize
            ItemCount     = if ($stats.ItemCount) { $stats.ItemCount } else { 0 }
            LastLogonTime = if ($stats.LastLogonTime) { $stats.LastLogonTime } else { "Never" }
        }
    } catch {
        Write-Warning "Stats retrieval failed for '$Identity' $(if ($Archive) {'(Archive)'}): $($_.Exception.Message)"
        return $null
    }
}

function Get-ArchiveStatus {
    <#
    .SYNOPSIS
        Checks if archive is active for a mailbox without throwing on missing mailbox.
    #>
    param([string]$Identity)
    try {
        $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop | 
               Select-Object ArchiveStatus, ArchiveDatabase, ArchiveState
        $hasArchive = ($mbx.ArchiveStatus -eq 'Active') -or 
                      ($mbx.ArchiveDatabase) -or 
                      ($mbx.ArchiveState -eq 'Local')
        return $hasArchive
    } catch {
        return $false
    }
}

#==============================================================================================
# LOAD INPUT LIST
#==============================================================================================
Write-Host "`n[STEP 1] Loading destination UPN list from: $DestUPNsCsvPath" -ForegroundColor Yellow

$csvData = Import-Csv -Path $DestUPNsCsvPath
if (-not $csvData) {
    Write-Error "CSV file is empty or invalid."
    exit 1
}

# Support common header variations
$destUPNs = $csvData | 
    Where-Object { $_ -and ($_.DestUPN -or $_.'Dest UPN' -or $_.'UPN' -or $_.'Destination UPN') } |
    ForEach-Object { 
        $upn = $_.DestUPN
        if (-not $upn) { $upn = $_.'Dest UPN' }
        if (-not $upn) { $upn = $_.'UPN' }
        if (-not $upn) { $upn = $_.'Destination UPN' }
        if ($upn -and $upn -match '@') { $upn.Trim() }
    } | 
    Where-Object { $_ } |
    Sort-Object -Unique

if (-not $destUPNs -or $destUPNs.Count -eq 0) {
    Write-Error "No valid DestUPN values found. CSV must have a column named DestUPN (or Dest UPN / UPN / Destination UPN)."
    exit 1
}

Write-Host "[OK] Found $($destUPNs.Count) unique destination UPN(s)." -ForegroundColor Green

#==============================================================================================
# SOURCE TENANT - COLLECT STATS
#==============================================================================================
Write-Host "`n[STEP 2] Connecting to SOURCE Exchange Online tenant..." -ForegroundColor Yellow

if (-not $SourceAdminUPN) {
    $SourceAdminUPN = Read-Host "Enter SOURCE tenant admin UPN (e.g. migadmin@contoso.onmicrosoft.com)"
}

try {
    Connect-ExchangeOnline -UserPrincipalName $SourceAdminUPN -ShowProgress $true -ErrorAction Stop
    Write-Host "[OK] Connected to SOURCE tenant as $SourceAdminUPN" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to SOURCE tenant: $($_.Exception.Message)"
    exit 1
}

$sourceStats = @{}
$counter = 0
foreach ($destUPN in $destUPNs) {
    $counter++
    $percent = [math]::Round(($counter / $destUPNs.Count) * 100, 1)
    Write-Progress -Activity "Collecting SOURCE tenant statistics" `
                   -Status "Processing $counter of $($destUPNs.Count) ($percent%)" `
                   -PercentComplete $percent `
                   -CurrentOperation $destUPN

    $localPart = ($destUPN -split '@')[0]
    $sourceUPN = "$localPart@$SourceDomain"

    $primary = Get-MailboxStatsSafe -Identity $sourceUPN
    $archive = $null
    if (Get-ArchiveStatus -Identity $sourceUPN) {
        $archive = Get-MailboxStatsSafe -Identity $sourceUPN -Archive
    }

    $sourceStats[$destUPN] = [PSCustomObject]@{
        SourceUPN              = $sourceUPN
        SourcePrimarySizeMB    = if ($primary) { $primary.SizeMB } else { "N/A" }
        SourcePrimaryItems     = if ($primary) { $primary.ItemCount } else { "N/A" }
        SourceArchiveSizeMB    = if ($archive) { $archive.SizeMB } else { 0 }
        SourceArchiveItems     = if ($archive) { $archive.ItemCount } else { 0 }
        SourceLastLogon        = if ($primary) { $primary.LastLogonTime } else { "N/A" }
    }
}

Write-Progress -Activity "Collecting SOURCE tenant statistics" -Completed
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "[OK] SOURCE data collected. Disconnected from source tenant." -ForegroundColor Green

#==============================================================================================
# DESTINATION TENANT - COLLECT STATS + BUILD REPORT
#==============================================================================================
Write-Host "`n[STEP 3] Connecting to DESTINATION Exchange Online tenant..." -ForegroundColor Yellow

if (-not $DestAdminUPN) {
    $DestAdminUPN = Read-Host "Enter DESTINATION tenant admin UPN (e.g. migadmin@fabrikam.com)"
}

try {
    Connect-ExchangeOnline -UserPrincipalName $DestAdminUPN -ShowProgress $true -ErrorAction Stop
    Write-Host "[OK] Connected to DESTINATION tenant as $DestAdminUPN" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to DESTINATION tenant: $($_.Exception.Message)"
    exit 1
}

$report = [System.Collections.Generic.List[object]]::new()
$counter = 0

foreach ($destUPN in $destUPNs) {
    $counter++
    $percent = [math]::Round(($counter / $destUPNs.Count) * 100, 1)
    Write-Progress -Activity "Collecting DESTINATION stats & building comparison report" `
                   -Status "Processing $counter of $($destUPNs.Count) ($percent%)" `
                   -PercentComplete $percent `
                   -CurrentOperation $destUPN

    $src = $sourceStats[$destUPN]
    if (-not $src) {
        $src = [PSCustomObject]@{
            SourceUPN = "N/A"; SourcePrimarySizeMB = "N/A"; SourcePrimaryItems = "N/A"
            SourceArchiveSizeMB = 0; SourceArchiveItems = 0; SourceLastLogon = "N/A"
        }
    }

    $primary = Get-MailboxStatsSafe -Identity $destUPN
    $archive = $null
    if (Get-ArchiveStatus -Identity $destUPN) {
        $archive = Get-MailboxStatsSafe -Identity $destUPN -Archive
    }

    $destPrimarySizeMB   = if ($primary) { $primary.SizeMB } else { "N/A" }
    $destPrimaryItems    = if ($primary) { $primary.ItemCount } else { "N/A" }
    $destArchiveSizeMB   = if ($archive) { $archive.SizeMB } else { 0 }
    $destArchiveItems    = if ($archive) { $archive.ItemCount } else { 0 }
    $destLastLogon       = if ($primary) { $primary.LastLogonTime } else { "N/A" }

    # Calculate diffs (only when both sides have numeric values)
    $primaryDiff = if ( ($destPrimarySizeMB -is [double] -or $destPrimarySizeMB -is [int]) -and 
                        ($src.SourcePrimarySizeMB -is [double] -or $src.SourcePrimarySizeMB -is [int]) ) {
        [math]::Round($destPrimarySizeMB - $src.SourcePrimarySizeMB, 2)
    } else { "N/A" }

    $archiveDiff = if ( ($destArchiveSizeMB -is [double] -or $destArchiveSizeMB -is [int]) -and 
                        ($src.SourceArchiveSizeMB -is [double] -or $src.SourceArchiveSizeMB -is [int]) ) {
        [math]::Round($destArchiveSizeMB - $src.SourceArchiveSizeMB, 2)
    } else { "N/A" }

    $srcTotal  = if ($src.SourcePrimarySizeMB -is [double] -or $src.SourcePrimarySizeMB -is [int]) { $src.SourcePrimarySizeMB } else { 0 } +
                 if ($src.SourceArchiveSizeMB -is [double] -or $src.SourceArchiveSizeMB -is [int]) { $src.SourceArchiveSizeMB } else { 0 }
    $destTotal = if ($destPrimarySizeMB -is [double] -or $destPrimarySizeMB -is [int]) { $destPrimarySizeMB } else { 0 } +
                 if ($destArchiveSizeMB -is [double] -or $destArchiveSizeMB -is [int]) { $destArchiveSizeMB } else { 0 }

    $totalDiff = if ( ($destTotal -is [double] -or $destTotal -is [int]) -and 
                      ($srcTotal -is [double] -or $srcTotal -is [int]) ) {
        [math]::Round($destTotal - $srcTotal, 2)
    } else { "N/A" }

    $notes = @()
    if ($src.SourcePrimarySizeMB -eq "N/A") { $notes += "Source mailbox not found or inaccessible" }
    if ($destPrimarySizeMB -eq "N/A")       { $notes += "Destination mailbox not found or inaccessible" }
    if ($primaryDiff -ne "N/A" -and [math]::Abs($primaryDiff) -gt 50) { 
        $notes += "Large primary size delta (>50MB) - review sync status" 
    }

    $report.Add([PSCustomObject]@{
        DestUPN                = $destUPN
        SourceUPN              = $src.SourceUPN
        Source_Primary_Size_MB = $src.SourcePrimarySizeMB
        Source_Primary_Items   = $src.SourcePrimaryItems
        Source_Archive_Size_MB = $src.SourceArchiveSizeMB
        Source_Archive_Items   = $src.SourceArchiveItems
        Source_Total_Size_MB   = [math]::Round($srcTotal, 2)
        Dest_Primary_Size_MB   = $destPrimarySizeMB
        Dest_Primary_Items     = $destPrimaryItems
        Dest_Archive_Size_MB   = $destArchiveSizeMB
        Dest_Archive_Items     = $destArchiveItems
        Dest_Total_Size_MB     = [math]::Round($destTotal, 2)
        Primary_Diff_MB        = $primaryDiff
        Archive_Diff_MB        = $archiveDiff
        Total_Diff_MB          = $totalDiff
        Source_LastLogon       = $src.SourceLastLogon
        Dest_LastLogon         = $destLastLogon
        Notes                  = ($notes -join "; ")
    })
}

Write-Progress -Activity "Collecting DESTINATION stats & building comparison report" -Completed
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "[OK] DESTINATION data collected. Disconnected from destination tenant." -ForegroundColor Green

#==============================================================================================
# EXPORT REPORT
#==============================================================================================
Write-Host "`n[STEP 4] Exporting comparison report..." -ForegroundColor Yellow

try {
    $report | Export-Csv -Path $ReportOutputPath -NoTypeInformation -Encoding UTF8 -Force
    $fullReportPath = (Resolve-Path $ReportOutputPath).Path
    Write-Host "[SUCCESS] Report saved to: $fullReportPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to export report: $($_.Exception.Message)"
    exit 1
}

#==============================================================================================
# SUMMARY
#==============================================================================================
Write-Host "`n==================================================================" -ForegroundColor Cyan
Write-Host "                        REPORT SUMMARY" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

$totalUsers       = $report.Count
$sourceFound      = ($report | Where-Object { $_.Source_Primary_Size_MB -ne "N/A" }).Count
$destFound        = ($report | Where-Object { $_.Dest_Primary_Size_MB -ne "N/A" }).Count
$bothFound        = ($report | Where-Object { 
    $_.Source_Primary_Size_MB -ne "N/A" -and $_.Dest_Primary_Size_MB -ne "N/A" 
}).Count

$avgTotalDiff = if ($bothFound -gt 0) {
    $numericDiffs = $report | Where-Object { $_.Total_Diff_MB -ne "N/A" } | Select-Object -ExpandProperty Total_Diff_MB
    if ($numericDiffs) { [math]::Round( ($numericDiffs | Measure-Object -Average).Average , 2) } else { "N/A" }
} else { "N/A" }

Write-Host "Total users in scope          : $totalUsers"
Write-Host "Source mailboxes found        : $sourceFound / $totalUsers"
Write-Host "Destination mailboxes found   : $destFound / $totalUsers"
Write-Host "Both sides found (comparable) : $bothFound / $totalUsers"
Write-Host "Average Total Size Diff (MB)  : $avgTotalDiff  (Dest - Source)"
Write-Host ""
Write-Host "Report file                   : $fullReportPath" -ForegroundColor White
Write-Host "End Time                      : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "==================================================================" -ForegroundColor Cyan

# Quick preview
Write-Host "`nFirst 5 rows preview (open full CSV in Excel for analysis):" -ForegroundColor Yellow
$report | Select-Object -First 5 | 
    Format-Table DestUPN, SourceUPN, Source_Total_Size_MB, Dest_Total_Size_MB, Total_Diff_MB, Notes -AutoSize

Write-Host "`n[INFO] Open the CSV in Microsoft Excel and filter/sort on:" -ForegroundColor Cyan
Write-Host "       - Total_Diff_MB (large negative = possible incomplete migration)" -ForegroundColor Cyan
Write-Host "       - Notes column for exceptions" -ForegroundColor Cyan
Write-Host "       - LastLogon columns to check activity" -ForegroundColor Cyan

Write-Host "`nScript completed successfully. Happy migrating!" -ForegroundColor Green
# End of script