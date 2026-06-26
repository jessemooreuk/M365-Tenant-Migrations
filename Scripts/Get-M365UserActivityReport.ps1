<#
.SYNOPSIS
    M365 Data Migration Script: User Activity & Service Access Report

.DESCRIPTION
    Generates a report showing whether users have successfully signed in and accessed 
    key Microsoft 365 services (Exchange Online, OneDrive, and general sign-in activity).

    Useful for:
    - Pre-migration baselining of active users
    - Post-migration verification that users have logged in and started using services
    - Identifying dormant accounts before/after tenant migration

    Checks performed:
    - Last successful sign-in (Microsoft Entra ID sign-in logs)
    - Last Exchange Online mailbox logon
    - OneDrive storage usage (via Microsoft Graph)

.PARAMETER UserUPNsCsvPath
    Optional. Path to CSV containing UPNs (column header: UPN).
    If omitted, the script processes ALL users in the tenant.

.PARAMETER OutputReportPath
    Optional. Path for the output CSV report.
    Default: .\M365_UserActivity_Report_YYYYMMDD_HHmmss.csv

.EXAMPLE
    # Process specific users from CSV
    .\Get-M365UserActivityReport.ps1 -UserUPNsCsvPath ".\MigrationUsers.csv"

.EXAMPLE
    # Process all users in the tenant (can be slow)
    .\Get-M365UserActivityReport.ps1

.NOTES
    Author      : Grok (xAI) - M365 Data Migration Scripts series
    Version     : 1.0 | 2026-06-26
    Requires    : 
        - ExchangeOnlineManagement module
        - Microsoft.Graph module (with appropriate permissions)
    Permissions : 
        - AuditLog.Read.All
        - User.Read.All
        - Directory.Read.All
        - Exchange admin rights for mailbox statistics

    Note on Teams: Full Teams activity reporting requires advanced Microsoft 365 usage reports 
    or Purview audit logs and is not included in this version for simplicity.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$UserUPNsCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputReportPath = ".\M365_UserActivity_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#==============================================================================================
# MODULE CHECK & CONNECTION
#==============================================================================================
$ErrorActionPreference = 'Stop'

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "  M365 DATA MIGRATION SCRIPT - User Activity & Service Access Report" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

# Check required modules
$RequiredModules = @('ExchangeOnlineManagement', 'Microsoft.Graph')
foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Host "[INFO] Installing module: $Module" -ForegroundColor Yellow
        Install-Module $Module -Scope CurrentUser -Force -AllowClobber
    }
}

Import-Module ExchangeOnlineManagement -Force
Import-Module Microsoft.Graph -Force

# Connect to Exchange Online
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline
}

# Connect to Microsoft Graph (requires interactive login with permissions)
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "AuditLog.Read.All", "User.Read.All", "Directory.Read.All" -NoWelcome

#==============================================================================================
# LOAD USERS
#==============================================================================================
$UsersToProcess = @()

if ($UserUPNsCsvPath) {
    if (-not (Test-Path $UserUPNsCsvPath)) {
        Write-Error "CSV file not found: $UserUPNsCsvPath"
        exit 1
    }
    $UsersToProcess = Import-Csv $UserUPNsCsvPath | Select-Object -ExpandProperty UPN
    Write-Host "Loaded $($UsersToProcess.Count) users from CSV." -ForegroundColor Green
} else {
    Write-Host "No CSV provided. Retrieving ALL users from tenant (this may take time)..." -ForegroundColor Yellow
    $UsersToProcess = Get-MgUser -All -Property UserPrincipalName, DisplayName | 
                      Select-Object -ExpandProperty UserPrincipalName
    Write-Host "Found $($UsersToProcess.Count) users in tenant." -ForegroundColor Green
}

#==============================================================================================
# PROCESS EACH USER
#==============================================================================================
$Report = [System.Collections.Generic.List[object]]::new()
$Counter = 0

foreach ($UPN in $UsersToProcess) {
    $Counter++
    $Percent = [math]::Round(($Counter / $UsersToProcess.Count) * 100, 1)
    Write-Progress -Activity "Analyzing user activity" `
                   -Status "$Counter of $($UsersToProcess.Count) - $UPN" `
                   -PercentComplete $Percent

    $Result = [PSCustomObject]@{
        UPN                        = $UPN
        DisplayName                = ""
        LastSuccessfulSignIn       = ""
        LastExchangeLogon          = ""
        OneDriveStorageUsedMB      = 0
        HasOneDrive                = $false
        HasSignedIn                = $false
        HasUsedExchange            = $false
        HasUsedOneDrive            = $false
        Notes                      = ""
    }

    try {
        # Get basic user info
        $User = Get-MgUser -UserId $UPN -Property DisplayName, UserPrincipalName -ErrorAction SilentlyContinue
        if ($User) {
            $Result.DisplayName = $User.DisplayName
        }

        # === Last Successful Sign-in (Entra ID) ===
        try {
            $SignIns = Get-MgAuditLogSignIn -Filter "userPrincipalName eq '$UPN' and status/errorCode eq 0" `
                                            -Top 1 -Sort "createdDateTime desc" -ErrorAction SilentlyContinue
            if ($SignIns) {
                $Result.LastSuccessfulSignIn = $SignIns[0].CreatedDateTime
                $Result.HasSignedIn = $true
            }
        } catch {
            $Result.Notes += "Sign-in logs unavailable; "
        }

        # === Last Exchange Logon ===
        try {
            $MailboxStats = Get-MailboxStatistics -Identity $UPN -ErrorAction SilentlyContinue
            if ($MailboxStats -and $MailboxStats.LastLogonTime) {
                $Result.LastExchangeLogon = $MailboxStats.LastLogonTime
                $Result.HasUsedExchange = $true
            }
        } catch {
            $Result.Notes += "Exchange stats unavailable; "
        }

        # === OneDrive Usage (via Microsoft Graph) ===
        try {
            $Drive = Get-MgUserDrive -UserId $UPN -ErrorAction SilentlyContinue
            if ($Drive -and $Drive.Quota -and $Drive.Quota.Used) {
                $Result.OneDriveStorageUsedMB = [math]::Round($Drive.Quota.Used / 1MB, 2)
                $Result.HasOneDrive = $true
                $Result.HasUsedOneDrive = ($Result.OneDriveStorageUsedMB -gt 0)
            }
        } catch {
            $Result.Notes += "OneDrive data unavailable; "
        }

    } catch {
        $Result.Notes += "General error: $($_.Exception.Message); "
    }

    $Report.Add($Result)
}

Write-Progress -Activity "Analyzing user activity" -Completed

#==============================================================================================
# EXPORT REPORT
#==============================================================================================
$Report | Export-Csv -Path $OutputReportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n==================================================================" -ForegroundColor Cyan
Write-Host "                        REPORT COMPLETE" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "Total users processed : $($Report.Count)"
Write-Host "Report saved to       : $OutputReportPath" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor Yellow
Write-Host "  Users who have signed in        : $(($Report | Where-Object HasSignedIn -eq $true).Count)"
Write-Host "  Users with Exchange activity    : $(($Report | Where-Object HasUsedExchange -eq $true).Count)"
Write-Host "  Users with OneDrive data        : $(($Report | Where-Object HasUsedOneDrive -eq $true).Count)"
Write-Host "==================================================================" -ForegroundColor Cyan

Write-Host "`nTip: Open the CSV in Excel and filter on 'HasSignedIn', 'HasUsedExchange', or 'HasUsedOneDrive' columns." -ForegroundColor Cyan
Write-Host "Script completed successfully." -ForegroundColor Green
# End of script