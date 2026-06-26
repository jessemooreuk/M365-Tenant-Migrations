# M365-Tenant-Migrations

**M365 Data Migration Scripts** — Repository for Microsoft 365 Tenant Migrations

This private repository hosts PowerShell scripts and utilities designed to simplify and verify cross-tenant migrations in Microsoft 365 / Exchange Online.

## 📁 Repository Structure

- `Scripts/` — Ready-to-use migration and reporting scripts
- `README.md` — This file (you are here)

## 🚀 Available Scripts

### 1. Exchange Online Mailbox Size Comparison Report

**Script:** `Scripts/ExchangeOnline_MailboxSize_Comparison_Report.ps1`

**Description:**  
Performs a precise size diff/comparison between source and destination Exchange Online mailboxes for a list of **specific users** being migrated.

**Key Features:**
- Uses **destination UPNs as the lookup** (your input CSV)
- Automatically derives source UPN from the username prefix before `@`
- Connects sequentially to Source → Destination tenants (avoids multi-tenant session issues)
- Reports **Primary + Archive** sizes (MB), item counts, last logon times
- Calculates **diffs** (Dest − Source) with smart notes for large deltas or missing mailboxes
- Outputs professional CSV ready for Excel filtering and stakeholder reporting

**Ideal For:**
- Pre-migration baselining of specific user cohorts
- Mid-migration progress checks during staged/batched moves
- Post-cutover verification that content arrived intact
- Supplementing reports from BitTitan, Quest, AvePoint, or native MRS moves

**Usage Example:**

```powershell
.\Scripts\ExchangeOnline_MailboxSize_Comparison_Report.ps1 `
    -DestUPNsCsvPath ".\Phase1-Users.csv" `
    -SourceDomain "contoso.onmicrosoft.com"
```

See full `.SYNOPSIS` / help inside the script for all parameters and examples.

---

### 2. Tenant-Wide Mailbox Size Limits (150 MB)

**Script:** `Scripts/Set-TenantWideMailboxSizeLimits.ps1`

**Description:**  
Applies **150 MB** `MaxSendSize` and `MaxReceiveSize` to **every mailbox** in the tenant (or targeted via Custom Attributes).

**Key Features:**
- Queries all current mailboxes live (`Get-Mailbox`)
- No input CSV required (optional support for filtering)
- Includes progress bar and detailed CSV logging
- Supports `-WhatIf` for safe testing
- Handles errors gracefully with per-mailbox status

**Ideal For:**
- Standardizing message size limits across a tenant after migration
- Ensuring consistency for all users

**Usage Example:**

```powershell
# Preview changes first
.\Scripts\Set-TenantWideMailboxSizeLimits.ps1 -WhatIf

# Apply to all mailboxes
.\Scripts\Set-TenantWideMailboxSizeLimits.ps1
```

---

### 3. M365 User Activity & Service Access Report

**Script:** `Scripts/Get-M365UserActivityReport.ps1`

**Description:**  
Generates a report showing whether users have **successfully signed in** and **accessed** key Microsoft 365 services.

**Checks performed:**
- Last successful sign-in (Microsoft Entra ID)
- Last Exchange Online mailbox logon
- OneDrive storage usage

**Key Features:**
- Supports specific users via CSV **or** all users in the tenant
- Clear boolean columns (`HasSignedIn`, `HasUsedExchange`, `HasUsedOneDrive`)
- Professional CSV output with summary
- Useful for post-migration verification

**Ideal For:**
- Verifying users have logged in and started using services after migration
- Identifying active vs dormant accounts
- Stakeholder reporting

**Usage Example:**

```powershell
# Specific migration users
.\Scripts\Get-M365UserActivityReport.ps1 -UserUPNsCsvPath ".\MigrationUsers.csv"

# All users in tenant
.\Scripts\Get-M365UserActivityReport.ps1
```

> **Note:** Full Teams activity reporting is more complex and not included in this version.

---

### Other Scripts

- `Scripts/EXOMailboxStats.ps1` — Existing mailbox statistics collection script

## 📋 How to Use

1. Clone or download the repo
2. Review the script header (`.SYNOPSIS`) for parameters and examples
3. Run in PowerShell 7.x with the required admin accounts
4. Review the generated CSV reports in Excel

## 🔒 Security & Best Practices

- Scripts use modern interactive or certificate-based authentication where possible
- Never hard-code credentials
- Always test with `-WhatIf` first on production tenants
- Review reports carefully before making decisions

## 📝 Contributing

This is a living collection of practical migration scripts.  
Suggestions, improvements, and new script ideas are welcome via pull requests or issues.

---

**Maintained as part of the M365 Data Migration Scripts series.**

**Last Updated:** June 2026