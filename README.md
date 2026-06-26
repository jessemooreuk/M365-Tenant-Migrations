# M365-Tenant-Migrations

**M365 Data Migration Scripts** — Repository for Microsoft 365 Tenant Migrations

This private repository hosts PowerShell scripts and utilities designed to simplify and verify cross-tenant migrations in Microsoft 365 / Exchange Online.

## 📁 Repository Structure

- `Scripts/` — Ready-to-use migration and reporting scripts
- `README.md` — This file (you are here)

## 🚀 Available Scripts

### Exchange Online Mailbox Size Comparison Report

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

See full `.SYNOPSIS` / help inside the script for all parameters and examples (`Get-Help .\Scripts\ExchangeOnline_MailboxSize_Comparison_Report.ps1 -Full`).

**Prerequisites:** PowerShell 5.1+ (7.x recommended), ExchangeOnlineManagement module, appropriate read permissions in both tenants.

---

### Other Scripts

- `Scripts/EXOMailboxStats.ps1` — Existing mailbox statistics collection script

## 📋 How to Use

1. Clone or download the repo
2. Review script header for parameters and examples
3. Run in an elevated PowerShell session with the required admin accounts
4. Review generated CSV reports in Excel (filter on `Total_Diff_MB`, `Notes`, etc.)

## 🔒 Security & Best Practices

- Scripts use modern interactive or certificate-based auth where possible
- Never hard-code credentials
- Test in a pilot tenant/user set first
- Review all diffs manually before declaring migration complete

## 📝 Contributing

This is a living collection of battle-tested migration scripts.  
Feel free to suggest improvements, new scripts, or report issues via pull requests or discussions.

---

**Maintained as part of the M365 Data Migration Scripts series.**

**Last Updated:** June 2026