# BitlockerKey-GlobalAuditor

Audits a hybrid Entra ID tenant for Windows devices that do not have a BitLocker recovery key escrowed to Entra.

## Features

- Queries all Windows devices (or hybrid-joined only) via Microsoft Graph
- Checks each device for BitLocker recovery key presence
- Flags devices as Active, Inactive, or No Sign-In Data
- Reports registered device owner and UPN for remediation
- Tracks BitLocker key count per device (useful for multi-volume machines)
- Exports a formatted XLSX workbook with three worksheets:
  - **Missing BitLocker Keys** - devices with no key escrowed
  - **Compliant Devices** - devices with keys escrowed
  - **Summary** - compliance rates, device counts, active vs inactive breakdown
- Logs all output to `C:\GI\BitLockerAudit_<timestamp>.log`

## Prerequisites

Run the prerequisite installer first:

```powershell
.\Install-BitLockerAuditPrereqs.ps1
```

This installs/updates the required PowerShell modules:
- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Identity.DirectoryManagement`
- `Microsoft.Graph.Identity.SignIns`
- `ImportExcel`

## Usage

```powershell
# Audit all Windows devices (default: flag inactive after 90 days)
.\Get-DevicesMissingBitLockerKeys.ps1

# Hybrid Entra-joined devices only
.\Get-DevicesMissingBitLockerKeys.ps1 -HybridOnly

# Custom inactive threshold and output path
.\Get-DevicesMissingBitLockerKeys.ps1 -InactiveDays 30 -ExportPath "C:\GI\report.xlsx"

# Specific tenant
.\Get-DevicesMissingBitLockerKeys.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

## Required Graph Permissions

- `Device.Read.All`
- `BitLockerKey.Read.All`

## Output

- **XLSX report**: `C:\GI\BitLockerAudit_<timestamp>.xlsx`
- **Log file**: `C:\GI\BitLockerAudit_<timestamp>.log`

## Remediation

The auditor finds *which* devices are missing a key; `Backup-BitLockerToEntra.ps1`
is the *fix* half. Run it on an affected device (as SYSTEM via PDQ Connect, or as
Administrator) to escrow the recovery key to Entra ID:

```powershell
.\Backup-BitLockerToEntra.ps1
```

It enumerates encrypted volumes, ensures each has a `RecoveryPassword` protector
(adding one if missing), and backs each protector up to Entra with
`BackupToAAD-BitLockerKeyProtector`. It self-verifies Hybrid Entra join via
`dsregcmd` first and fails loudly if the device isn't registered.

### Why keys go missing

- **TPM-only protector** — silent/Autopilot encryption often creates only a `Tpm`
  protector and no `RecoveryPassword`, so there was never a key to escrow. This is
  the most common cause; the script handles it by adding a recovery password first.
- **Encrypted before policy applied** — a `RecoveryPassword` exists but the
  backup-to-Entra step never ran (manual encryption, or the Intune policy landed
  after encryption). Re-running the escrow fixes it.
- **Join-type mismatch** — Entra-joined devices escrow to Entra ID; Hybrid-joined
  devices default to on-prem AD DS. If the fleet is hybrid, keys may be in AD DS
  instead of Entra.

### Intune policy (prevention, going forward)

Endpoint security → Disk encryption → BitLocker policy → OS drive:

- `Save BitLocker recovery information to Microsoft Entra ID` = **Yes**
- `Store recovery information in Microsoft Entra ID before enabling BitLocker` = **Require**
- Ensure a recovery password protector is created (not TPM-only)

Policy only helps devices going forward. Already-encrypted devices need the active
remediation above, ideally delivered fleet-wide via **Intune → Devices →
Remediations** (a detection script flagging "no `RecoveryPassword` protector on C:"
plus `Backup-BitLockerToEntra.ps1` as the remediation). Note that Intune remediation
scripts run as **SYSTEM under Windows PowerShell 5.1**, not PowerShell 7 — the
BitLocker cmdlets are available there, but don't rely on PS7-only syntax.
