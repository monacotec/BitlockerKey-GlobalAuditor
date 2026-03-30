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
