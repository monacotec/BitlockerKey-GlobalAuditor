# Runbook — Deploy BitLocker Entra Escrow via Intune Remediations

Fleet-wide "fix" for devices flagged by `Get-DevicesMissingBitLockerKeys.ps1`.
This runbook packages the detection + remediation scripts into an **Intune
Remediation** so every non-compliant device escrows its BitLocker recovery key to
Entra ID automatically, on a schedule.

## The find → fix → verify loop

| Stage | Tool | Where it runs |
|---|---|---|
| **Find** | `Get-DevicesMissingBitLockerKeys.ps1` | Your workstation, against Graph (source of truth) |
| **Fix (detect)** | `Detect-BitLockerEntraEscrow.ps1` | Each device, as SYSTEM, via Intune |
| **Fix (remediate)** | `Backup-BitLockerToEntra.ps1` | Each device, as SYSTEM, via Intune (only when detection fails) |
| **Verify** | `Get-DevicesMissingBitLockerKeys.ps1` again | Your workstation — the missing-key count should drop |

The device-side detection is a *gate only*; it cannot confirm the key actually
reached Entra. The Graph auditor remains authoritative — always re-run it to
confirm a real drop in the missing-key count.

## How the pair converges

- **Detection** flags a device NON-COMPLIANT when a protected volume has no
  `RecoveryPassword` protector, **or** when there is no success stamp at
  `HKLM:\SOFTWARE\GI\BitLockerEscrow`.
- **Remediation** ensures a `RecoveryPassword` protector exists, escrows every
  protector with `BackupToAAD-BitLockerKeyProtector`, then writes the success
  stamp.
- On the next detection cycle the device reports COMPLIANT and remediation stops
  firing. Re-escrow is idempotent, so a repeat run is harmless.

## Prerequisites

- **Intune role**: Intune Administrator, or a custom role with
  *Remediations* create/assign rights.
- **Licensing**: Remediations require Windows Enterprise E3/E5, Education A3/A5,
  or Windows 365 (the usual Intune "Advanced" entitlement).
- **Devices**: Windows 10/11, Entra-joined or Hybrid Entra-joined, enrolled in
  Intune. BitLocker enabled (or an Intune disk-encryption policy that enables it).
- **BitLocker policy** (recommended, prevents *new* gaps): Endpoint security →
  Disk encryption → OS drive → `Save BitLocker recovery information to Microsoft
  Entra ID` = **Yes**; `Store recovery information in Microsoft Entra ID before
  enabling BitLocker` = **Require**.
- **Script files** from this repo:
  - `Detect-BitLockerEntraEscrow.ps1` (detection)
  - `Backup-BitLockerToEntra.ps1` (remediation)

## Step 1 — Create the remediation

1. Sign in to the **Microsoft Intune admin center** (intune.microsoft.com).
2. Go to **Devices → Remediations → Create script package**.
3. **Basics**
   - **Name**: `BitLocker — Escrow recovery key to Entra ID`
   - **Description**: `Detects volumes missing an Entra-escrowed BitLocker recovery key and escrows them. Companion to the BitLocker Global Auditor.`
   - **Publisher**: your team name.

## Step 2 — Attach the scripts and set run options

On the **Settings** page:

- **Detection script file**: upload `Detect-BitLockerEntraEscrow.ps1`
- **Remediation script file**: upload `Backup-BitLockerToEntra.ps1`

Then set these toggles exactly:

| Option | Value | Why |
|---|---|---|
| Run this script using the logged-on credentials | **No** | Must run as **SYSTEM** — `BackupToAAD-BitLockerKeyProtector` and `dsregcmd` need it |
| Enforce script signature check | **No** | Scripts are unsigned (set **Yes** only if you code-sign them first) |
| Run script in 64-bit PowerShell | **Yes** | Ensures the `HKLM:\SOFTWARE\GI` stamp is written/read in the same registry view |

## Step 3 — Assign and schedule

1. **Assignments**: start with a **pilot device group** (a ring of 10–50 test
   devices), not All Devices.
2. Edit the schedule for the assignment:
   - **Frequency**: **Daily** (recommended). Hourly is overkill — escrow is a
     one-time fix per device.
   - Detection runs on schedule; remediation runs immediately after, only on
     devices detection marked non-compliant.
3. **Review + create**.

## Step 4 — Monitor

- **Intune → Devices → Remediations →** *your package* → **Device status**.
  Columns to watch:
  - **Detection status** / **Remediation status** (Success / Failed)
  - **Pre-remediation detection output** — the `Non-compliant: ...` reason
  - **Post-remediation detection output** — should read `Compliant: ...`
- **On a device** (spot check): `Get-ItemProperty 'HKLM:\SOFTWARE\GI\BitLockerEscrow'`
  should show a recent `LastBackupUtc`.
- **Authoritative check**: re-run `Get-DevicesMissingBitLockerKeys.ps1`. The
  Missing BitLocker Keys count should fall as devices remediate. **This is the
  real success metric** — device-side "Compliant" only means the escrow call
  succeeded locally.

## Step 5 — Roll out in rings

Once the pilot shows the missing-key count dropping and no unexpected remediation
failures, broaden the assignment ring by ring (pilot → IT → broad → all). Keep the
daily schedule so newly-imaged or newly-encrypted devices self-heal.

## Rollback

Low risk — the scripts are non-destructive:
- Remove the assignment (or delete the package) to stop it running.
- The remediation only **adds** a `RecoveryPassword` protector and escrows keys;
  it never removes protectors, decrypts, or rotates existing keys.
- The only footprint left behind is the `HKLM:\SOFTWARE\GI\BitLockerEscrow` stamp,
  which is inert.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Detection always non-compliant, remediation "Success" | Stamp not being written (escrow raised a warning, not an error) | Check remediation output; confirm `HKLM:\SOFTWARE\GI\BitLockerEscrow\LastBackupUtc` exists after a run |
| Remediation fails: "not Azure AD / Entra joined" | Device registration broken or lagging | On device (as SYSTEM): `dsregcmd /status`; repair Entra/hybrid registration, then let it re-run |
| Escrow fails: access denied | Not running as SYSTEM, or 32-bit host | Confirm *logged-on credentials = No* and *64-bit PowerShell = Yes* |
| Keys appear in on-prem AD, not Entra | Hybrid device escrowing to AD DS by policy | Adjust BitLocker policy/join type; this script targets Entra explicitly via `BackupToAAD-...` |
| Detection Compliant but auditor still flags the device | Local stamp exists but key never reached Entra (registration issue at escrow time) | Trust the Graph auditor; clear the stamp to force a re-escrow: `Remove-Item 'HKLM:\SOFTWARE\GI\BitLockerEscrow' -Recurse` |

## Notes & limitations

- **Runtime is Windows PowerShell 5.1**, not PowerShell 7. Both scripts are
  5.1-compatible; do not add PS7-only syntax.
- **64-bit is required** so the registry stamp is consistent between detection and
  remediation.
- Device-side detection **cannot** verify Entra received the key. Pair every
  rollout with a Graph auditor run for true compliance numbers.
- Output captured by Intune is truncated (~2 KB per script) — keep script output
  concise (the scripts already do).
