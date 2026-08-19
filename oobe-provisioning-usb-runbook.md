# Runbook — Build the OOBE Provisioning USB (.ppkg)

Create a USB provisioning package that, applied during **OOBE**, joins a new or
reimaged device to the directory and enrolls it in Intune — so the BitLocker
disk-encryption policy and the Entra-escrow remediation take over automatically and
the device is **compliant from day one**. This is the "front of the funnel" for the
BitLocker Global Auditor: stop new gaps before they happen.

## Where this fits

| Stage | Tool | This runbook's role |
|---|---|---|
| **Provision** | `.ppkg` on USB (this runbook) | Get the device joined + Intune-enrolled at OOBE |
| **Encrypt + escrow** | Intune BitLocker disk-encryption policy | Enables BitLocker and escrows the key to Entra |
| **Backfill gaps** | Intune Remediations pair ([intune-remediation-runbook.md](intune-remediation-runbook.md)) | Fixes any device that slipped through |
| **Verify** | `Get-DevicesMissingBitLockerKeys.ps1` (Graph) | Authoritative compliance check |

> **The .ppkg does not escrow BitLocker keys itself.** It only joins and enrolls the
> device. Escrow is done by the Intune disk-encryption policy once the device checks
> in. Both pieces are required for a day-one-compliant device.

## Decision — join type (read first)

A provisioning package can join the device one of two ways. **This changes where the
BitLocker key escrows**, so choose deliberately:

| Join type | What the .ppkg does | BitLocker escrow target | Use when |
|---|---|---|---|
| **Entra join (cloud)** | Uses a *bulk enrollment token* to join Entra directly; device is Entra-joined (cloud-only) | Entra ID (clean, direct) | You want the simplest path to keys-in-Entra; device doesn't need on-prem domain membership |
| **AD domain join (hybrid)** | Joins on-prem AD with domain credentials; device hybrid-registers to Entra via Entra Connect | AD DS by default; Entra only if the Intune policy is set to save to Entra | Device must be on-prem domain-joined (GPO, on-prem resources) |

Your current fleet is **hybrid**. If the goal is simply to get recovery keys into
Entra with the least friction, **cloud Entra join is the cleaner target**. If these
devices must remain on-prem domain-joined, use AD domain join **and** confirm the
Intune BitLocker policy saves recovery info to Entra (see the escrow policy note
below), or keys will land only in on-prem AD DS.

> **Chosen path: AD domain join (hybrid).** Devices join on-prem AD using a
> dedicated join account, `hybridjoin@gipartners.com`, then hybrid-register to Entra
> via Entra Connect. Because the Intune BitLocker policy is configured to save
> recovery info to Entra, keys escrow to Entra ID. Set up the join account per the
> next section **before** building the package.

## Join account — hybridjoin@gipartners.com

This is an **on-prem Active Directory** account whose only job is to create the
computer object during domain join. Its credentials are **embedded in the .ppkg**,
so it must be least-privileged and the package treated as a secret.

### Permissions summary

| Scope | Needs | Explicitly does NOT need |
|---|---|---|
| **On-prem AD** | Delegated **Create/Delete Computer objects** + specific property writes on **one target OU** | Domain Admin, Account Operators, or any built-in admin group |
| **Entra ID** | *Nothing* — hybrid registration is done by Entra Connect, not this account | Any Entra role (Global Admin, Cloud Device Admin, Intune Admin) |
| **Intune** | *Nothing* — auto-enrollment runs in the device/user context after hybrid join | Any Intune role or license |

Least privilege matters here specifically because the password ships inside the
package. Delegating rights on a single OU means a leaked package can, at worst,
create/rename computer objects in that one OU — not compromise the domain.

### Step A — Create the account

Run on a machine with the RSAT Active Directory PowerShell module (adjust the OU DN
to your directory):

```powershell
New-ADUser -Name 'hybridjoin' -SamAccountName 'hybridjoin' `
  -UserPrincipalName 'hybridjoin@gipartners.com' `
  -Path 'OU=Service Accounts,DC=gipartners,DC=com' `
  -AccountPassword (Read-Host -AsSecureString 'Enter a strong 25+ char password') `
  -Enabled $true -PasswordNeverExpires $true -CannotChangePassword $true `
  -Description 'Delegated computer-join account for OOBE provisioning packages'
```

Keep it a plain **Domain Users** member — the join rights come from OU delegation in
the next step, not from group membership.

### Step B — Delegate Create/Delete Computer on the target OU (GUI, recommended)

Do this on the OU where provisioned computers should land (e.g.
`OU=Workstations,DC=gipartners,DC=com`):

1. **Active Directory Users and Computers** → **View → Advanced Features** (on).
2. Right-click the target OU → **Delegate Control…** → **Next**.
3. **Add** `hybridjoin` → **Next**.
4. **Create a custom task to delegate** → **Next**.
5. **Only the following objects in the folder** → check **Computer objects**, then
   check both **Create selected objects in this folder** and **Delete selected
   objects in this folder** → **Next**.
6. Permissions — check **General** and **Property-specific**, then grant:
   - Read All Properties / Write All Properties
   - Read Permissions
   - Reset Password / Change Password
   - Read and write Account Restrictions
   - Validated write to DNS host name
   - Validated write to service principal name
7. **Next → Finish**.

### Step B (alt) — dsacls (scripted)

The create/delete grant is simple and robust to script; use the GUI wizard above for
the property-level bits:

```cmd
dsacls "OU=Workstations,DC=gipartners,DC=com" /I:T /G "GIPARTNERS\hybridjoin:CCDC;computer"
```

`CCDC;computer` = create + delete **computer** child objects; `/I:T` applies down the
subtree (drop to default/`/I:S` if you don't want sub-OUs included).

### Step C — Point the package at the same OU

In WCD, the domain-join **AccountOU must match the delegated OU**. In the *Provision
desktop devices* wizard's **Join Active Directory** step, or in *Advanced
provisioning* under **Runtime settings → Accounts → ComputerAccount**, set:

- **Account / UserName**: `GIPARTNERS\hybridjoin` (or `hybridjoin@gipartners.com`)
- **Password**: the account password
- **AccountOU**: `OU=Workstations,DC=gipartners,DC=com`
- **DomainName**: `gipartners.com`

If you don't set AccountOU, computers land in the default **Computers** container —
where `hybridjoin` has no delegated rights — and the join fails. Either set AccountOU
or redirect the default container with `redircmp`.

### Step D — Harden the account

- **No admin groups** — verify it's only in Domain Users.
- **Deny interactive logon**: via GPO grant it *Deny log on locally*, *Deny log on
  through Remote Desktop Services*, *Deny log on as a batch job/service*. Domain join
  is a network operation and does not need any of these.
- **MachineAccountQuota**: OU delegation works regardless of the domain
  `ms-DS-MachineAccountQuota`. If that quota is still the default 10 for all
  authenticated users, consider setting it to 0 domain-wide so only delegated
  accounts can join — tightens the whole domain, not just this account.
- **Rotate**: change the password after each provisioning campaign (or on any
  suspected package exposure) and rebuild the package. Always password-protect the
  .ppkg (Step Finish → *Protect your package*).

## Prerequisites

- **Windows Configuration Designer (WCD)** — install from the Microsoft Store
  ("Windows Configuration Designer") or as part of the Windows ADK. Run it on a
  Windows admin workstation.
- **Directory permissions**:
  - *Entra join path*: an account allowed to join devices to Entra and create a
    **bulk enrollment token** (typically Intune Administrator / Cloud Device
    Administrator). That account should be **excluded from Conditional Access / MFA
    prompts that would block the unattended join**, and must not exceed the Entra
    "maximum devices per user" quota.
  - *AD domain join path*: domain credentials permitted to join computers to the
    target OU.
- **Intune auto-enrollment**: Entra → Mobility (MDM) → Microsoft Intune → MDM user
  scope covers the enrolling users, and devices are licensed for Intune.
- **USB drive** formatted **FAT32** (OOBE reads FAT32 reliably; NTFS can work but
  FAT32 is safest). 8 GB is plenty for a .ppkg.
- **BitLocker Intune policy already configured** (see [README](README.md#intune-policy-prevention-going-forward)) —
  otherwise the provisioned device joins but never encrypts/escrows.

## Step 1 — Start the package in WCD

1. Launch **Windows Configuration Designer**.
2. Choose **Provision desktop devices** (the guided wizard — sufficient for
   join + enroll; use *Advanced provisioning* only if you need custom settings).
3. **Name** the project, e.g. `GI-OOBE-Provisioning`, and note the export folder.

## Step 2 — Fill in the wizard

**① Set up device**
- **Device name**: use a template so each device is unique, e.g. `GI-%SERIAL%`
  (or `GI-%RAND:5%`). Keep within the 15-char NetBIOS limit if domain-joining.
- **Enter product key** / edition upgrade: only if you need to change the Windows
  edition.
- **Configure devices for shared use**: **Off** (leave default) unless building
  kiosks.

**② Set up network** (so OOBE can reach the directory)
- Add the corporate **Wi-Fi SSID + credentials** if devices provision over wireless.
  Skip if they are always on wired Ethernet at provisioning time.

**③ Account management** — this is the join/enroll step
- **Entra join path**: select **Enroll in Azure AD** → **Get Bulk Token** → sign in
  with the permitted account → complete auth. A bulk token (valid up to **180 days**)
  is embedded in the package. **Set the shortest practical expiry.**
- **AD domain join path** (chosen): select **Join Active Directory**, enter
  `gipartners.com`, the `hybridjoin` credentials, and the **AccountOU** that matches
  the delegation — see [Join account — hybridjoin@gipartners.com](#join-account--hybridjoingipartnerscom).

**④ Add applications / ⑤ Add certificates**
- Usually skip for a join-only package. Add root/enterprise certs here only if OOBE
  needs them to reach your network.

**⑥ Finish**
- **Protect your package**: set a **password** (and sign it if you have a code-signing
  cert). Treat this as mandatory — see Security below.

## Step 3 — Build and stage to USB

1. In the wizard's finish screen (or **Export → Provisioning package**), **Build**.
   WCD produces `<ProjectName>.ppkg` plus supporting files in the export folder.
2. Copy the **`.ppkg`** file to the **root of the USB drive**. (Copying the whole
   export folder is fine too; the device only needs the `.ppkg`.)
3. Label the USB and record which package/version and token-expiry date it carries.

## Step 4 — Apply during OOBE

1. Boot the target device to the **first OOBE screen** (region/"Is this the right
   country or region?").
2. **Insert the USB.** Provisioning usually auto-launches. If it doesn't, **press the
   Windows key five times** on that screen to open the "Provision this device" flow.
3. Select the **`.ppkg`**, enter the package password if set, and confirm. The device
   applies the package, joins the directory, and begins Intune enrollment.

> **Post-OOBE alternative** (device already past OOBE): Settings → Accounts →
> **Access work or school** → **Add or remove a provisioning package** → **Add a
> package** → select the `.ppkg`.

## Step 5 — Verify the device

- **Join state** (run as SYSTEM/admin on the device): `dsregcmd /status`
  - Entra join path → `AzureAdJoined : YES`
  - Hybrid path → `DomainJoined : YES` and, after sync, `AzureAdJoined : YES`
- **Directory/Intune**: the device appears in Entra → Devices and Intune → Devices.
- **BitLocker**: after the Intune disk-encryption policy applies, the OS drive
  encrypts and the key escrows. Confirm on-device with
  `manage-bde -protectors -get C:` (a `RecoveryPassword` protector present).
- **Authoritative check**: re-run `Get-DevicesMissingBitLockerKeys.ps1`; the new
  device should show as compliant (key present in Entra).

## Security — handle the USB as sensitive

The bulk enrollment token embedded in the `.ppkg` can join devices to your tenant as
the token account for the life of the token (up to 180 days). Therefore:

- **Password-protect** (and sign, if possible) the package in WCD.
- **Store the USB securely**; don't leave it in provisioning areas unattended.
- **Use the shortest token expiry** you can operate with, and **rebuild before it
  expires**. Expiry does not affect already-joined devices.
- If a USB is lost, **rotate**: disable/rotate the token account's credentials and
  rebuild the package. Review recently joined devices in Entra.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Provisioning doesn't launch at OOBE | USB not FAT32, or `.ppkg` not at root | Reformat FAT32, put `.ppkg` at root; press Windows key 5× on the region screen |
| Package applies but Entra join fails | Bulk token expired, join account lacks rights, or CA/MFA blocked the unattended join | Rebuild with a fresh token; grant device-join rights; exclude the account from blocking CA/MFA |
| "Maximum number of devices reached" | Entra per-user device quota hit by the token account | Raise the quota (Entra → Devices → Device settings) or use a dedicated join account |
| Joins Entra but never enrolls in Intune | MDM user scope not set, or unlicensed | Entra → Mobility (MDM) → Intune: set MDM user scope; assign Intune licenses |
| No network during OOBE | No Wi-Fi profile in package, no wired link | Add a Wi-Fi profile in the wizard, or provision on wired Ethernet |
| Joined but BitLocker key not in Entra | Disk-encryption policy missing/not applied, or domain-join path escrowing to AD DS only | Confirm the Intune BitLocker policy saves recovery info to Entra; let the Remediations pair backfill |

## Notes & limitations

- A join-only `.ppkg` **enrolls**; it does not encrypt or escrow. The Intune
  BitLocker policy does that — keep both in place.
- Bulk tokens expire (≤180 days); this package **must be rebuilt periodically**.
- FAT32 USB and package-at-root are the reliability keys for OOBE pickup.
- For fully cloud, zero-touch provisioning at scale, Windows Autopilot is the
  alternative to a USB package — out of scope for this runbook, which covers the
  USB `.ppkg` path.
