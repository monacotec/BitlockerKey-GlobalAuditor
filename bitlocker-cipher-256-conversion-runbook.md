# Runbook - Convert BitLocker AES-128 to AES-256

Identify devices encrypted at XTS-AES-128 (below the AES-256 policy) and convert them
in place, then re-escrow the new recovery key to Entra. Detection is done tenant-side
with Intune; the conversion itself is delivered by PDQ Connect.

## Read this first - the decrypt window

BitLocker cannot change cipher strength in place. Converting 128 -> 256 requires a
FULL DECRYPT followed by a RE-ENCRYPT. During the decrypt phase the drive is NOT
protected by BitLocker. That window is unavoidable, so:

- Convert only physically secure, on-network devices, ideally on AC power.
- Throttle to small batches; do not blanket the whole fleet at once.
- The converter refuses to start if the TPM is not present/ready (so a device can
  always be re-protected unattended after decrypt).
- TPM+PIN devices are out of scope - this re-encrypts unattended with a TPM-only plus
  RecoveryPassword protector.

Existing AES-128 is still strong. If the decrypt window is unacceptable for a given
device, the alternative is to reimage/reprovision it (the OOBE provisioning path
encrypts at 256 from the start) - see oobe-provisioning-usb-runbook.md.

## Pieces

| File | Role | Delivery |
|---|---|---|
| `Detect-BitLockerCipherStrength.ps1` | Detection: flags any volume not XtsAes256 | Intune Remediation (detection only) |
| `Install-BitLocker256ConversionTask.ps1` | Stages the converter and registers a SYSTEM scheduled task | PDQ Connect (as SYSTEM) |
| `Convert-BitLockerTo256.ps1` | Engine: decrypt -> re-encrypt at 256 -> re-add protectors -> escrow | Run by the scheduled task |

Why a scheduled task instead of running the converter directly from PDQ: the
decrypt+re-encrypt takes hours and spans reboots. The task runs as SYSTEM shortly
after install and again at every startup, resuming from the live on-disk state each
time, until the drive reaches XTS-AES-256 - then it removes itself. PDQ just kicks it
off and returns immediately.

## Prerequisites

- Devices have a present and ready TPM (the converter enforces this).
- The Intune BitLocker policy is set to XTS-AES-256 for the OS/fixed drives, so any
  background re-encryption stays at 256 and new devices never regress.
- PDQ Connect available and able to run as SYSTEM (the default).
- BitLocker recovery keys for the current (128) state are already escrowed to Entra
  before you start - so a device is recoverable if it reboots mid-convert. Confirm
  with Get-DevicesMissingBitLockerKeys.ps1 first.

## Step 1 - Identify the AES-128 population (Intune)

Deploy `Detect-BitLockerCipherStrength.ps1` as a detection-only remediation:

1. Intune admin center -> Devices -> Remediations -> Create script package.
2. Name: `BitLocker - Report AES-128 volumes`.
3. Detection script file: `Detect-BitLockerCipherStrength.ps1`. Leave the remediation
   script empty (report only).
4. Run options: Run using logged-on credentials = No; Enforce signature check = No;
   Run in 64-bit PowerShell = Yes.
5. Assign broadly, Daily. Devices at 128 come back "With issues", and the detection
   output names the volume and method (e.g. `C:=XtsAes128`).
6. Export the run states to size the population and track it down over time.

## Step 2 - Enforce 256 going forward (policy)

Confirm the Intune disk-encryption policy sets encryption method to XTS-AES-256 for OS
and fixed drives. This does not touch already-encrypted drives, but guarantees any
re-encryption (including this conversion) and any new/reprovisioned device lands at
256.

## Step 3 - Convert via PDQ Connect (throttled)

Package both scripts together and deploy as SYSTEM to a SMALL target group:

1. In PDQ Connect, create a package that runs `Install-BitLocker256ConversionTask.ps1`
   with `Convert-BitLockerTo256.ps1` included as a package file next to it (the
   installer copies the converter from its own folder to C:\GI).
   - Optional switches: `-UsedSpaceOnly` (faster, less thorough) and `-RequireACPower`
     (abort before decrypt if on battery). Default is full-disk, battery-allowed.
2. Target a pilot ring first (a handful of devices you can watch).
3. Deploy. The installer returns immediately; within ~2 minutes the scheduled task
   starts the converter as SYSTEM.
4. Widen the ring only after the pilot reaches 256 cleanly.

What happens on each device:

- Converter checks the live state. If already 256+fully encrypted, it just escrows and
  removes the task.
- Otherwise: decrypt (waits, logging percent), then Enable-BitLocker at XtsAes256 with
  a fresh TPM protector, add a RecoveryPassword protector, escrow it to Entra, write
  the escrow stamp, and remove the scheduled task. Encryption then finishes in the
  background.
- If the device reboots mid-way, the startup trigger re-runs the converter and it
  resumes from wherever the volume actually is.

## Step 4 - Verify

- On a device: `Get-BitLockerVolume -MountPoint C: | Select-Object MountPoint,
  EncryptionMethod,VolumeStatus,EncryptionPercentage` -> expect `XtsAes256`.
- Progress/troubleshooting log: `C:\GI\BitLockerConvert_<computername>.log`.
- State marker: `HKLM:\SOFTWARE\GI\BitLockerConvert` (Phase = Done when complete).
- Re-run the Step 1 Intune detection - the "With issues" count should drop.
- Re-run `Get-DevicesMissingBitLockerKeys.ps1` to confirm the new recovery key is in
  Entra (the conversion re-escrows because decrypt removes the old protectors).

## Throttling and comms

- Small rings only. Decryption + full re-encryption is disk-intensive for hours and
  users will notice performance impact - schedule accordingly and warn them.
- Prefer AC power (`-RequireACPower`) for laptops.
- Watch the pilot logs before scaling. Track the population down with the Step 1
  detection export.

## Risks and rollback

- The decrypt window is the main risk - see the top of this runbook.
- There is no clean "undo": once decrypt starts, the safe direction is forward to 256.
  If Enable-BitLocker fails, the task retries on next boot; the drive stays decrypted
  until it succeeds, so investigate failing pilots before widening.
- Recoverability: because current 128 keys are escrowed before you start (prereq) and
  the converter re-escrows the new key immediately after re-enabling, a device is
  recoverable at both ends. The only gap is the decrypt window, when there is no
  BitLocker protector at all by design.

## Notes and limitations

- Runs in Windows PowerShell 5.1 as SYSTEM. Scripts are ASCII-only.
- TPM+PIN, network-unlock, and non-TPM configurations are out of scope.
- Full-disk encryption is the default on re-encrypt (safer than used-space-only for a
  drive that previously held recoverable cleartext).
- The converter does not wait for background encryption to finish before exiting; the
  cipher method is already 256 at that point, which is what detection and policy check.
