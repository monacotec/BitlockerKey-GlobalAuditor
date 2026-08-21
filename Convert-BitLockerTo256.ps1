<#
.SYNOPSIS
    Converts a BitLocker volume from XTS-AES-128 to XTS-AES-256.
.DESCRIPTION
    BitLocker cannot change cipher strength in place, so this DECRYPTS the volume
    fully and then RE-ENCRYPTS it with XTS-AES-256, re-adds TPM + RecoveryPassword
    protectors, and escrows the new recovery key to Entra ID.

    ****  WARNING  ****
    There is an UNAVOIDABLE window, during the decrypt phase, when the drive is NOT
    protected by BitLocker. Run only on physically secure devices, ideally on AC
    power, throttled to small batches. Not intended for TPM+PIN devices - this
    re-encrypts unattended with a TPM-only + RecoveryPassword protector.

    Intended to run as SYSTEM to completion, driven by the scheduled task that
    Install-BitLocker256ConversionTask.ps1 registers (so it survives reboots and
    is not bound by the Intune 60-minute limit). Idempotent: each run reads the
    live on-disk state and resumes. Self-removes its scheduled task when done.
    ASCII-only.
.PARAMETER MountPoint
    Volume to convert. Default: the OS drive.
.PARAMETER UsedSpaceOnly
    Encrypt used space only (faster). Default is FULL encryption, which is safer
    when converting a drive whose freed sectors may hold recoverable cleartext.
.PARAMETER RequireACPower
    Abort BEFORE decrypting if the device is running on battery.
.NOTES
    Requirements: Administrator/SYSTEM, a present+ready TPM, BitLocker cmdlets.
#>
[CmdletBinding()]
param(
    [string]$MountPoint = $env:SystemDrive,
    [switch]$UsedSpaceOnly,
    [switch]$RequireACPower
)

$ErrorActionPreference = 'Stop'
$targetMethod = 'XtsAes256'
$stampKey     = 'HKLM:\SOFTWARE\GI\BitLockerEscrow'
$stateKey     = 'HKLM:\SOFTWARE\GI\BitLockerConvert'
$taskName     = 'GI-BitLocker256Conversion'
$log          = "C:\GI\BitLockerConvert_$env:COMPUTERNAME.log"
$maxWaitMin   = 480   # 8h ceiling per wait phase

if (-not (Test-Path 'C:\GI')) { New-Item -Path 'C:\GI' -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = ('{0} [{1}] {2}' -f (Get-Date).ToString('s'), $Level, $Message)
    Add-Content -Path $log -Value $line
    Write-Output $line
}

function Set-State {
    param([string]$Phase)
    if (-not (Test-Path $stateKey)) { New-Item -Path $stateKey -Force | Out-Null }
    Set-ItemProperty -Path $stateKey -Name 'Phase' -Value $Phase
    Set-ItemProperty -Path $stateKey -Name 'UpdatedUtc' -Value ((Get-Date).ToUniversalTime().ToString('o'))
}

function Wait-ForStatus {
    param([string]$Mount, [string]$DesiredVolumeStatus)
    $deadline = (Get-Date).AddMinutes($maxWaitMin)
    while ($true) {
        $v = Get-BitLockerVolume -MountPoint $Mount
        Write-Log ("  {0}: {1}, {2}%" -f $Mount, $v.VolumeStatus, $v.EncryptionPercentage)
        if ($v.VolumeStatus -eq $DesiredVolumeStatus) { return $v }
        if ((Get-Date) -gt $deadline) { throw "Timed out waiting for $DesiredVolumeStatus on $Mount after $maxWaitMin min." }
        Start-Sleep -Seconds 30
    }
}

# --- Elevation ---
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Must run as Administrator or SYSTEM. Current: $($id.Name)" 'ERROR'
    exit 1
}

try {
    Write-Log "=== BitLocker 128->256 conversion on $MountPoint (as $($id.Name)) ==="

    $vol = Get-BitLockerVolume -MountPoint $MountPoint
    Write-Log ("Current: Method={0}, VolumeStatus={1}, Protection={2}" -f $vol.EncryptionMethod, $vol.VolumeStatus, $vol.ProtectionStatus)

    if ("$($vol.EncryptionMethod)" -eq $targetMethod -and $vol.VolumeStatus -eq 'FullyEncrypted') {
        Write-Log "Already $targetMethod and fully encrypted. Ensuring escrow only."
        Set-State 'Done'
    }
    else {
        # TPM guard - never decrypt if we cannot re-protect unattended.
        $tpm = Get-Tpm
        if (-not $tpm.TpmPresent -or -not $tpm.TpmReady) {
            Write-Log "TPM not present/ready - refusing to decrypt (would strand the drive)." 'ERROR'
            exit 1
        }

        if ($RequireACPower) {
            $bat  = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
            $onAC = (-not $bat) -or ($bat.BatteryStatus -contains 2)
            if (-not $onAC) {
                Write-Log "On battery and -RequireACPower set - aborting before decrypt." 'ERROR'
                exit 1
            }
        }

        # --- Decrypt phase ---
        if ($vol.VolumeStatus -ne 'FullyDecrypted') {
            if ($vol.VolumeStatus -ne 'DecryptionInProgress') {
                Write-Log "Starting decryption (current method $($vol.EncryptionMethod))."
                Set-State 'Decrypting'
                Disable-BitLocker -MountPoint $MountPoint | Out-Null
            }
            else {
                Write-Log "Decryption already in progress - waiting."
            }
            $vol = Wait-ForStatus -Mount $MountPoint -DesiredVolumeStatus 'FullyDecrypted'
            Write-Log "Fully decrypted."
        }

        # --- Encrypt phase at 256 ---
        Write-Log ("Enabling BitLocker with {0} (full disk = {1})." -f $targetMethod, (-not $UsedSpaceOnly))
        Set-State 'Encrypting'
        $enableParams = @{
            MountPoint       = $MountPoint
            EncryptionMethod = $targetMethod
            TpmProtector     = $true
            SkipHardwareTest = $true
        }
        if ($UsedSpaceOnly) { $enableParams['UsedSpaceOnly'] = $true }
        Enable-BitLocker @enableParams | Out-Null

        $vol = Get-BitLockerVolume -MountPoint $MountPoint
        if (-not ($vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })) {
            Add-BitLockerKeyProtector -MountPoint $MountPoint -RecoveryPasswordProtector | Out-Null
        }
        Resume-BitLocker -MountPoint $MountPoint -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Encryption started (continues in the background)."
        Set-State 'Escrowing'
    }

    # --- Escrow the (new) recovery key to Entra ---
    $vol = Get-BitLockerVolume -MountPoint $MountPoint
    $rp  = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
    $escrowOk = $true
    foreach ($p in $rp) {
        try {
            BackupToAAD-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $p.KeyProtectorId | Out-Null
            Write-Log "Escrowed protector $($p.KeyProtectorId) to Entra."
        }
        catch {
            $escrowOk = $false
            Write-Log "Escrow failed for $($p.KeyProtectorId): $_" 'WARN'
        }
    }
    if ($escrowOk -and $rp) {
        if (-not (Test-Path $stampKey)) { New-Item -Path $stampKey -Force | Out-Null }
        Set-ItemProperty -Path $stampKey -Name 'LastBackupUtc' -Value ((Get-Date).ToUniversalTime().ToString('o'))
    }

    $vol = Get-BitLockerVolume -MountPoint $MountPoint
    Write-Log ("Result: Method={0}, VolumeStatus={1}, Protection={2}, {3}%" -f $vol.EncryptionMethod, $vol.VolumeStatus, $vol.ProtectionStatus, $vol.EncryptionPercentage)

    if ("$($vol.EncryptionMethod)" -ne $targetMethod) {
        Write-Log "Encryption method is still not $targetMethod - leaving task in place to retry." 'ERROR'
        exit 1
    }

    Set-State 'Done'
    # Conversion reached 256 (encryption may still be finishing in the background).
    # Remove the resume task so it stops re-running on boot.
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "[OK] Conversion to $targetMethod complete. Scheduled task removed."
    exit 0
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" 'ERROR'
    Write-Log "If the drive is currently decrypted, the resume task will re-run and re-encrypt at $targetMethod on next boot." 'ERROR'
    exit 1
}
