<#
.SYNOPSIS
    Intune Remediations DETECTION script for BitLocker Entra ID key escrow.
.DESCRIPTION
    Reports the device NON-COMPLIANT (exit 1 -> triggers the remediation script)
    when a BitLocker-protected volume either:
      - lacks a RecoveryPassword protector (nothing exists to escrow), or
      - has no recorded successful escrow to Entra ID
        (no stamp written by Backup-BitLockerToEntra.ps1).

    Reports COMPLIANT (exit 0) when there are no protected volumes, or when every
    protected volume has a RecoveryPassword protector and a success stamp exists.

    Read-only: never changes BitLocker state. Designed to run as SYSTEM in
    Windows PowerShell 5.1, 64-bit, under Intune -> Devices -> Remediations.
.NOTES
    This is a device-side gate only. It cannot confirm the key actually landed in
    Entra. Authoritative, tenant-side verification is Get-DevicesMissingBitLockerKeys.ps1
    (Microsoft Graph), which remains the source of truth for compliance reporting.

    Run in 64-bit PowerShell so the HKLM:\SOFTWARE\GI stamp is read from the same
    registry view the remediation script writes to.
#>

$ErrorActionPreference = 'Stop'
$stampKey = 'HKLM:\SOFTWARE\GI\BitLockerEscrow'

try {
    $volumes = Get-BitLockerVolume | Where-Object {
        $_.ProtectionStatus -eq 'On' -or $_.VolumeStatus -ne 'FullyDecrypted'
    }

    if (-not $volumes) {
        Write-Output 'Compliant: no BitLocker-protected volumes.'
        exit 0
    }

    $reasons = @()

    foreach ($v in $volumes) {
        $hasRecovery = $v.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
        if (-not $hasRecovery) {
            $reasons += "$($v.MountPoint) has no RecoveryPassword protector"
        }
    }

    # Success stamp written by Backup-BitLockerToEntra.ps1 after a clean escrow.
    $stamped = $false
    if (Test-Path $stampKey) {
        $last = (Get-ItemProperty -Path $stampKey -Name 'LastBackupUtc' -ErrorAction SilentlyContinue).LastBackupUtc
        if ($last) { $stamped = $true }
    }
    if (-not $stamped) {
        $reasons += 'no successful Entra escrow recorded'
    }

    if ($reasons.Count -gt 0) {
        Write-Output ('Non-compliant: ' + ($reasons -join '; '))
        exit 1
    }

    Write-Output 'Compliant: all protected volumes have a RecoveryPassword protector and a recorded Entra escrow.'
    exit 0
}
catch {
    # Fail toward remediation so a detection error does not silently skip escrow.
    Write-Output "Detection error, flagging for remediation: $($_.Exception.Message)"
    exit 1
}
