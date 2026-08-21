<#
.SYNOPSIS
    Intune Remediations DETECTION script: flag volumes NOT encrypted with XTS-AES-256.
.DESCRIPTION
    Reports the device NON-COMPLIANT (exit 1 -> "With issues") when any BitLocker-
    protected volume uses an encryption method other than XtsAes256 (typically
    XtsAes128 from Windows silent/default encryption applied before the AES-256
    policy landed). Reports COMPLIANT (exit 0) when all protected volumes are
    XtsAes256, or nothing is encrypted.

    Detection only - it never changes encryption. Use it to size and track the
    AES-128 population; Convert-BitLockerTo256.ps1 performs the actual conversion.

    Runs as SYSTEM in Windows PowerShell 5.1, 64-bit, under Intune. ASCII-only
    (non-ASCII characters break Intune remediation scripts on every device).
.NOTES
    Intune's built-in Encryption report shows status, not cipher strength - this
    script is how you surface 128 vs 256 across the fleet.
#>

$ErrorActionPreference = 'Stop'

try {
    $vols = Get-BitLockerVolume | Where-Object {
        $_.ProtectionStatus -eq 'On' -or $_.VolumeStatus -ne 'FullyDecrypted'
    }

    if (-not $vols) {
        Write-Output 'Compliant: no encrypted volumes.'
        exit 0
    }

    $bad = @()
    foreach ($v in $vols) {
        if ("$($v.EncryptionMethod)" -ne 'XtsAes256') {
            $bad += "$($v.MountPoint)=$($v.EncryptionMethod)"
        }
    }

    if ($bad.Count -gt 0) {
        Write-Output ('Non-compliant: not XTS-AES-256: ' + ($bad -join ', '))
        exit 1
    }

    Write-Output 'Compliant: all protected volumes are XtsAes256.'
    exit 0
}
catch {
    # Fail visible so a detection error is not read as compliant.
    Write-Output "Detection error: $($_.Exception.Message)"
    exit 1
}
