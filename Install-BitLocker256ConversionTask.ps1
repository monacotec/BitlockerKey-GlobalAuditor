<#
.SYNOPSIS
    Stages Convert-BitLockerTo256.ps1 and registers a SYSTEM scheduled task to run it.
.DESCRIPTION
    The 128->256 conversion decrypts then re-encrypts, which takes hours and spans
    reboots - too long for an Intune remediation and unsafe to leave half-done. This
    installer copies the converter to C:\GI and registers a scheduled task that runs
    it as SYSTEM shortly after install and again at every startup, until the converter
    reaches XTS-AES-256 and removes the task itself.

    Deploy THIS script (with Convert-BitLockerTo256.ps1 alongside it) via PDQ Connect
    as SYSTEM, throttled to small batches. Runs fast and returns immediately; the
    scheduled task does the long work.

    Runs as SYSTEM in Windows PowerShell 5.1. ASCII-only.
.PARAMETER UsedSpaceOnly
    Pass through to the converter: encrypt used space only (faster, less thorough).
.PARAMETER RequireACPower
    Pass through to the converter: abort before decrypting if on battery.
#>
[CmdletBinding()]
param(
    [switch]$UsedSpaceOnly,
    [switch]$RequireACPower
)

$ErrorActionPreference = 'Stop'
$taskName = 'GI-BitLocker256Conversion'
$dest     = 'C:\GI\Convert-BitLockerTo256.ps1'
$source   = Join-Path $PSScriptRoot 'Convert-BitLockerTo256.ps1'

if (-not (Test-Path 'C:\GI')) { New-Item -Path 'C:\GI' -ItemType Directory -Force | Out-Null }

if (-not (Test-Path $source)) {
    Write-Error "Convert-BitLockerTo256.ps1 not found next to this installer ($source)."
    exit 1
}
Copy-Item -Path $source -Destination $dest -Force
Write-Output "Staged converter to $dest"

# Build the argument string for the converter from the pass-through switches.
$argLine = "-NoProfile -ExecutionPolicy Bypass -File `"$dest`""
if ($UsedSpaceOnly)  { $argLine += ' -UsedSpaceOnly' }
if ($RequireACPower) { $argLine += ' -RequireACPower' }

$action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argLine
$trigStart = New-ScheduledTaskTrigger -AtStartup
$trigNow   = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(2))
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
                -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 12)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigStart, $trigNow `
    -Principal $principal -Settings $settings -Force | Out-Null

Write-Output "Registered scheduled task '$taskName'. Conversion begins within ~2 minutes and"
Write-Output "resumes on each boot until the drive is XTS-AES-256, then the task self-removes."
Write-Output "Progress log: C:\GI\BitLockerConvert_$env:COMPUTERNAME.log"
exit 0
