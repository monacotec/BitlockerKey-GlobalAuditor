#Requires -Version 7.0

<#
.SYNOPSIS
    Installs and verifies prerequisites for Get-DevicesMissingBitLockerKeys.ps1.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Ensure Install-PSResource is available. PSResourceGet itself has no sub-dependencies
# so Install-Module won't hit the "Dependencies\Desktop" temp-path access-denied bug.
if (-not (Get-Command Install-PSResource -ErrorAction SilentlyContinue)) {
    Write-Host "Bootstrapping Microsoft.PowerShell.PSResourceGet..." -ForegroundColor Cyan
    Install-Module -Name Microsoft.PowerShell.PSResourceGet -Scope CurrentUser -Force -AllowClobber
    Import-Module Microsoft.PowerShell.PSResourceGet
    Write-Host "PSResourceGet ready." -ForegroundColor Green
}

# Redirect TEMP to a plain directory. The Graph .nupkg contains a Dependencies\Desktop
# subfolder; Windows treats 'Desktop' as a shell namespace under %TEMP%, causing access
# denied during extraction. A non-special path avoids this entirely.
$customTemp = 'C:\GI\Temp'
if (-not (Test-Path $customTemp)) {
    New-Item -Path $customTemp -ItemType Directory -Force | Out-Null
}
$originalTemp = $env:TEMP
$originalTmp  = $env:TMP
$env:TEMP = $customTemp
$env:TMP  = $customTemp

function Install-OrUpdateModule {
    param([string]$Name, [switch]$Reinstall)
    $params = @{ Name = $Name; Scope = 'CurrentUser'; TrustRepository = $true }
    if ($Reinstall) { $params['Reinstall'] = $true }
    Install-PSResource @params
}

$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Identity.SignIns',
    'ImportExcel'
)

$logDir = 'C:\GI'
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    Write-Host "Created log directory: $logDir" -ForegroundColor Green
}
else {
    Write-Host "Log directory exists:  $logDir" -ForegroundColor Green
}

foreach ($mod in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $mod | Sort-Object Version -Descending | Select-Object -First 1

    if ($installed) {
        Write-Host "[OK]      $mod v$($installed.Version)" -ForegroundColor Green

        $online = Find-PSResource -Name $mod -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($online -and $online.Version -gt $installed.Version) {
            Write-Host "          Updating $mod $($installed.Version) -> $($online.Version)..." -ForegroundColor Yellow
            Install-OrUpdateModule -Name $mod -Reinstall
            Write-Host "          Updated." -ForegroundColor Green
        }
    }
    else {
        Write-Host "[MISSING] $mod - installing..." -ForegroundColor Yellow
        Install-OrUpdateModule -Name $mod
        $ver = (Get-Module -ListAvailable -Name $mod | Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-Host "          Installed $mod v$ver" -ForegroundColor Green
    }
}

$env:TEMP = $originalTemp
$env:TMP  = $originalTmp

# Verify all modules can be imported in a fresh subprocess to avoid
# "assembly already loaded" errors from a prior version in this session.
Write-Host "`nVerifying module imports..." -ForegroundColor Cyan
$allGood = $true
foreach ($mod in $requiredModules) {
    $result = pwsh -NoProfile -NonInteractive -Command "
        try { Import-Module '$mod' -ErrorAction Stop; 'OK' }
        catch { 'FAIL: ' + `$_.Exception.Message }
    "
    if ($result -eq 'OK') {
        Write-Host "[OK]      $mod imported successfully." -ForegroundColor Green
    }
    else {
        Write-Host "[FAIL]    $mod failed to import: $result" -ForegroundColor Red
        $allGood = $false
    }
}

# PowerShell version check
Write-Host "`nPowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "[WARN]    PowerShell 7+ is required. You are running $($PSVersionTable.PSVersion)." -ForegroundColor Red
    $allGood = $false
}
else {
    Write-Host "[OK]      PowerShell 7+ detected." -ForegroundColor Green
}

Write-Host ""
if ($allGood) {
    Write-Host "All prerequisites are met. You can now run Get-DevicesMissingBitLockerKeys.ps1" -ForegroundColor Green
}
else {
    Write-Host "Some checks failed. Review the output above." -ForegroundColor Red
}
