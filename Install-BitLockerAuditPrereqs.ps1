#Requires -Version 7.0

<#
.SYNOPSIS
    Installs and verifies prerequisites for Get-DevicesMissingBitLockerKeys.ps1.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

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

        $online = Find-Module -Name $mod -ErrorAction SilentlyContinue
        if ($online -and $online.Version -gt $installed.Version) {
            Write-Host "          Updating $mod $($installed.Version) -> $($online.Version)..." -ForegroundColor Yellow
            Update-Module -Name $mod -Scope CurrentUser -Force
            Write-Host "          Updated." -ForegroundColor Green
        }
    }
    else {
        Write-Host "[MISSING] $mod - installing..." -ForegroundColor Yellow
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
        $ver = (Get-Module -ListAvailable -Name $mod | Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-Host "          Installed $mod v$ver" -ForegroundColor Green
    }
}

# Verify all modules can be imported
Write-Host "`nVerifying module imports..." -ForegroundColor Cyan
$allGood = $true
foreach ($mod in $requiredModules) {
    try {
        Import-Module $mod -ErrorAction Stop
        Write-Host "[OK]      $mod imported successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "[FAIL]    $mod failed to import: $_" -ForegroundColor Red
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
