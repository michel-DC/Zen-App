param(
    [switch] $Cleanup,
    [string] $InstallPath
)

$ErrorActionPreference = 'Stop'
$expectedPath = [IO.Path]::GetFullPath(
    (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\Zen'))

if (-not $Cleanup) {
    $temporaryScript = Join-Path ([IO.Path]::GetTempPath()) "Zen-Uninstall-$PID.ps1"
    Copy-Item -LiteralPath $PSCommandPath -Destination $temporaryScript -Force
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $temporaryScript,
        '-Cleanup',
        '-InstallPath', $expectedPath)
    return
}

$resolvedInstallPath = [IO.Path]::GetFullPath($InstallPath)
if ($resolvedInstallPath -ne $expectedPath) {
    throw 'Le dossier de désinstallation de Zen est invalide.'
}

$installedExecutable = Join-Path $resolvedInstallPath 'Zen.exe'
Get-Process Zen -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -eq $installedExecutable } |
    Stop-Process -Force

$shortcutPath = Join-Path ([Environment]::GetFolderPath('Programs')) 'Zen.lnk'
if (Test-Path -LiteralPath $shortcutPath) {
    Remove-Item -LiteralPath $shortcutPath -Force
}

$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Zen'
if (Test-Path -LiteralPath $uninstallKey) {
    Remove-Item -LiteralPath $uninstallKey -Recurse -Force
}

if (Test-Path -LiteralPath $resolvedInstallPath) {
    Remove-Item -LiteralPath $resolvedInstallPath -Recurse -Force
}

Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
