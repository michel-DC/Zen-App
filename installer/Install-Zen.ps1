param(
    [string] $SourcePath = (Join-Path $PSScriptRoot '..\Zen-Windows11'),
    [switch] $DoNotLaunch
)

$ErrorActionPreference = 'Stop'

$source = [IO.Path]::GetFullPath($SourcePath)
$sourceExecutable = Join-Path $source 'Zen.exe'
if (-not (Test-Path -LiteralPath $sourceExecutable)) {
    throw "Zen.exe est introuvable dans le dossier source : $source"
}

$localPrograms = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs'
$installPath = [IO.Path]::GetFullPath((Join-Path $localPrograms 'Zen'))
$expectedPath = [IO.Path]::GetFullPath((Join-Path $localPrograms 'Zen'))
if ($installPath -ne $expectedPath) { throw "Le dossier d’installation de Zen est invalide." }

$installedExecutable = Join-Path $installPath 'Zen.exe'
Get-Process Zen -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -eq $installedExecutable } |
    Stop-Process -Force

if (Test-Path -LiteralPath $installPath) {
    Remove-Item -LiteralPath $installPath -Recurse -Force
}
New-Item -ItemType Directory -Path $installPath -Force | Out-Null
Copy-Item -Path (Join-Path $source '*') -Destination $installPath -Recurse -Force

$uninstallerSource = Join-Path $PSScriptRoot 'Uninstall-Zen.ps1'
$uninstallerDestination = Join-Path $installPath 'Uninstall-Zen.ps1'
Copy-Item -LiteralPath $uninstallerSource -Destination $uninstallerDestination -Force

$programsFolder = [Environment]::GetFolderPath('Programs')
$shortcutPath = Join-Path $programsFolder 'Zen.lnk'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $installedExecutable
$shortcut.WorkingDirectory = $installPath
$shortcut.IconLocation = "$installedExecutable,0"
$shortcut.Description = 'Zen - Outils locaux pour documents et images'
$shortcut.Save()

$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Zen'
New-Item -Path $uninstallKey -Force | Out-Null
$uninstallCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$uninstallerDestination`""
$estimatedSize = [int] ((Get-ChildItem -LiteralPath $installPath -Recurse -File | Measure-Object Length -Sum).Sum / 1KB)
New-ItemProperty -Path $uninstallKey -Name DisplayName -Value 'Zen' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name DisplayVersion -Value '0.1.0' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name Publisher -Value 'michel-DC' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name InstallLocation -Value $installPath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name DisplayIcon -Value "$installedExecutable,0" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name UninstallString -Value $uninstallCommand -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name NoModify -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name NoRepair -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $uninstallKey -Name EstimatedSize -Value $estimatedSize -PropertyType DWord -Force | Out-Null

if (-not $DoNotLaunch) {
    Start-Process -FilePath $installedExecutable
}

[pscustomobject]@{
    Name = 'Zen'
    Version = '0.1.0'
    InstallPath = $installPath
    Shortcut = $shortcutPath
}
