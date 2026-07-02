param(
    [string]$AppRoot = "C:\letter-generator",
    [string]$Host = "0.0.0.0",
    [string]$Port = "8080",
    [string]$StorageRoot = "C:\letter-generator\data",
    [string]$Domain = "",
    [string]$TaskName = "LetterGeneratorApp",
    [string]$PythonLauncher = "py"
)

$ErrorActionPreference = "Stop"

$InstallScript = Join-Path $PSScriptRoot "install.ps1"
$CaddyScript = Join-Path $PSScriptRoot "install_caddy.ps1"

& $InstallScript -AppRoot $AppRoot -Host $Host -Port $Port -StorageRoot $StorageRoot -TaskName $TaskName -PythonLauncher $PythonLauncher

if ($Domain) {
    & $CaddyScript -Domain $Domain -AppPort $Port
} else {
    Write-Host ""
    Write-Host "No domain supplied. HTTPS was not configured." -ForegroundColor Yellow
}
