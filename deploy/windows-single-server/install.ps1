param(
    [string]$AppRoot = "C:\letter-generator",
    [string]$Host = "0.0.0.0",
    [string]$Port = "8080",
    [string]$StorageRoot = "C:\letter-generator\data",
    [string]$TaskName = "LetterGeneratorApp",
    [string]$PythonLauncher = "py"
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

if (-not (Test-Path $AppRoot)) {
    throw "AppRoot '$AppRoot' does not exist. Copy the project to the Windows server first."
}

Set-Location $AppRoot

Write-Step "Creating virtual environment"
if (-not (Test-Path ".venv")) {
    & $PythonLauncher -m venv .venv
}

$PythonExe = Join-Path $AppRoot ".venv\Scripts\python.exe"
$PipExe = Join-Path $AppRoot ".venv\Scripts\pip.exe"

Write-Step "Installing Python dependencies"
& $PythonExe -m pip install --upgrade pip
& $PipExe install -r windows_requirements.txt

Write-Step "Preparing storage"
New-Item -ItemType Directory -Force -Path $StorageRoot | Out-Null

$RunScript = Join-Path $AppRoot "run_app_production.ps1"
@"
`$env:HOST = "$Host"
`$env:PORT = "$Port"
`$env:STORAGE_ROOT = "$StorageRoot"
`$env:PDF_ENGINE = "local_word"
`$env:PDF_PROFILE_DEFAULT = "exact"
Set-Location "$AppRoot"
& "$PythonExe" -m waitress --listen=$Host`:$Port app:app
"@ | Set-Content -Path $RunScript -Encoding UTF8

Write-Step "Registering startup task"
$PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$TaskCommand = "`"$PowerShellExe`" -NoProfile -ExecutionPolicy Bypass -File `"$RunScript`""
schtasks /Delete /TN $TaskName /F 2>$null | Out-Null
schtasks /Create /TN $TaskName /SC ONLOGON /RL HIGHEST /TR $TaskCommand /F | Out-Null

Write-Step "Starting application now"
Start-Process -FilePath $PowerShellExe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$RunScript`"" -WindowStyle Hidden

Write-Step "Done"
Write-Host "Application root: $AppRoot"
Write-Host "Storage root: $StorageRoot"
Write-Host "Port: $Port"
Write-Host "Scheduled task: $TaskName"
Write-Host ""
Write-Host "Next:"
Write-Host "1. Install Microsoft Word on this Windows server."
Write-Host "2. Install the David font if it is not already present."
Write-Host "3. Configure auto-login for the dedicated server user if you want Word automation to survive reboot without manual login."
Write-Host "4. Put Caddy or IIS in front of port $Port for HTTPS and a fixed domain."
