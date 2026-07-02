param(
    [string]$SetupExe = ".\setup.exe",
    [string]$ConfigXml = ".\office-ltsc-2024-config.xml"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $SetupExe)) {
    throw "Office Deployment Tool setup.exe not found at $SetupExe"
}

if (-not (Test-Path $ConfigXml)) {
    throw "Office configuration XML not found at $ConfigXml"
}

Write-Host "Installing Office LTSC 2024..." -ForegroundColor Cyan
& $SetupExe /configure $ConfigXml

Write-Host ""
Write-Host "Done. Verify that Word opens once under the dedicated server user." -ForegroundColor Green
