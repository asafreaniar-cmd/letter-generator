param(
    [string]$BaseUrl = "http://127.0.0.1:8080/"
)

$ErrorActionPreference = "Stop"

$AppRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$PythonExe = Join-Path $AppRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $PythonExe)) {
    throw "Python virtual environment not found at $PythonExe"
}

& $PythonExe "$PSScriptRoot\verify_exact_pipeline.py" --base-url $BaseUrl
exit $LASTEXITCODE
