param(
    [Parameter(Mandatory = $true)][string]$Domain,
    [string]$AppPort = "8080",
    [string]$CaddyRoot = "C:\caddy",
    [string]$ServiceName = "LetterGeneratorCaddy"
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

Write-Step "Preparing directories"
New-Item -ItemType Directory -Force -Path $CaddyRoot | Out-Null

$ZipPath = Join-Path $CaddyRoot "caddy.zip"
$ExtractPath = Join-Path $CaddyRoot "extract"
$CaddyExe = Join-Path $CaddyRoot "caddy.exe"
$Caddyfile = Join-Path $CaddyRoot "Caddyfile"

Write-Step "Downloading Caddy"
Invoke-WebRequest -Uri "https://caddyserver.com/api/download?os=windows&arch=amd64" -OutFile $ZipPath

if (Test-Path $ExtractPath) {
    Remove-Item -Recurse -Force $ExtractPath
}
New-Item -ItemType Directory -Force -Path $ExtractPath | Out-Null

Write-Step "Extracting Caddy"
Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
$DownloadedExe = Get-ChildItem -Path $ExtractPath -Filter "caddy.exe" -Recurse | Select-Object -First 1
if (-not $DownloadedExe) {
    throw "caddy.exe was not found after extraction"
}
Copy-Item -Force $DownloadedExe.FullName $CaddyExe

Write-Step "Writing Caddyfile"
@"
$Domain {
    encode gzip
    reverse_proxy 127.0.0.1:$AppPort
}
"@ | Set-Content -Path $Caddyfile -Encoding UTF8

Write-Step "Opening firewall for 80/443"
New-NetFirewallRule -DisplayName "LetterGenerator HTTP 80" -Direction Inbound -Protocol TCP -LocalPort 80 -Action Allow -ErrorAction SilentlyContinue | Out-Null
New-NetFirewallRule -DisplayName "LetterGenerator HTTPS 443" -Direction Inbound -Protocol TCP -LocalPort 443 -Action Allow -ErrorAction SilentlyContinue | Out-Null

Write-Step "Registering Caddy Windows service"
sc.exe stop $ServiceName | Out-Null 2>$null
sc.exe delete $ServiceName | Out-Null 2>$null
$BinPath = "`"$CaddyExe`" run --environ --config `"$Caddyfile`""
sc.exe create $ServiceName binPath= $BinPath start= auto | Out-Null
sc.exe description $ServiceName "Caddy reverse proxy for Letter Generator" | Out-Null
sc.exe start $ServiceName | Out-Null

Write-Step "Done"
Write-Host "Domain: $Domain"
Write-Host "Caddy root: $CaddyRoot"
Write-Host "Reverse proxy target: 127.0.0.1:$AppPort"
Write-Host ""
Write-Host "Important:"
Write-Host "1. Point the DNS A record of $Domain to this server's public IP."
Write-Host "2. Wait for DNS to propagate."
Write-Host "3. Re-run this script if you change the domain."
