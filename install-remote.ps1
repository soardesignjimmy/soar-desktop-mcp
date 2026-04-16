# SOAR Desktop MCP - One-Line Remote Installer
# Usage: irm https://raw.githubusercontent.com/soardesignjimmy/soar-desktop-mcp/main/install-remote.ps1 | iex

$ErrorActionPreference = "Stop"
$installDir = "$env:USERPROFILE\soar-desktop-mcp"
$zipUrl = "https://github.com/soardesignjimmy/soar-desktop-mcp/archive/refs/heads/main.zip"
$zipFile = "$env:TEMP\soar-desktop-mcp.zip"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SOAR Desktop MCP - One-Click Installer" -ForegroundColor Cyan
Write-Host "  Control any Windows app via Accessibility Tree" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Download
Write-Host "[1/4] Downloading from GitHub ..." -ForegroundColor Yellow
if (Test-Path $zipFile) { Remove-Item $zipFile -Force }
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing
Write-Host "  Downloaded OK" -ForegroundColor Green

# 2. Extract
Write-Host "[2/4] Extracting ..." -ForegroundColor Yellow
if (Test-Path $installDir) {
    Write-Host "  Existing install found - backing up venv ..."
    if (Test-Path "$installDir\venv") {
        Rename-Item "$installDir\venv" "$env:TEMP\soar-desktop-venv-backup" -Force -ErrorAction SilentlyContinue
    }
    Remove-Item $installDir -Recurse -Force
}
Expand-Archive -Path $zipFile -DestinationPath $env:TEMP -Force
Move-Item "$env:TEMP\soar-desktop-mcp-main" $installDir -Force
# Restore venv if backed up
if (Test-Path "$env:TEMP\soar-desktop-venv-backup") {
    Move-Item "$env:TEMP\soar-desktop-venv-backup" "$installDir\venv" -Force
}
Remove-Item $zipFile -Force
Write-Host "  Extracted to: $installDir" -ForegroundColor Green

# 3. Run install.bat
Write-Host "[3/4] Running install.bat ..." -ForegroundColor Yellow
Push-Location $installDir
cmd /c "install.bat"
Pop-Location

# 4. Show config
$runBat = "$installDir\run_mcp.bat"
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Installed to: $installDir" -ForegroundColor White
Write-Host ""
Write-Host "  Add to Claude Cowork / Claude Code config:" -ForegroundColor White
Write-Host ""
Write-Host "  {" -ForegroundColor Gray
Write-Host "    `"mcpServers`": {" -ForegroundColor Gray
Write-Host "      `"soar-desktop`": {" -ForegroundColor Gray
Write-Host "        `"command`": `"$($runBat -replace '\\','\\')`"" -ForegroundColor White
Write-Host "      }" -ForegroundColor Gray
Write-Host "    }" -ForegroundColor Gray
Write-Host "  }" -ForegroundColor Gray
Write-Host ""
