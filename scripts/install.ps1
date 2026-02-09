# BondMaster Excel Add-in Installation Script (PowerShell)
# Run: powershell -ExecutionPolicy Bypass -File install.ps1

$ErrorActionPreference = "Stop"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "BondMaster Excel Add-in Installer" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python not found. Please install Python 3.11+ first." -ForegroundColor Red
    exit 1
}

# Install packages
Write-Host ""
Write-Host "[1/5] Installing Python packages..." -ForegroundColor Yellow
pip install bondmaster xloil httpx --upgrade
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install packages." -ForegroundColor Red
    exit 1
}

# Install bondmaster-excel (from local if available, or skip)
Write-Host ""
Write-Host "[2/5] Installing bondmaster-excel..." -ForegroundColor Yellow
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
if (Test-Path "$projectDir\pyproject.toml") {
    pip install -e $projectDir
} else {
    Write-Host "Note: Installing from local source not available. Using pip package." -ForegroundColor Gray
    pip install bondmaster-excel --upgrade 2>$null
}

# Install xlOil add-in
Write-Host ""
Write-Host "[3/5] Installing xlOil Excel Add-in..." -ForegroundColor Yellow
python -m xloil install
if ($LASTEXITCODE -ne 0) {
    Write-Host "WARNING: xlOil install returned non-zero. May still work." -ForegroundColor Yellow
}

# Load seed data
Write-Host ""
Write-Host "[4/5] Loading bond seed data..." -ForegroundColor Yellow
python -m bondmaster.cli fetch --seed-only
if ($LASTEXITCODE -ne 0) {
    Write-Host "WARNING: Failed to load seed data. You can do this later with: bondmaster fetch --seed-only" -ForegroundColor Yellow
}

# Create desktop shortcuts
Write-Host ""
Write-Host "[5/5] Creating shortcuts..." -ForegroundColor Yellow

$desktopPath = [Environment]::GetFolderPath("Desktop")

# API Server shortcut
$apiScript = @"
@echo off
title BondMaster API Server
echo Starting BondMaster API on http://127.0.0.1:8000
echo Press Ctrl+C to stop
python -m bondmaster.cli serve
pause
"@
$apiScript | Out-File -FilePath "$desktopPath\Start BondMaster API.bat" -Encoding ASCII

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Double-click 'Start BondMaster API.bat' on your Desktop"
Write-Host "2. Open Excel"
Write-Host "3. Try: =BONDSTATIC(`"GB00BYZW3G56`", `"coupon_rate`")"
Write-Host ""
Write-Host "Available functions:" -ForegroundColor Cyan
Write-Host "  =BONDSTATIC(isin, field)    - Get single field"
Write-Host "  =BONDINFO(isin)             - Get all fields (row)"
Write-Host "  =BONDLIST(country)          - List all ISINs"
Write-Host "  =BONDSEARCH(field, value)   - Search/filter"
Write-Host "  =BONDCOUNT(country)         - Count bonds"
Write-Host "  =BONDAPI_STATUS()           - Check connection"
Write-Host ""

Read-Host "Press Enter to exit"
