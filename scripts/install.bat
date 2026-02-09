@echo off
REM BondMaster Excel Add-in Installation Script
REM Run this as Administrator if needed for system-wide install

echo ============================================
echo BondMaster Excel Add-in Installer
echo ============================================
echo.

REM Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.11+ first.
    pause
    exit /b 1
)

echo [1/4] Installing Python packages...
pip install bondmaster bondmaster-excel xloil httpx --upgrade
if errorlevel 1 (
    echo ERROR: Failed to install packages.
    pause
    exit /b 1
)

echo.
echo [2/4] Installing xlOil Excel Add-in...
python -m xloil install
if errorlevel 1 (
    echo ERROR: Failed to install xlOil add-in.
    pause
    exit /b 1
)

echo.
echo [3/4] Loading bond seed data...
python -m bondmaster.cli fetch --seed-only
if errorlevel 1 (
    echo WARNING: Failed to load seed data. You can do this later.
)

echo.
echo [4/4] Creating start script...
echo @echo off > "%USERPROFILE%\Desktop\Start BondMaster API.bat"
echo echo Starting BondMaster API server... >> "%USERPROFILE%\Desktop\Start BondMaster API.bat"
echo python -m bondmaster.cli serve >> "%USERPROFILE%\Desktop\Start BondMaster API.bat"

echo.
echo ============================================
echo Installation Complete!
echo ============================================
echo.
echo Next steps:
echo 1. Double-click "Start BondMaster API.bat" on your Desktop
echo 2. Open Excel
echo 3. Use functions like =BONDSTATIC("GB00BYZW3G56", "coupon_rate")
echo.
echo The xlOil add-in should load automatically when Excel starts.
echo If not, go to: File ^> Options ^> Add-ins ^> Manage: COM Add-ins
echo.
pause
