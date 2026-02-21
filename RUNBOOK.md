# BondMaster Excel Add-in Operational Runbook

## Overview

BondMaster Excel is an xlOil-based Excel add-in that provides real-time bond data via Python UDFs. This runbook covers installation, troubleshooting, and maintenance.

---

## Quick Reference

| Item | Value |
|------|-------|
| Python Version | 3.11+ |
| Excel Version | 2016+ (64-bit recommended) |
| Dependencies | xlOil, httpx, pydantic |
| API Endpoint | Configurable in `xlOil.ini` |

---

## Installation

### Prerequisites

1. **Python 3.11+** installed and in PATH
2. **Excel 2016+** (64-bit recommended for large datasets)
3. **BondMaster API** running (local or remote)

### Install Steps

```powershell
# 1. Clone or download release
git clone https://github.com/JeffExec/bondmaster-excel.git
cd bondmaster-excel

# 2. Install dependencies
pip install uv
uv sync

# 3. Configure API endpoint (edit xlOil.ini)
notepad xlOil.ini
# Set BONDMASTER_URL=http://localhost:8000

# 4. Register add-in in Excel
# Open Excel > File > Options > Add-ins
# Manage: COM Add-ins > Go
# Browse to: <install-path>\xlOil.xll
# Click OK
```

### Verify Installation

In Excel, enter in any cell:
```
=BONDAPI_STATUS()
```

Should return: `Connected to http://localhost:8000`

---

## Configuration

### xlOil.ini

```ini
[xlOil]
Plugins=xlOil_Python

[xlOil_Python]
# Path to Python interpreter (leave empty to use PATH)
PYTHONEXECUTABLE=

# Module to load
LoadModules=bondmaster_excel.functions

[Environment]
# BondMaster API URL
BONDMASTER_URL=http://localhost:8000
```

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `BONDMASTER_URL` | API server URL | `http://localhost:8000` |
| `BONDMASTER_TIMEOUT` | Request timeout (seconds) | `30` |

---

## Available Functions

| Function | Description | Example |
|----------|-------------|---------|
| `BONDAPI_STATUS()` | Check API connectivity | `=BONDAPI_STATUS()` |
| `BONDSTATIC(isin, field)` | Get static field | `=BONDSTATIC("US912810TM60","coupon_rate")` |
| `BONDINFO(isin)` | Get all bond info | `=BONDINFO("US912810TM60")` |
| `BONDSEARCH(query)` | Search bonds | `=BONDSEARCH("OATEI 2030")` |
| `BONDSBYTENOR(country, tenor)` | List by tenor | `=BONDSBYTENOR("US","10Y")` |
| `BONDSBYCOUNTRY(country)` | List by country | `=BONDSBYCOUNTRY("GB")` |

---

## Troubleshooting

### "xlOil Py" Tab Not Visible

1. **Check COM Add-ins**
   - Excel > File > Options > Add-ins
   - Manage: COM Add-ins > Go
   - Verify "xlOil Core" is checked

2. **Check xlOil logs**
   ```
   %APPDATA%\xlOil\xlOil.log
   ```

3. **Reinstall xlOil**
   ```powershell
   pip uninstall xloil xloil-core
   pip install xloil
   ```

### #VALUE! Error in Formulas

1. **Check API server is running**
   ```powershell
   curl http://localhost:8000/health
   ```

2. **Check network connectivity**
   - Ensure firewall allows port 8000
   - Try `=BONDAPI_STATUS()` to diagnose

3. **Check xlOil Python console**
   - xlOil Py tab > Log Window
   - Look for Python exceptions

### Slow Formula Calculation

1. **Enable caching**
   - BondMaster Excel caches responses for 5 minutes
   - First call may be slow; subsequent calls are cached

2. **Reduce API calls**
   - Use `BONDINFO()` once vs. multiple `BONDSTATIC()` calls
   - Consider using array formulas

3. **Check API server performance**
   ```powershell
   curl -w "Time: %{time_total}s\n" http://localhost:8000/health
   ```

### Excel Crashes on Load

1. **Check Python version**
   - xlOil requires Python 3.9-3.12
   - Ensure 64-bit Python for 64-bit Excel

2. **Check for conflicting add-ins**
   - Disable other Python add-ins (PyXLL, etc.)
   - Try loading in safe mode: `excel.exe /safe`

3. **Check Windows Event Viewer**
   - Application logs for Excel crashes
   - Look for faulting module (xlOil.xll, Python312.dll, etc.)

---

## Maintenance

### Update Add-in

```powershell
cd bondmaster-excel
git pull
uv sync

# Restart Excel to load new code
```

### Clear Cache

```powershell
# Delete xlOil cache
Remove-Item -Recurse -Force "$env:APPDATA\xlOil\cache"

# Or in Python console:
# import bondmaster_excel.functions as f
# f._cache.clear()
```

### Reset xlOil

```powershell
# Unregister add-in
# Excel > File > Options > Add-ins > COM Add-ins > Go
# Uncheck xlOil Core

# Clear config
Remove-Item -Recurse -Force "$env:APPDATA\xlOil"

# Reinstall
pip install --force-reinstall xloil
```

---

## Deployment (Enterprise)

### Network Deployment

1. **Shared network drive**
   ```
   \\server\apps\bondmaster-excel\
   ├── bondmaster_excel\
   ├── xlOil.ini
   ├── xlOil.xll
   └── install.bat
   ```

2. **Group Policy**
   - Deploy add-in registration via GPO
   - Set HKCU\Software\Microsoft\Office\16.0\Excel\Add-in Manager

3. **Centralized API**
   - Set `BONDMASTER_URL` in xlOil.ini to shared API
   - Example: `http://bondmaster.internal:8000`

### Silent Install

```powershell
# install-silent.ps1
param([string]$ApiUrl = "http://localhost:8000")

# Install Python dependencies
pip install uv
uv sync

# Configure API URL
(Get-Content xlOil.ini) -replace 'BONDMASTER_URL=.*', "BONDMASTER_URL=$ApiUrl" | Set-Content xlOil.ini

# Register add-in (requires admin for HKLM)
$addinPath = Join-Path $PWD "xlOil.xll"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$openCount = (Get-ItemProperty $regPath -Name OPEN -ErrorAction SilentlyContinue).OPEN
$nextOpen = if ($openCount) { "OPEN$([int]$openCount + 1)" } else { "OPEN" }
New-ItemProperty -Path $regPath -Name $nextOpen -Value $addinPath -PropertyType String -Force
```

---

## Monitoring

### Health Check Script

```powershell
# health-check.ps1
$apiUrl = "http://localhost:8000"

try {
    $response = Invoke-RestMethod -Uri "$apiUrl/health" -TimeoutSec 5
    if ($response.status -eq "healthy") {
        Write-Host "API: OK" -ForegroundColor Green
        exit 0
    }
} catch {
    Write-Host "API: FAILED - $_" -ForegroundColor Red
    exit 1
}
```

### Log Locations

| Log | Location |
|-----|----------|
| xlOil | `%APPDATA%\xlOil\xlOil.log` |
| Python | Excel > xlOil Py > Log Window |
| API | `journalctl -u bondmaster` (Linux) |

---

## Contacts

| Role | Contact |
|------|---------|
| Primary On-Call | Set up PagerDuty/Opsgenie |
| Excel Support | - |
| API Support | - |

---

## Changelog

- **2026-02-21**: Initial runbook created
