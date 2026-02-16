# BondMaster Excel Add-in

**Native Excel functions for government bond reference data.** No macro warnings. Blazing fast. Works offline.

```excel
=BONDSTATIC("US912810TM58", "coupon_rate")  →  4.625
=BONDYEARSTOMAT("GB00BYZW3G56")             →  5.42
=BONDLIST("DE")                              →  [list of German bond ISINs]
```

## ✨ Features

- **18 Excel functions** covering all bond-master capabilities
- **Native XLL format** — no macro security warnings
- **Sub-millisecond lookups** via intelligent caching
- **Works offline** — local SQLite database
- **User-friendly errors** — clear messages, not just #N/A
- **Built-in help** — `=BONDHELP()` in any cell

## 📦 Supported Markets

| Market | Bonds | Coverage |
|--------|-------|----------|
| 🇺🇸 US Treasury | 400+ | Full |
| 🇬🇧 UK Gilts | 100+ | Full |
| 🇩🇪 Germany | 90+ | Full |
| 🇫🇷 France | 30+ | Full |
| 🇮🇹 Italy | 200+ | Full |
| 🇪🇸 Spain | 20+ | Full |
| 🇯🇵 Japan | 30+ | Full |
| 🇳🇱 Netherlands | 15+ | Full |

---

## 🚀 Installation (Windows)

### Prerequisites
- Windows 10/11
- Microsoft Excel (desktop version, not web)
- Python 3.11 or 3.12
- Git (for cloning repositories)

### Step 1: Create a project folder and virtual environment

Open **PowerShell** and run:

```powershell
cd ~\PythonProjects   # or wherever you keep projects
mkdir bondmaster-excel
cd bondmaster-excel
python -m venv .venv
.venv\Scripts\activate
```

### Step 2: Install packages from GitHub

Both `bondmaster` (the core library) and `bondmaster-excel` (the Excel add-in) are installed from GitHub:

```powershell
pip install git+https://github.com/JeffExec/bond-master.git git+https://github.com/JeffExec/bondmaster-excel.git xlOil httpx
```

### Step 3: Install xlOil into Excel

```powershell
xloil install
```

This registers xlOil with Excel. You should see:
```
Installed C:\Users\<you>\AppData\Roaming\Microsoft\Excel\XLSTART\xlOil.xll
Edited C:\Users\<you>\AppData\Roaming\xlOil\xlOil.ini to point to <your-venv> python distribution.
```

### Step 4: Configure xlOil to load bondmaster-excel

Open the xlOil config file at:
```
%APPDATA%\xlOil\xlOil.ini
```

#### 4a. Add the Python path

Find the `[[xlOil_Python.Environment]]` section with `XLOIL_PYTHON_PATH` and change it to point to your venv's site-packages:

```ini
XLOIL_PYTHON_PATH='''C:\Users\<you>\PythonProjects\bondmaster-excel\.venv\Lib\site-packages'''
```

> **Important:** Use triple single quotes `'''` for paths (TOML literal strings), not double quotes.

#### 4b. Add bondmaster_excel to LoadModules

Find the `[xlOil_Python]` section and update `LoadModules`:

```ini
[xlOil_Python]
LoadModules=["xloil.xloil_ribbon", "bondmaster_excel.udfs"]
```

> **Important:** Use `bondmaster_excel.udfs` (not just `bondmaster_excel`) — the functions are in the `udfs` submodule.

### Step 5: Load bond data

```powershell
bondmaster fetch --seed-only
```

This downloads reference data for all supported markets (~1000 bonds).

### Step 6: Start the API server

```powershell
bondmaster serve
```

**Keep this terminal open** while using Excel. You should see:
```
INFO:     Uvicorn running on http://127.0.0.1:8000
```

### Step 7: Open Excel and test

1. Open Excel
2. Look for the **xlOil Py** tab in the ribbon (confirms xlOil loaded)
3. In any cell, type: `=BONDAPI_STATUS()`
4. Press Enter

If you see **✓ Connected** — you're done! 🎉

---

## 🔧 Troubleshooting Installation

### xlOil ribbon tab doesn't appear

1. Check Excel Add-ins: File → Options → Add-ins → Manage "Excel Add-ins" → Go
2. Is xlOil.xll listed and checked? If unchecked, Excel disabled it after a crash.
3. Re-run `xloil install` and restart Excel.

### "Error parsing settings file" on Excel startup

Your `xlOil.ini` has a syntax error. Common issues:
- **Use triple quotes for paths:** `'''C:\path\to\folder'''` (not `"C:\path..."`)
- **Escape backslashes in double quotes:** `"C:\\path\\to\\folder"` OR use triple quotes

### Functions show #NAME? error

The module failed to load. Click **Open Log** in the xlOil ribbon to see the error.

**Common causes:**
1. **Wrong LoadModules:** Must be `bondmaster_excel.udfs`, not `bondmaster_excel`
2. **Path not set:** `XLOIL_PYTHON_PATH` must point to your venv's `site-packages`
3. **Import error:** Test with `python -c "import bondmaster_excel.udfs"` in your venv

### "Cannot connect" error in cells

The API server isn't running. Open a new terminal, activate your venv, and run:
```powershell
.venv\Scripts\activate
bondmaster serve
```

### xlOil Log shows "TypeError: func() got an unexpected keyword argument 'category'"

Your xlOil version doesn't support the `category` parameter. Update bondmaster-excel:
```powershell
pip install --upgrade git+https://github.com/JeffExec/bondmaster-excel.git
```

---

## 📖 Function Reference

### Core Data Functions

| Function | Description | Example |
|----------|-------------|---------|
| `BONDSTATIC(isin, field)` | Get any field value | `=BONDSTATIC("US912810TM58", "coupon")` |
| `BONDINFO(isin, headers)` | Get all fields as row | `=BONDINFO("GB00BYZW3G56", TRUE)` |
| `BONDLIST(country, type)` | List ISINs by country | `=BONDLIST("DE", "NOMINAL")` |
| `BONDSEARCH(f1, v1, ...)` | Search with filters | `=BONDSEARCH("country", "US", "security_type", "INDEX_LINKED")` |
| `BONDCOUNT(country)` | Count bonds | `=BONDCOUNT("GB")` |

### Analytics Functions

| Function | Description | Example |
|----------|-------------|---------|
| `BONDYEARSTOMAT(isin)` | Years to maturity | `=BONDYEARSTOMAT("GB00BYZW3G56")` |
| `BONDMATURITYRANGE(from, to, country)` | Bonds maturing in range | `=BONDMATURITYRANGE("2025-01-01", "2030-12-31", "US")` |
| `BONDCOUPONFREQ(isin)` | Payment frequency | `=BONDCOUPONFREQ("US912810TM58")` → "Semi-annual" |
| `BONDISLINKER(isin)` | Is inflation-linked? | `=BONDISLINKER("GB00B3LZBF68")` → TRUE |

### Enterprise Functions

| Function | Description | Example |
|----------|-------------|---------|
| `BONDLINEAGE(isin, field)` | Data source attribution | `=BONDLINEAGE("DE0001102580", "coupon_rate")` |
| `BONDHISTORY(isin, limit)` | Change history | `=BONDHISTORY("US912810TM58", 10)` |
| `BONDACTIONS(type, days)` | Corporate actions | `=BONDACTIONS("MATURED", 30)` |

### Utility Functions

| Function | Description | Example |
|----------|-------------|---------|
| `BONDAPI_STATUS()` | Check API connection | `=BONDAPI_STATUS()` → "✓ Connected" |
| `BONDCACHE_CLEAR()` | Clear cache | `=BONDCACHE_CLEAR()` |
| `BONDCACHE_STATS()` | Cache performance | `=BONDCACHE_STATS()` |
| `BONDHELP(topic)` | Built-in help | `=BONDHELP("fields")` |
| `BONDISINVALID(isin)` | Validate ISIN | `=BONDISINVALID("GB00BYZW3G56")` → TRUE |

---

## 📋 Available Fields

Use these with `BONDSTATIC(isin, field)`:

| Field | Description | Shortcut |
|-------|-------------|----------|
| `coupon_rate` | Coupon rate (as %) | `coupon` |
| `maturity_date` | Maturity date | `maturity` |
| `issue_date` | Issue date | `issue` |
| `security_type` | NOMINAL or INDEX_LINKED | `type` |
| `coupon_frequency` | Payments per year | `freq` |
| `currency` | Currency code | |
| `country` | Country code | |
| `issuer` | Issuer name | |
| `name` | Full bond name | |
| `outstanding_amount` | Amount outstanding | |

---

## 🎯 Common Use Cases

### Build a Portfolio Tracker

| A (ISIN) | B (Coupon) | C (Maturity) | D (Years) |
|----------|------------|--------------|-----------|
| GB00BYZW3G56 | `=BONDSTATIC(A2, "coupon")` | `=BONDSTATIC(A2, "maturity")` | `=BONDYEARSTOMAT(A2)` |
| US912810TM58 | `=BONDSTATIC(A3, "coupon")` | `=BONDSTATIC(A3, "maturity")` | `=BONDYEARSTOMAT(A3)` |

### Find Bonds Maturing Soon

```excel
=BONDMATURITYRANGE("2025-01-01", "2025-12-31", "US")
```

### List All Inflation-Linked Bonds

```excel
=BONDLIST("GB", "INDEX_LINKED")    → UK index-linked gilts
=BONDLIST("US", "INDEX_LINKED")    → US TIPS
```

---

## ⚙️ Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `BONDMASTER_API_URL` | `http://127.0.0.1:8000` | API server URL |
| `BONDMASTER_CACHE_TTL` | `300` | Cache TTL in seconds |

### Remote API Server

If running the API on another machine:
```powershell
$env:BONDMASTER_API_URL = "http://bondserver.company.com:8000"
```

---

## 🏗️ Architecture

```
┌─────────────────────┐     HTTP/REST     ┌──────────────────────┐
│  Excel + xlOil      │ ◄───────────────► │  BondMaster API      │
│  (XLL Add-in)       │  localhost:8000   │  (bondmaster serve)  │
│                     │                   │                      │
│  • TTL Cache        │                   │  • SQLite Storage    │
│  • Input Validation │                   │  • Multi-source      │
│  • Error Formatting │                   │  • Enterprise MDM    │
└─────────────────────┘                   └──────────────────────┘
```

---

## 📄 License

MIT License

---

**Need help?** Type `=BONDHELP()` in Excel or open an issue on GitHub.
