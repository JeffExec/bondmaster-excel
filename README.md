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

- Windows 10/11 or Windows Server 2019+
- Microsoft Excel (desktop version, not web)
- Python 3.11 or 3.12
- Git (for cloning repositories)

> ⚠️ **Architecture must match:** If your Excel is 64-bit (most common), you need 64-bit Python.
> Check Excel: File → Account → About Excel → look for "64-bit" or "32-bit".
> Check Python: `python -c "import struct; print(struct.calcsize('P')*8, 'bit')"`

---

### Option A: Virtual Environment (Recommended for developers)

#### Step 1: Create a project folder and virtual environment

Open **PowerShell** and run:

```powershell
cd ~\PythonProjects   # or wherever you keep projects
mkdir bondmaster-excel
cd bondmaster-excel
python -m venv .venv
.venv\Scripts\activate
```

#### Step 2: Install packages from GitHub

> ⚠️ **Authentication Required:** The `bond-master` repository is private. You need a GitHub Personal Access Token (PAT) to install it.

**Option A: Use GitHub CLI (recommended)**
```powershell
# Install GitHub CLI if you haven't: https://cli.github.com/
gh auth login
pip install git+https://github.com/JeffExec/bond-master.git git+https://github.com/JeffExec/bondmaster-excel.git xlOil httpx
```

**Option B: Use a Personal Access Token**
```powershell
# Create a PAT at https://github.com/settings/tokens with 'repo' scope
# Then use it in the URL (replace YOUR_TOKEN):
pip install git+https://YOUR_TOKEN@github.com/JeffExec/bond-master.git git+https://github.com/JeffExec/bondmaster-excel.git xlOil httpx
```

**Option C: Clone and install locally**
```powershell
git clone https://github.com/JeffExec/bond-master.git
pip install ./bond-master xlOil httpx
pip install git+https://github.com/JeffExec/bondmaster-excel.git
```

#### Step 3: Install xlOil into Excel

```powershell
xloil install
```

You should see:
```
Installed C:\Users\<you>\AppData\Roaming\Microsoft\Excel\XLSTART\xlOil.xll
```

> **Note:** If the XLSTART folder doesn't exist, create it manually first:
> ```powershell
> mkdir "$env:APPDATA\Microsoft\Excel\XLSTART"
> ```

#### Step 4: Configure xlOil

Open the xlOil config file at `%APPDATA%\xlOil\xlOil.ini`.

Find and update these sections:

```toml
[xlOil_Python]
LoadModules=["xloil.xloil_ribbon", "bondmaster_excel.udfs"]

[[xlOil_Python.Environment]]
XLOIL_PYTHON_PATH='''C:\Users\<you>\PythonProjects\bondmaster-excel\.venv\Lib\site-packages'''

[[xlOil_Python.Environment]]
PYTHONEXECUTABLE='''C:\Users\<you>\PythonProjects\bondmaster-excel\.venv\Scripts\python.exe'''
```

> **Important:** Use triple single quotes `'''` for paths with spaces (TOML literal strings).

#### Step 5: Load bond data

```powershell
bondmaster fetch --seed-only
```

#### Step 6: Start the API server

```powershell
bondmaster serve
```

Keep this terminal open while using Excel.

#### Step 7: Open Excel and test

1. Open Excel normally (from Start menu or double-click a file)
2. Look for the **xlOil Py** tab in the ribbon
3. In any cell: `=BONDAPI_STATUS()`
4. If you see **✓ Connected** — you're done! 🎉

---

### Option B: System-wide Install (Simpler for servers/non-developers)

For Windows servers or users who don't need isolated environments:

> ⚠️ **Authentication Required:** See Step 2 above for GitHub authentication options (the `bond-master` repo is private).

```powershell
# Install packages (uses user site-packages if no admin)
# Use gh auth login first, or include your PAT in the URL
pip install --user git+https://github.com/JeffExec/bond-master.git git+https://github.com/JeffExec/bondmaster-excel.git xlOil httpx
```

> ⚠️ **PATH issue:** pip installs scripts to a folder not in PATH. Find your Scripts folder:
> ```powershell
> pip show xloil | Select-String "Location"
> # Example output: Location: C:\Users\<you>\AppData\Roaming\Python\Python312\site-packages
> # Scripts are at: C:\Users\<you>\AppData\Roaming\Python\Python312\Scripts
> ```

```powershell
# Set Scripts path for this session (adjust path from above)
$scripts = "$env:APPDATA\Python\Python312\Scripts"

# For user installs, set XLOIL_BIN_DIR (binaries are in a different location)
$env:XLOIL_BIN_DIR = "$env:APPDATA\Python\share\xloil"

# Close Excel first, then install xlOil
& "$scripts\xloil.exe" install
```

**Edit config:** Open `%APPDATA%\xlOil\xlOil.ini` and update:

```toml
[xlOil_Python]
LoadModules=["xloil.xloil_ribbon", "bondmaster_excel.udfs"]

[[xlOil_Python.Environment]]
# Point to YOUR user site-packages (not a venv)
XLOIL_PYTHON_PATH='''C:\Users\<you>\AppData\Roaming\Python\Python312\site-packages'''
```

```powershell
# Load data and start server
& "$scripts\bondmaster.exe" fetch --seed-only
& "$scripts\bondmaster.exe" serve
```

> 💡 **Tip:** To avoid typing full paths, add Scripts to PATH permanently:
> ```powershell
> [Environment]::SetEnvironmentVariable("PATH", "$env:PATH;$scripts", "User")
> # Restart PowerShell after this
> ```

---

## 🔧 Troubleshooting Installation

### xlOil ribbon tab doesn't appear

1. **Check Excel Add-ins:** File → Options → Add-ins → Manage "Excel Add-ins" → Go
2. Is xlOil.xll listed and checked? If unchecked, Excel disabled it.
3. **Re-run install:**
   ```powershell
   xloil install
   ```
4. **Restart Excel** (close all Excel windows completely)

### Functions show #NAME? error

The Python module failed to load. Check the xlOil log:
- Click **Open Log** in the xlOil ribbon, OR
- Check `%APPDATA%\xlOil\xloil.log`

**Common causes:**

1. **Wrong LoadModules value:**
   ```toml
   # ❌ Wrong
   LoadModules=["bondmaster_excel"]
   
   # ✅ Correct
   LoadModules=["xloil.xloil_ribbon", "bondmaster_excel.udfs"]
   ```

2. **XLOIL_PYTHON_PATH not set:** Must point to your `site-packages` folder.

3. **Python can't import the module:** Test in terminal:
   ```powershell
   python -c "import bondmaster_excel.udfs; print('OK')"
   ```

4. **Missing xlOil DLLs:** If log shows DLL errors, copy the supporting files:
   ```powershell
   $src = (pip show xloil | Select-String "Location").Line.Split(": ")[1]
   $src = "$src\share\xloil"
   $dst = "$env:APPDATA\Microsoft\Excel\XLSTART"
   
   Copy-Item "$src\xlOil.dll" $dst
   Copy-Item "$src\xlOil_Python.dll" $dst
   Copy-Item "$src\xlOil_Utils.dll" $dst
   
   # Copy the PYD matching your Python version:
   # Python 3.12: xlOil_Python312.pyd
   # Python 3.11: xlOil_Python311.pyd
   Copy-Item "$src\xlOil_Python312.pyd" $dst
   ```

### "Error parsing settings file" on Excel startup

Your `xlOil.ini` has a TOML syntax error.

**Quick fix:** Start fresh from the default config (⚠️ backs up your existing config first):
```powershell
$src = (pip show xloil | Select-String "Location").Line.Split(": ")[1]
$cfg = "$env:APPDATA\xlOil\xlOil.ini"
if (Test-Path $cfg) { Copy-Item $cfg "$cfg.bak" }  # Backup existing
Copy-Item "$src\share\xloil\xlOil.ini" $cfg
```

Then add `bondmaster_excel.udfs` to the `LoadModules` line.

**Common syntax issues:**
- Use `'''triple quotes'''` for paths with spaces
- Use `["array", "syntax"]` for lists
- Section names are case-sensitive: `[xlOil_Python]` not `[xloil_python]`

### "Cannot connect" error in cells

The API server isn't running.

```powershell
# In a new terminal:
.venv\Scripts\activate  # if using venv
bondmaster serve
```

### No log file created at all

xlOil core isn't loading. Check:

1. **Architecture mismatch:** 64-bit Excel needs 64-bit Python
2. **Missing Visual C++ Redistributable:** Install from [Microsoft](https://aka.ms/vs/17/release/vc_redist.x64.exe)
3. **XLL blocked by Windows:** Right-click xlOil.xll → Properties → Unblock

### xlOil loads but bondmaster functions missing

Check the log for import errors. Common issues:

1. **bondmaster package not installed:**
   ```powershell
   pip install git+https://github.com/JeffExec/bond-master.git
   ```

2. **httpx not installed:**
   ```powershell
   pip install httpx
   ```

### xlOil Log shows "TypeError: func() got an unexpected keyword argument 'category'"

You have an older bondmaster-excel with xlOil 0.21+. The `category` parameter was removed in xlOil 0.21.

**Fix:** Update bondmaster-excel:
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
| `BONDNAMESEARCH(query)` | Search by bond name (v2.0) | `=BONDNAMESEARCH("OATEI 2030")` |
| `BONDCOUNT(country)` | Count bonds | `=BONDCOUNT("GB")` |

> **v2.0 Auto-Lookup:** When you request a bond that isn't in the database, the cell will show "🔄 Looking up..." while the API fetches it in the background. Refresh the cell after a few seconds to see the result.

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
