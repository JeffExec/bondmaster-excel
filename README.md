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

### Step 1: Install Python packages

Open **Command Prompt** (or PowerShell) and run:

```cmd
pip install bondmaster bondmaster-excel xlOil httpx
```

### Step 2: Install xlOil into Excel

```cmd
python -m xloil install
```

This registers xlOil with Excel. You should see:
```
xlOil registered for Excel
```

### Step 3: Configure xlOil to load bondmaster-excel

Find your xlOil config file:
```cmd
python -c "import xloil; print(xloil.config_path())"
```

Open that file (usually `%APPDATA%\xlOil\xlOil.ini`) and add:

```ini
[xlOil]
Plugins = bondmaster_excel
```

### Step 4: Load bond data

```cmd
bondmaster fetch --seed-only
```

This downloads reference data for all supported markets (~1000 bonds).

### Step 5: Start the API server

```cmd
bondmaster serve
```

**Keep this terminal open** while using Excel. You should see:
```
INFO:     Uvicorn running on http://127.0.0.1:8000
```

### Step 6: Open Excel and test

1. Open Excel
2. In any cell, type: `=BONDAPI_STATUS()`
3. Press Enter

If you see **✓ Connected** — you're done! 🎉

If you see **✗ Disconnected** — make sure the API server is running (Step 5).

---

## 🔧 Troubleshooting Installation

### "xlOil not found" in Excel

1. Close Excel completely
2. Re-run: `python -m xloil install`
3. Restart Excel

### "bondmaster_excel not found"

Make sure you installed with pip:
```cmd
pip show bondmaster-excel
```

If not found, install it:
```cmd
pip install bondmaster-excel
```

### Functions show #NAME! error

xlOil isn't loading. Check:
1. Excel → File → Options → Add-ins
2. Look for "xlOil" in the list
3. If not there, re-run `python -m xloil install`

### "Cannot connect" error

The API server isn't running. Open a new terminal and run:
```cmd
bondmaster serve
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
```cmd
set BONDMASTER_API_URL=http://bondserver.company.com:8000
```

---

## 📁 Examples

See the `examples/` folder:
- `GettingStarted.md` — Step-by-step tutorial
- `PortfolioTemplate.csv` — Import as starting point

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
