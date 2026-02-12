# BondMaster Excel Add-in

**Native Excel functions for government bond reference data.** No macro warnings. Blazing fast. Works offline.

```excel
=BONDSTATIC("US912810TM58", "coupon_rate")  →  4.625
=BONDYEARSTOMAT("GB00BYZW3G56")             →  5.42
=BONDLIST("DE")                              →  [list of German bond ISINs]
```

## ✨ Features

- **20+ Excel functions** covering all bond-master capabilities
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

## 🚀 Quick Start

### 1. Install
```bash
pip install bondmaster bondmaster-excel xloil httpx
python -m xloil install
```

### 2. Load Data
```bash
bondmaster fetch --seed-only
```

### 3. Start API
```bash
bondmaster serve
```

### 4. Use in Excel
```excel
=BONDAPI_STATUS()                        → ✓ Connected
=BONDSTATIC("US912810TM58", "coupon")    → 4.625
=BONDLIST("GB")                          → [UK gilt ISINs]
```

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

## 🎯 Common Use Cases

### Build a Portfolio Tracker

```excel
| A (ISIN)        | B (Coupon)                    | C (Maturity)                      | D (Years)                |
|-----------------|-------------------------------|-----------------------------------|--------------------------|
| GB00BYZW3G56    | =BONDSTATIC(A2, "coupon")     | =BONDSTATIC(A2, "maturity")       | =BONDYEARSTOMAT(A2)      |
| US912810TM58    | =BONDSTATIC(A3, "coupon")     | =BONDSTATIC(A3, "maturity")       | =BONDYEARSTOMAT(A3)      |
```

### Find Bonds Maturing Soon

```excel
=BONDMATURITYRANGE("2025-01-01", "2025-12-31", "US")
```
Returns ISIN and maturity date for all matching bonds.

### List All Inflation-Linked Bonds

```excel
=BONDLIST("GB", "INDEX_LINKED")    → UK index-linked gilts
=BONDLIST("US", "INDEX_LINKED")    → US TIPS
```

### Data Quality Check

```excel
=BONDLINEAGE("DE0001102580", "coupon_rate")
→ "deutsche_finanzagentur (confidence: 95%)"
```

## ⚙️ Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `BONDMASTER_API_URL` | `http://127.0.0.1:8000` | API server URL |
| `BONDMASTER_CACHE_TTL` | `300` | Cache TTL in seconds |

### Custom API URL

For remote servers:
```powershell
set BONDMASTER_API_URL=http://bondserver.company.com:8000
```

## 🔧 Troubleshooting

### "✗ Disconnected" Error

**Problem:** API server not running

**Solution:**
```bash
bondmaster serve
```
Keep terminal open while using Excel.

### #VALUE! Errors

**Problem:** Invalid input

**Check:**
- ISIN is 12 characters
- Field name is spelled correctly
- Run `=BONDHELP("fields")` to see valid fields

### Slow First Lookup

**Normal behavior.** First lookup fetches from API. Subsequent lookups use cache (5-minute TTL).

Check cache stats:
```excel
=BONDCACHE_STATS()
→ "Size: 50/500 | Hit Rate: 95% | TTL: 300s"
```

### Outdated Data

Clear cache and retry:
```excel
=BONDCACHE_CLEAR()
```

## 📁 Examples

See the `examples/` folder:
- `GettingStarted.md` — Step-by-step tutorial
- `PortfolioTemplate.csv` — Import as starting point

## 🏗️ Architecture

```
┌─────────────────────┐     HTTP/REST     ┌──────────────────────┐
│  Excel + xlOil      │ ◄───────────────► │  BondMaster API      │
│  (XLL Add-in)       │  localhost:8000   │  (Python/FastAPI)    │
│                     │                   │                      │
│  • TTL Cache        │                   │  • SQLite Storage    │
│  • Input Validation │                   │  • Multi-source      │
│  • Error Formatting │                   │  • Enterprise MDM    │
└─────────────────────┘                   └──────────────────────┘
```

## 🧪 Development

```bash
# Clone
git clone https://github.com/JeffExec/bondmaster-excel.git
cd bondmaster-excel

# Install dev dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Type check
mypy bondmaster_excel/
```

## 📄 License

MIT License

---

**Need help?** Use `=BONDHELP()` in Excel or open an issue on GitHub.
