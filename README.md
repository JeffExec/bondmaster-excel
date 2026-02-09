# BondMaster Excel Add-in

Excel add-in for accessing government bond reference data. Native XLL format - no macro warnings.

## Features

- **Native Excel functions** - Use like any built-in function
- **No macro security warnings** - XLL add-in format
- **Fast** - Uses xlOil (2000x faster than COM-based solutions)
- **Offline capable** - Works with local BondMaster API

## Supported Markets

| Market | Bonds | Coverage |
|--------|-------|----------|
| 🇬🇧 UK Gilts | 100+ | Full |
| 🇺🇸 US Treasury | 400+ | Full (requires internet) |
| 🇩🇪 Germany Bunds | 20+ | Major issues |
| 🇯🇵 Japan JGBs | 25+ | Major issues |
| 🇫🇷 France OATs | 20+ | Major issues |
| 🇮🇹 Italy BTPs | 20+ | Major issues |
| 🇪🇸 Spain Bonos | 20+ | Major issues |
| 🇳🇱 Netherlands DSLs | 10+ | Major issues |

## Installation

### Quick Install (Windows)

```powershell
# Run PowerShell as Administrator
powershell -ExecutionPolicy Bypass -File scripts\install.ps1
```

### Manual Install

```bash
# 1. Install packages
pip install bondmaster xloil httpx

# 2. Install xlOil Excel add-in
python -m xloil install

# 3. Load bond data
bondmaster fetch --seed-only

# 4. Start API server
bondmaster serve
```

## Usage

### Start the API Server

Before using Excel functions, start the BondMaster API:

```bash
bondmaster serve
```

Or double-click "Start BondMaster API.bat" on your Desktop.

### Excel Functions

#### BONDSTATIC - Get a single field

```excel
=BONDSTATIC("GB00BYZW3G56", "coupon_rate")     → 1.5
=BONDSTATIC("GB00BYZW3G56", "maturity_date")   → 2026-07-22
=BONDSTATIC("GB00BYZW3G56", "issuer")          → UK DMO
=BONDSTATIC("GB00BYZW3G56", "currency")        → GBP
=BONDSTATIC("GB00BYZW3G56", "security_type")   → NOMINAL
```

**Available fields:**
- `isin`, `cusip`, `name`
- `country`, `issuer`, `currency`
- `coupon_rate`, `coupon_frequency`
- `maturity_date`, `issue_date`, `first_coupon_date`
- `security_type` (NOMINAL, INDEX_LINKED)
- `outstanding_amount`

#### BONDINFO - Get all fields as a row

```excel
=BONDINFO("GB00BYZW3G56")           → Spills across columns
=BONDINFO("GB00BYZW3G56", TRUE)     → Includes header row
```

#### BONDLIST - Get ISINs for a country

```excel
=BONDLIST("GB")                    → All UK gilt ISINs (spills down)
=BONDLIST("US", "INDEX_LINKED")    → US TIPS only
=BONDLIST("DE", "NOMINAL")         → German nominal bonds
```

#### BONDSEARCH - Search with filters

```excel
=BONDSEARCH("country", "US")
=BONDSEARCH("country", "GB", "security_type", "INDEX_LINKED")
```

#### BONDCOUNT - Count bonds

```excel
=BONDCOUNT()       → Total bonds in database
=BONDCOUNT("US")   → US bonds only
```

#### Utility Functions

```excel
=BONDAPI_STATUS()      → "Connected" or error message
=BONDCACHE_CLEAR()     → Clear cache (after data updates)
```

## Architecture

```
┌─────────────────┐     HTTP/REST    ┌──────────────────┐
│  Excel + xlOil  │ ◄──────────────► │  BondMaster API  │
│  (XLL Add-in)   │   localhost:8000 │  (Python/FastAPI)│
└─────────────────┘                  └──────────────────┘
                                              │
                                              ▼
                                     ┌──────────────────┐
                                     │  SQLite + Seed   │
                                     │      Data        │
                                     └──────────────────┘
```

## Troubleshooting

### "Disconnected" or #N/A errors

1. Ensure BondMaster API is running: `bondmaster serve`
2. Check http://127.0.0.1:8000/health in browser
3. Clear cache: `=BONDCACHE_CLEAR()`

### Add-in not loading

1. Open Excel → File → Options → Add-ins
2. Manage: COM Add-ins → Go
3. Check "xlOil" is listed and enabled
4. If not, run: `python -m xloil install`

### Functions not appearing

1. Restart Excel completely
2. Type `=BOND` and check autocomplete
3. Functions are in the "BondMaster" category in Insert Function

## Development

```bash
# Clone repository
git clone https://github.com/JeffExec/bondmaster-excel.git
cd bondmaster-excel

# Install in development mode
pip install -e ".[dev]"

# Run tests
pytest
```

## License

MIT License
