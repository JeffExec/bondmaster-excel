# BondMaster Excel - Getting Started

## Quick Setup (5 minutes)

### Step 1: Install BondMaster
```powershell
pip install bondmaster bondmaster-excel
```

### Step 2: Load Bond Data
```bash
bondmaster fetch --seed-only
```

### Step 3: Start the API Server
```bash
bondmaster serve
```
Leave this running in a terminal window.

### Step 4: Verify in Excel
Open Excel and enter:
```
=BONDAPI_STATUS()
```
Should show: `✓ Connected`

---

## Your First Formulas

### Get bond coupon rate
```excel
=BONDSTATIC("US912810TM58", "coupon_rate")
```
Returns: `4.625` (as a percentage)

### Get maturity date
```excel
=BONDSTATIC("US912810TM58", "maturity_date")
```
Returns: `2054-02-15`

### Get full bond info
```excel
=BONDINFO("GB00BYZW3G56", TRUE)
```
Returns: Header row + data row spanning multiple columns

---

## Common Use Cases

### Portfolio Analysis

**List all UK gilts:**
```excel
=BONDLIST("GB")
```

**List only UK inflation-linked gilts:**
```excel
=BONDLIST("GB", "INDEX_LINKED")
```

**Count US Treasury bonds:**
```excel
=BONDCOUNT("US")
```

### Maturity Analysis

**Years to maturity:**
```excel
=BONDYEARSTOMAT("GB00BYZW3G56")
```

**Bonds maturing in 2025:**
```excel
=BONDMATURITYRANGE("2025-01-01", "2025-12-31", "US")
```

### Data Validation

**Check ISIN is valid:**
```excel
=BONDISINVALID("GB00BYZW3G56")
```

**Check if bond is inflation-linked:**
```excel
=BONDISLINKER("GB00B3LZBF68")
```

---

## Building a Bond Portfolio Tracker

Create a simple portfolio tracker:

| A (ISIN) | B (Coupon) | C (Maturity) | D (Years) | E (Type) |
|----------|------------|--------------|-----------|----------|
| GB00BYZW3G56 | =BONDSTATIC(A2,"coupon") | =BONDSTATIC(A2,"maturity") | =BONDYEARSTOMAT(A2) | =BONDSTATIC(A2,"type") |

Then copy formulas down for each bond in your portfolio.

---

## Available Fields

| Field | Description | Example |
|-------|-------------|---------|
| `coupon_rate` | Coupon as % | 1.5 |
| `maturity_date` | Maturity | 2030-07-22 |
| `issue_date` | Issue date | 2020-07-22 |
| `issuer` | Issuer name | UK DMO |
| `currency` | Currency | GBP |
| `security_type` | Type | NOMINAL |
| `country` | Country code | GB |
| `name` | Full name | UK Treasury 1.5% 2030 |
| `coupon_frequency` | Payments/year | 2 |

**Shortcuts:** `coupon` → `coupon_rate`, `maturity` → `maturity_date`, `type` → `security_type`

---

## Country Codes

| Code | Country | Bonds |
|------|---------|-------|
| US | United States | 400+ |
| GB | United Kingdom | 100+ |
| DE | Germany | 90+ |
| FR | France | 30+ |
| IT | Italy | 200+ |
| ES | Spain | 20+ |
| JP | Japan | 30+ |
| NL | Netherlands | 15+ |

---

## All Functions

### Core Data
- `BONDSTATIC(isin, field)` - Get a single field
- `BONDINFO(isin, headers)` - Get all fields
- `BONDLIST(country, type)` - List ISINs
- `BONDSEARCH(field, value, ...)` - Search with filters
- `BONDCOUNT(country)` - Count bonds

### Analytics
- `BONDYEARSTOMAT(isin)` - Years to maturity
- `BONDMATURITYRANGE(from, to, country)` - Bonds in date range
- `BONDCOUPONFREQ(isin)` - Payment frequency text
- `BONDISLINKER(isin)` - Check if inflation-linked

### Enterprise
- `BONDLINEAGE(isin, field)` - Data source attribution
- `BONDHISTORY(isin)` - Change history
- `BONDACTIONS(type, days)` - Corporate actions

### Utilities
- `BONDAPI_STATUS()` - Check connection
- `BONDCACHE_CLEAR()` - Clear cache
- `BONDCACHE_STATS()` - Cache statistics
- `BONDHELP(topic)` - Help
- `BONDISINVALID(isin)` - Validate ISIN

---

## Troubleshooting

### "✗ Disconnected" error
1. Open terminal/command prompt
2. Run: `bondmaster serve`
3. Keep it running while using Excel

### #VALUE! errors
- Check ISIN format: 12 characters, starts with country code
- Check field name spelling
- Use `=BONDHELP("fields")` to see valid fields

### Slow responses
- First call may be slow (cache warming)
- Subsequent calls are cached (5 min TTL)
- Use `=BONDCACHE_STATS()` to check hit rate

### Data seems outdated
- Run `=BONDCACHE_CLEAR()` to force refresh
- Then retry your formula

---

## Need More Help?

```excel
=BONDHELP()           → Overview
=BONDHELP("fields")   → All field names
=BONDHELP("countries") → Country codes
=BONDHELP("functions") → All functions
```
