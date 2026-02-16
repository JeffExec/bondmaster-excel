# Architecture

## System Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                           Excel Workbook                             │
│  ┌────────────────┐  ┌────────────────┐  ┌────────────────┐        │
│  │ =BONDSTATIC()  │  │ =BONDLIST()    │  │ =BONDINFO()    │        │
│  └───────┬────────┘  └───────┬────────┘  └───────┬────────┘        │
└──────────┼───────────────────┼───────────────────┼──────────────────┘
           │                   │                   │
           ▼                   ▼                   ▼
┌─────────────────────────────────────────────────────────────────────┐
│                     xlOil Python Plugin                              │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │                    bondmaster_excel.udfs                      │  │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────┐   │  │
│  │  │ Input       │  │ TTL Cache   │  │ Error Formatting    │   │  │
│  │  │ Validation  │  │ (5 min)     │  │ User-friendly msgs  │   │  │
│  │  └─────────────┘  └─────────────┘  └─────────────────────┘   │  │
│  └──────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────┘
                               │
                               │ HTTP/REST (localhost:8000)
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      BondMaster API Server                           │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │  FastAPI REST API                                             │  │
│  │  /bonds/{isin}  /bonds/search  /bonds/list  /stats           │  │
│  └──────────────────────────────────────────────────────────────┘  │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │  SQLite Storage                                               │  │
│  │  bonds | bond_history | bond_lineage | corporate_actions      │  │
│  └──────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────┘
```

## Components

### 1. Excel UDF Layer (`bondmaster_excel.udfs`)

**Purpose:** Expose bond data as native Excel functions via xlOil.

**Key Design Decisions:**
- **TTL Cache (5 min):** Bond data is semi-static; short TTL balances freshness vs. performance
- **Thread-safe HTTP client:** Excel may call functions from multiple threads
- **User-friendly errors:** Return `⚠️ Error message` instead of cryptic #VALUE!
- **Input validation:** Validate ISINs before hitting the API

**File:** `bondmaster_excel/udfs.py` (~1000 lines)

### 2. TTL Cache

**Purpose:** Reduce API calls for repeated lookups.

```python
class _TTLCache:
    """Thread-safe LRU cache with TTL expiration."""
    - maxsize: 500 entries
    - ttl: 300 seconds (5 minutes)
    - Hit rate tracking for observability
```

**Why not `functools.lru_cache`?**
- No TTL support (stale data forever)
- Not observable (no hit rate stats)
- Bond data changes, needs expiration

### 3. HTTP Client

**Design:** Singleton `httpx.Client` with connection pooling.

```python
_client: httpx.Client | None = None
_client_lock = threading.Lock()
```

**Why singleton?**
- Connection reuse (faster)
- Thread-safe initialization
- Proper cleanup on module unload

### 4. Function Categories

| Category | Functions | Purpose |
|----------|-----------|---------|
| Core | BONDSTATIC, BONDINFO, BONDLIST, BONDSEARCH, BONDCOUNT | Basic data retrieval |
| Analytics | BONDYEARSTOMAT, BONDMATURITYRANGE, BONDCOUPONFREQ, BONDISLINKER | Derived calculations |
| Enterprise | BONDLINEAGE, BONDHISTORY, BONDACTIONS | MDM/audit features |
| Utility | BONDAPI_STATUS, BONDCACHE_CLEAR, BONDCACHE_STATS, BONDHELP | Diagnostics |

## Data Flow

### Simple Lookup: `=BONDSTATIC("GB00BYZW3G56", "coupon_rate")`

```
1. Excel calls BONDSTATIC("GB00BYZW3G56", "coupon_rate")
2. Validate ISIN format (12 chars, check digit)
3. Check TTL cache for "GB00BYZW3G56"
   - HIT: Return cached bond data
   - MISS: Continue to step 4
4. HTTP GET http://localhost:8000/bonds/GB00BYZW3G56
5. Parse response, extract "coupon_rate" field
6. Store in cache with 5-min TTL
7. Return value to Excel (e.g., 1.5)
```

### Array Formula: `=BONDLIST("DE")`

```
1. Excel calls BONDLIST("DE")
2. HTTP GET http://localhost:8000/bonds/list?country=DE
3. Parse response as list of ISINs
4. Return as vertical array (spills in Excel 365)
```

## Error Handling Strategy

### User-Facing Errors

All errors returned to Excel cells use this format:
```
⚠️ [Short description]: [Actionable fix]
```

Examples:
- `⚠️ Invalid ISIN: Must be 12 characters`
- `⚠️ Bond not found: Check ISIN or run bondmaster fetch`
- `⚠️ API offline: Start server with 'bondmaster serve'`

### Internal Errors

- Logged to xlOil log file
- Never expose stack traces to users
- HTTP errors mapped to user-friendly messages

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `BONDMASTER_API_URL` | `http://127.0.0.1:8000` | API server URL |
| `BONDMASTER_CACHE_TTL` | `300` | Cache TTL in seconds |

### xlOil Configuration

Required in `xlOil.ini`:
```ini
[xlOil_Python]
LoadModules=["bondmaster_excel.udfs"]
```

## Dependencies

| Package | Purpose |
|---------|---------|
| `xloil` | Excel-Python bridge (XLL add-in framework) |
| `httpx` | Modern HTTP client with connection pooling |
| `bondmaster` | Core bond data library (API server) |

## Testing Strategy

- **Unit tests:** Mocked HTTP responses, cache behavior
- **Integration tests:** Real API server (pytest fixtures)
- **Coverage target:** >90%

## Performance Considerations

1. **Cache hit rate:** Target >80% in normal use
2. **Cold start:** First call slower (no cache)
3. **Batch operations:** Use BONDLIST for bulk, not individual calls
4. **Connection reuse:** Singleton HTTP client

## Security

- **Local only:** Default API binds to 127.0.0.1
- **No secrets in code:** All config via environment
- **Input validation:** ISINs validated before API calls
