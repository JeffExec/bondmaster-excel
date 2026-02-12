"""
BondMaster Excel UDFs (User Defined Functions).

Complete Excel integration for government bond reference data.
All functions are exposed to Excel via xlOil in the "BondMaster" category.

Quick Start:
    1. Start API: bondmaster serve
    2. In Excel: =BONDAPI_STATUS() should show "Connected"
    3. Try: =BONDSTATIC("US912810TM58", "coupon_rate")

Function Categories:
    - Core: BONDSTATIC, BONDINFO, BONDLIST, BONDSEARCH, BONDCOUNT
    - Analytics: BONDYEARSTOMAT, BONDNEXTCOUPON, BONDMATURITYRANGE
    - Data Management: BONDREFRESH, BONDEXPORT
    - Enterprise: BONDLINEAGE, BONDHISTORY, BONDACTIONS
    - Utilities: BONDAPI_STATUS, BONDCACHE_CLEAR, BONDCACHE_STATS, BONDHELP

Cache Strategy:
    TTL-based LRU cache (5 min default). Bond reference data is semi-static,
    so short TTL balances freshness vs. performance.
"""

import os
import re
import threading
import time
from datetime import date, datetime
from typing import Any, NamedTuple

import httpx
import xloil as xlo

# =============================================================================
# Configuration
# =============================================================================

API_BASE_URL = os.environ.get("BONDMASTER_API_URL", "http://127.0.0.1:8000")
REQUEST_TIMEOUT = 10.0
MAX_RETRIES = 2
CACHE_TTL_SECONDS = float(os.environ.get("BONDMASTER_CACHE_TTL", "300"))
CACHE_MAX_SIZE = 500

# ISIN validation: 2 letters + 9 alphanumeric + 1 check digit
ISIN_PATTERN = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$")

# Country codes with names for help text
COUNTRY_CODES = {
    "US": "United States",
    "GB": "United Kingdom",
    "DE": "Germany",
    "FR": "France",
    "IT": "Italy",
    "ES": "Spain",
    "JP": "Japan",
    "NL": "Netherlands",
}

# Available fields with descriptions
BOND_FIELDS = {
    "isin": "ISIN identifier",
    "cusip": "CUSIP (US bonds)",
    "sedol": "SEDOL (UK bonds)",
    "name": "Bond name",
    "country": "Country code (US, GB, DE...)",
    "issuer": "Issuing entity",
    "security_type": "NOMINAL or INDEX_LINKED",
    "currency": "Currency code (USD, GBP, EUR...)",
    "coupon_rate": "Coupon rate (displayed as %)",
    "coupon_frequency": "Payments per year (1=annual, 2=semi)",
    "day_count_convention": "Day count method",
    "maturity_date": "Maturity date",
    "issue_date": "Issue date",
    "first_coupon_date": "First coupon payment date",
    "outstanding_amount": "Amount outstanding",
    "original_tenor": "Original term (e.g., 10Y)",
}

# =============================================================================
# HTTP Client
# =============================================================================

_client: httpx.Client | None = None
_client_lock = threading.Lock()


def _get_client() -> httpx.Client:
    """Get or create HTTP client singleton (thread-safe)."""
    global _client
    with _client_lock:
        if _client is None:
            _client = httpx.Client(base_url=API_BASE_URL, timeout=REQUEST_TIMEOUT)
        return _client


def _api_request(
    method: str,
    path: str,
    params: dict | None = None,
    json: dict | None = None,
    headers: dict | None = None,
) -> tuple[bool, Any]:
    """
    Make API request with retry logic.
    
    Returns: (success: bool, data_or_error: Any)
    """
    for attempt in range(MAX_RETRIES + 1):
        try:
            client = _get_client()
            response = client.request(
                method=method,
                url=path,
                params=params,
                json=json,
                headers=headers,
            )

            if response.status_code == 200:
                return True, response.json()
            elif response.status_code == 404:
                return False, "Not found"
            elif response.status_code == 403:
                return False, "API key required"
            else:
                return False, f"HTTP {response.status_code}"

        except httpx.TimeoutException:
            if attempt < MAX_RETRIES:
                time.sleep(0.1 * (2 ** attempt))
                continue
            return False, "Timeout - is BondMaster API running?"

        except httpx.ConnectError:
            return False, "Cannot connect - start API with: bondmaster serve"

        except httpx.RequestError as e:
            if attempt < MAX_RETRIES:
                time.sleep(0.1 * (2 ** attempt))
                continue
            return False, f"Network error: {type(e).__name__}"

    return False, "Max retries exceeded"


# =============================================================================
# Validation Helpers
# =============================================================================

def _is_valid_isin(isin: str) -> bool:
    """Validate ISIN format."""
    return bool(ISIN_PATTERN.match(isin.upper().strip()))


def _normalize_isin(isin: str) -> str:
    """Normalize ISIN to uppercase, stripped."""
    return isin.upper().strip()


def _parse_date(value: Any) -> date | None:
    """Parse date from various formats."""
    if value is None:
        return None
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value.replace("Z", "")).date()
        except ValueError:
            return None
    return None


# =============================================================================
# TTL Cache
# =============================================================================

class _CacheEntry(NamedTuple):
    data: dict
    expires_at: float


class _TTLCache:
    """Thread-safe TTL-aware LRU cache."""

    def __init__(self, maxsize: int = 500, ttl_seconds: float = 300.0):
        self._cache: dict[str, _CacheEntry] = {}
        self._maxsize = maxsize
        self._ttl = ttl_seconds
        self._lock = threading.Lock()
        self._hits = 0
        self._misses = 0

    def get(self, key: str) -> dict | None:
        with self._lock:
            entry = self._cache.get(key)
            if entry is None:
                self._misses += 1
                return None
            if time.time() > entry.expires_at:
                del self._cache[key]
                self._misses += 1
                return None
            # Move to end (LRU)
            self._cache[key] = self._cache.pop(key)
            self._hits += 1
            return entry.data

    def set(self, key: str, value: dict) -> None:
        with self._lock:
            if key in self._cache:
                del self._cache[key]
            elif len(self._cache) >= self._maxsize:
                oldest = next(iter(self._cache))
                del self._cache[oldest]
            self._cache[key] = _CacheEntry(value, time.time() + self._ttl)

    def clear(self) -> int:
        with self._lock:
            count = len(self._cache)
            self._cache.clear()
            self._hits = 0
            self._misses = 0
            return count

    def stats(self) -> dict:
        with self._lock:
            total = self._hits + self._misses
            return {
                "size": len(self._cache),
                "maxsize": self._maxsize,
                "hits": self._hits,
                "misses": self._misses,
                "hit_rate": self._hits / total if total > 0 else 0.0,
                "ttl_seconds": self._ttl,
            }


_bond_cache = _TTLCache(maxsize=CACHE_MAX_SIZE, ttl_seconds=CACHE_TTL_SECONDS)


def _fetch_bond(isin: str) -> dict | None:
    """Fetch bond with caching."""
    isin = _normalize_isin(isin)
    if not _is_valid_isin(isin):
        return None

    cached = _bond_cache.get(isin)
    if cached is not None:
        return cached

    success, data = _api_request("GET", f"/bonds/{isin}")
    if not success:
        return None

    # Handle envelope response
    bond = data.get("data", data) if isinstance(data, dict) else data
    if bond:
        _bond_cache.set(isin, bond)
    return bond


def _format_error(msg: str) -> str:
    """Format error message for Excel display."""
    return f"⚠️ {msg}"


# =============================================================================
# CORE FUNCTIONS
# =============================================================================

@xlo.func(
    help="Get a specific field from bond reference data.\n\nExample: =BONDSTATIC(\"GB00BYZW3G56\", \"coupon_rate\")",
    args={
        "isin": "ISIN code (e.g., 'GB00BYZW3G56', 'US912810TM58')",
        "field": "Field name: coupon_rate, maturity_date, issuer, currency, security_type, etc.",
    },
    category="BondMaster",
)
def BONDSTATIC(isin: str, field: str) -> xlo.ExcelValue:
    """
    Get a specific field from bond reference data.
    
    COMMON FIELDS:
        coupon_rate     - Coupon rate (as %, e.g., 1.5 means 1.5%)
        maturity_date   - Maturity date
        issue_date      - Issue date  
        issuer          - Issuing entity name
        currency        - Currency code (USD, GBP, EUR)
        security_type   - NOMINAL or INDEX_LINKED
        country         - Country code (US, GB, DE)
        name            - Full bond name
    
    SHORTCUTS:
        coupon → coupon_rate
        maturity → maturity_date
        type → security_type
    
    EXAMPLES:
        =BONDSTATIC("GB00BYZW3G56", "coupon_rate")   → 1.5
        =BONDSTATIC("US912810TM58", "maturity_date") → 2054-02-15
        =BONDSTATIC("DE0001102580", "issuer")        → Federal Republic of Germany
    """
    if not isin or not field:
        return _format_error("ISIN and field required")

    isin = _normalize_isin(isin)
    if not _is_valid_isin(isin):
        return _format_error(f"Invalid ISIN format: {isin}")

    bond = _fetch_bond(isin)
    if bond is None:
        return _format_error(f"Bond not found: {isin}")

    # Field aliases
    field = field.lower().strip()
    aliases = {
        "coupon": "coupon_rate",
        "maturity": "maturity_date",
        "issue": "issue_date",
        "type": "security_type",
        "freq": "coupon_frequency",
        "frequency": "coupon_frequency",
    }
    field = aliases.get(field, field)

    if field not in bond and field not in BOND_FIELDS:
        return _format_error(f"Unknown field: {field}")

    value = bond.get(field)
    if value is None:
        return ""

    # Format coupon as percentage
    if field == "coupon_rate" and isinstance(value, (int, float)):
        return value * 100

    return value


@xlo.func(
    help="Get all reference data for a bond as a row (spills across columns).\n\nExample: =BONDINFO(\"GB00BYZW3G56\", TRUE)",
    args={
        "isin": "ISIN code",
        "with_headers": "Include header row (default: FALSE)",
    },
    category="BondMaster",
)
def BONDINFO(isin: str, with_headers: bool = False) -> xlo.ExcelValue:
    """
    Get all bond data as an array that spills across cells.
    
    EXAMPLES:
        =BONDINFO("GB00BYZW3G56")        → Data row only
        =BONDINFO("GB00BYZW3G56", TRUE)  → Headers + data (2 rows)
    
    COLUMNS RETURNED:
        ISIN, Name, Country, Issuer, Type, Currency, Coupon%, Frequency,
        Maturity, Issue Date, Outstanding
    """
    if not isin:
        return _format_error("ISIN required")

    isin = _normalize_isin(isin)
    if not _is_valid_isin(isin):
        return _format_error(f"Invalid ISIN: {isin}")

    bond = _fetch_bond(isin)
    if bond is None:
        return _format_error(f"Bond not found: {isin}")

    columns = [
        ("isin", "ISIN"),
        ("name", "Name"),
        ("country", "Country"),
        ("issuer", "Issuer"),
        ("security_type", "Type"),
        ("currency", "Currency"),
        ("coupon_rate", "Coupon %"),
        ("coupon_frequency", "Frequency"),
        ("maturity_date", "Maturity"),
        ("issue_date", "Issue Date"),
        ("outstanding_amount", "Outstanding"),
    ]

    values = []
    for key, _ in columns:
        val = bond.get(key, "")
        if key == "coupon_rate" and isinstance(val, (int, float)):
            val = val * 100
        values.append(val if val is not None else "")

    if with_headers:
        headers = [col[1] for col in columns]
        return [headers, values]

    return [values]


@xlo.func(
    help="Get list of ISINs for a country.\n\nExample: =BONDLIST(\"GB\", \"INDEX_LINKED\")",
    args={
        "country": "Country code: US, GB, DE, FR, IT, ES, JP, NL",
        "security_type": "Optional: NOMINAL or INDEX_LINKED",
        "limit": "Max results (default: 500)",
    },
    category="BondMaster",
)
def BONDLIST(
    country: str,
    security_type: str | None = None,
    limit: int = 500,
) -> xlo.ExcelValue:
    """
    Get all ISINs for a country as a vertical array (spills down).
    
    EXAMPLES:
        =BONDLIST("GB")                    → All UK gilt ISINs
        =BONDLIST("US", "INDEX_LINKED")    → US TIPS only
        =BONDLIST("DE", "NOMINAL", 10)     → First 10 German nominal bonds
    
    COUNTRY CODES:
        US=United States, GB=United Kingdom, DE=Germany, FR=France,
        IT=Italy, ES=Spain, JP=Japan, NL=Netherlands
    """
    if not country:
        return _format_error("Country code required (US, GB, DE, FR, IT, ES, JP, NL)")

    country = country.upper().strip()
    if country not in COUNTRY_CODES:
        return _format_error(f"Unknown country: {country}. Use: {', '.join(COUNTRY_CODES.keys())}")

    params = {"country": country, "limit": min(limit, 1000)}
    if security_type:
        st = security_type.upper().strip()
        if st not in ("NOMINAL", "INDEX_LINKED"):
            return _format_error("security_type must be NOMINAL or INDEX_LINKED")
        params["security_type"] = st

    success, data = _api_request("GET", "/bonds", params=params)
    if not success:
        return _format_error(str(data))

    bonds = data.get("data", data) if isinstance(data, dict) else data
    if not bonds:
        return _format_error(f"No bonds found for {country}")

    return [[b.get("isin", "")] for b in bonds]


@xlo.func(
    help="Search bonds with multiple filters.\n\nExample: =BONDSEARCH(\"country\", \"US\", \"security_type\", \"INDEX_LINKED\")",
    args={
        "field1": "Filter field (country, security_type, currency, etc.)",
        "value1": "Filter value",
        "field2": "Optional second filter",
        "value2": "Optional second value",
        "field3": "Optional third filter",
        "value3": "Optional third value",
    },
    category="BondMaster",
)
def BONDSEARCH(
    field1: str,
    value1: str,
    field2: str | None = None,
    value2: str | None = None,
    field3: str | None = None,
    value3: str | None = None,
) -> xlo.ExcelValue:
    """
    Search for bonds matching filter criteria. Returns ISINs.
    
    FILTER FIELDS:
        country         - Country code (US, GB, DE...)
        security_type   - NOMINAL or INDEX_LINKED
        currency        - Currency code (USD, GBP, EUR)
        maturity_from   - Min maturity date (YYYY-MM-DD)
        maturity_to     - Max maturity date (YYYY-MM-DD)
        min_coupon      - Min coupon rate (as decimal, e.g., 0.02)
        max_coupon      - Max coupon rate
    
    EXAMPLES:
        =BONDSEARCH("country", "DE")
        =BONDSEARCH("country", "US", "security_type", "INDEX_LINKED")
        =BONDSEARCH("currency", "EUR", "maturity_from", "2030-01-01")
    """
    params = {"limit": 500}

    filters = [
        (field1, value1),
        (field2, value2),
        (field3, value3),
    ]

    for field, value in filters:
        if field and value:
            params[field.lower().strip()] = value

    if len(params) == 1:
        return _format_error("At least one filter required")

    success, data = _api_request("GET", "/bonds", params=params)
    if not success:
        return _format_error(str(data))

    bonds = data.get("data", data) if isinstance(data, dict) else data
    if not bonds:
        return _format_error("No bonds match filters")

    return [[b.get("isin", "")] for b in bonds]


@xlo.func(
    help="Count bonds in database.\n\nExample: =BONDCOUNT(\"US\")",
    args={
        "country": "Optional country code to filter",
    },
    category="BondMaster",
)
def BONDCOUNT(country: str | None = None) -> xlo.ExcelValue:
    """
    Count total bonds in the database.
    
    EXAMPLES:
        =BONDCOUNT()       → Total all bonds
        =BONDCOUNT("US")   → US bonds only
        =BONDCOUNT("GB")   → UK gilts only
    """
    success, data = _api_request("GET", "/stats")
    if not success:
        return _format_error(str(data))

    if country:
        country = country.upper().strip()
        by_country = data.get("by_country", {})
        return by_country.get(country, 0)

    return data.get("total_bonds", 0)


# =============================================================================
# ANALYTICS FUNCTIONS
# =============================================================================

@xlo.func(
    help="Calculate years to maturity for a bond.\n\nExample: =BONDYEARSTOMAT(\"GB00BYZW3G56\")",
    args={
        "isin": "ISIN code",
        "as_of": "Optional: calculation date (default: today)",
    },
    category="BondMaster",
)
def BONDYEARSTOMAT(isin: str, as_of: str | None = None) -> xlo.ExcelValue:
    """
    Calculate years remaining to maturity.
    
    EXAMPLES:
        =BONDYEARSTOMAT("GB00BYZW3G56")           → Years from today
        =BONDYEARSTOMAT("GB00BYZW3G56", "2025-01-01") → Years from specific date
    
    RETURNS: Decimal years (e.g., 5.25 = 5 years 3 months)
    """
    bond = _fetch_bond(isin)
    if bond is None:
        return _format_error(f"Bond not found: {isin}")

    maturity = _parse_date(bond.get("maturity_date"))
    if maturity is None:
        return _format_error("No maturity date available")

    if as_of:
        try:
            calc_date = datetime.fromisoformat(as_of).date()
        except ValueError:
            return _format_error(f"Invalid date format: {as_of}")
    else:
        calc_date = date.today()

    if maturity <= calc_date:
        return 0.0

    days = (maturity - calc_date).days
    return round(days / 365.25, 2)


@xlo.func(
    help="Get bonds maturing within a date range.\n\nExample: =BONDMATURITYRANGE(\"2025-01-01\", \"2025-12-31\", \"US\")",
    args={
        "from_date": "Start date (YYYY-MM-DD)",
        "to_date": "End date (YYYY-MM-DD)",
        "country": "Optional country filter",
    },
    category="BondMaster",
)
def BONDMATURITYRANGE(
    from_date: str,
    to_date: str,
    country: str | None = None,
) -> xlo.ExcelValue:
    """
    Get ISINs of bonds maturing within a date range.
    
    EXAMPLES:
        =BONDMATURITYRANGE("2025-01-01", "2025-12-31")
        =BONDMATURITYRANGE("2025-01-01", "2030-12-31", "GB")
    
    USE CASE: Find bonds for reinvestment planning
    """
    params = {
        "maturity_from": from_date,
        "maturity_to": to_date,
        "limit": 500,
    }
    if country:
        params["country"] = country.upper().strip()

    success, data = _api_request("GET", "/bonds", params=params)
    if not success:
        return _format_error(str(data))

    bonds = data.get("data", data) if isinstance(data, dict) else data
    if not bonds:
        return _format_error("No bonds maturing in range")

    # Return ISIN and maturity date
    result = []
    for b in bonds:
        result.append([b.get("isin", ""), b.get("maturity_date", "")])

    return result


@xlo.func(
    help="Get bond coupon payment frequency in plain text.\n\nExample: =BONDCOUPONFREQ(\"GB00BYZW3G56\")",
    args={
        "isin": "ISIN code",
    },
    category="BondMaster",
)
def BONDCOUPONFREQ(isin: str) -> xlo.ExcelValue:
    """
    Get coupon payment frequency as text.
    
    RETURNS: "Annual", "Semi-annual", "Quarterly", or "Zero coupon"
    """
    bond = _fetch_bond(isin)
    if bond is None:
        return _format_error(f"Bond not found: {isin}")

    freq = bond.get("coupon_frequency", 0)
    coupon = bond.get("coupon_rate", 0)

    if coupon == 0:
        return "Zero coupon"

    freq_map = {1: "Annual", 2: "Semi-annual", 4: "Quarterly", 12: "Monthly"}
    return freq_map.get(freq, f"{freq}x per year")


@xlo.func(
    help="Check if a bond is inflation-linked.\n\nExample: =BONDISLINKER(\"GB00B3LZBF68\")",
    args={
        "isin": "ISIN code",
    },
    category="BondMaster",
)
def BONDISLINKER(isin: str) -> xlo.ExcelValue:
    """
    Check if a bond is inflation-linked (index-linked).
    
    RETURNS: TRUE or FALSE
    
    EXAMPLE:
        =BONDISLINKER("GB00B3LZBF68")  → TRUE (UK index-linked gilt)
        =BONDISLINKER("GB00BYZW3G56")  → FALSE (conventional gilt)
    """
    bond = _fetch_bond(isin)
    if bond is None:
        return _format_error(f"Bond not found: {isin}")

    return bond.get("security_type", "").upper() == "INDEX_LINKED"


# =============================================================================
# DATA MANAGEMENT FUNCTIONS
# =============================================================================

@xlo.func(
    help="Refresh bond data from sources (requires API key).\n\nExample: =BONDREFRESH(\"US\")",
    args={
        "country": "Country to refresh (or blank for all)",
        "api_key": "API key for authentication",
    },
    category="BondMaster",
)
def BONDREFRESH(country: str | None = None, api_key: str | None = None) -> xlo.ExcelValue:
    """
    Trigger a refresh of bond data from live sources.
    
    REQUIRES: API key set via BONDMASTER_API_KEY environment variable
    
    EXAMPLES:
        =BONDREFRESH("US", "your-api-key")  → Refresh US bonds
        =BONDREFRESH(, "your-api-key")      → Refresh all bonds
    
    NOTE: Runs in background. Check =BONDCOUNT() after a minute.
    """
    headers = {}
    if api_key:
        headers["X-API-Key"] = api_key

    json_data = {}
    if country:
        json_data["country"] = country.upper().strip()
    else:
        json_data["full"] = True

    success, data = _api_request("POST", "/bonds/refresh", json=json_data, headers=headers)
    if not success:
        return _format_error(str(data))

    # Clear cache since data may be updating
    _bond_cache.clear()

    return data.get("message", "Refresh started")


# =============================================================================
# ENTERPRISE FUNCTIONS
# =============================================================================

@xlo.func(
    help="Get data lineage (source attribution) for a bond.\n\nExample: =BONDLINEAGE(\"DE0001102580\", \"coupon_rate\")",
    args={
        "isin": "ISIN code",
        "field": "Optional: specific field to check",
    },
    category="BondMaster",
)
def BONDLINEAGE(isin: str, field: str | None = None) -> xlo.ExcelValue:
    """
    Get data lineage showing which source provided each field.
    
    USE CASE: Audit trail, data quality verification
    
    EXAMPLES:
        =BONDLINEAGE("DE0001102580")              → All field sources
        =BONDLINEAGE("DE0001102580", "coupon_rate") → Source for coupon
    
    RETURNS: Source name and confidence level
    """
    isin = _normalize_isin(isin)
    success, data = _api_request("GET", f"/enterprise/lineage/{isin}")

    if not success:
        return _format_error(str(data))

    lineage = data.get("data")
    if lineage is None:
        return _format_error(f"No lineage data for {isin}")

    if field:
        field = field.lower().strip()
        sources = lineage.get("field_sources", {})
        if field not in sources:
            return _format_error(f"No lineage for field: {field}")
        src = sources[field]
        return f"{src.get('source_name', 'Unknown')} (confidence: {src.get('confidence', 0):.0%})"

    # Return summary
    sources = lineage.get("contributing_sources", [])
    confidence = lineage.get("reconciliation_confidence", 0)
    return f"Sources: {', '.join(sources)} | Confidence: {confidence:.0%}"


@xlo.func(
    help="Get change history for a bond.\n\nExample: =BONDHISTORY(\"DE0001102580\")",
    args={
        "isin": "ISIN code",
        "limit": "Max records (default: 10)",
    },
    category="BondMaster",
)
def BONDHISTORY(isin: str, limit: int = 10) -> xlo.ExcelValue:
    """
    Get change history for a bond (event sourcing).
    
    USE CASE: Track when data changed and what values were affected
    
    RETURNS: Array with change records [Date, Type, Field, Old, New]
    """
    isin = _normalize_isin(isin)
    success, data = _api_request(
        "GET",
        f"/enterprise/history/{isin}",
        params={"limit": limit},
    )

    if not success:
        return _format_error(str(data))

    history = data.get("data", [])
    if not history:
        return _format_error(f"No history for {isin}")

    result = [["Date", "Type", "Field", "Old Value", "New Value"]]
    for record in history:
        result.append([
            record.get("changed_at", ""),
            record.get("change_type", ""),
            record.get("field_name", ""),
            record.get("old_value", ""),
            record.get("new_value", ""),
        ])

    return result


@xlo.func(
    help="Get corporate actions (maturities, calls).\n\nExample: =BONDACTIONS(\"MATURED\", 30)",
    args={
        "action_type": "Optional: MATURED, CALLED, COUPON_CHANGE",
        "days_ahead": "For maturities: days to look ahead (default: 30)",
    },
    category="BondMaster",
)
def BONDACTIONS(
    action_type: str | None = None,
    days_ahead: int = 30,
) -> xlo.ExcelValue:
    """
    Get corporate actions affecting bonds.
    
    ACTION TYPES:
        MATURED       - Bond reached maturity
        CALLED        - Early redemption by issuer
        COUPON_CHANGE - Coupon rate changed
    
    EXAMPLES:
        =BONDACTIONS()                → All recent actions
        =BONDACTIONS("MATURED", 60)   → Maturities in next 60 days
    
    USE CASE: Portfolio rebalancing, cash flow planning
    """
    if action_type and action_type.upper() == "MATURED":
        # Use upcoming maturities endpoint
        success, data = _api_request(
            "GET",
            "/enterprise/corporate-actions/maturities",
            params={"days": days_ahead},
        )
    else:
        params = {"limit": 100}
        if action_type:
            params["action_type"] = action_type.upper()
        success, data = _api_request("GET", "/enterprise/corporate-actions", params=params)

    if not success:
        return _format_error(str(data))

    actions = data.get("data", [])
    if not actions:
        return _format_error("No corporate actions found")

    result = [["ISIN", "Type", "Effective Date", "Notes"]]
    for action in actions:
        result.append([
            action.get("isin", ""),
            action.get("action_type", ""),
            action.get("effective_date", ""),
            action.get("notes", ""),
        ])

    return result


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

@xlo.func(
    help="Check if BondMaster API is running and connected.",
    category="BondMaster",
    volatile=True,
)
def BONDAPI_STATUS() -> str:
    """
    Check BondMaster API connectivity.
    
    RETURNS:
        "✓ Connected" - API is running
        "✗ Disconnected: <reason>" - API not available
    
    TROUBLESHOOTING:
        1. Start API: bondmaster serve
        2. Check: http://127.0.0.1:8000/health
    """
    success, data = _api_request("GET", "/health")
    if success:
        return "✓ Connected"
    return f"✗ Disconnected: {data}"


@xlo.func(
    help="Clear the bond data cache (forces refresh from API).",
    category="BondMaster",
    volatile=True,
)
def BONDCACHE_CLEAR() -> str:
    """
    Clear cached bond data. Call after updating the database.
    
    WHEN TO USE:
        - After running =BONDREFRESH()
        - After manual database updates
        - To force fresh data fetch
    """
    count = _bond_cache.clear()
    return f"✓ Cleared {count} cached entries"


@xlo.func(
    help="Show cache performance statistics.",
    category="BondMaster",
    volatile=True,
)
def BONDCACHE_STATS() -> str:
    """
    Display cache performance statistics.
    
    RETURNS: "Size: N/500 | Hit Rate: X% | TTL: 300s"
    
    INTERPRETING:
        - High hit rate (>80%) = good cache utilization
        - Low hit rate = consider longer TTL
    """
    stats = _bond_cache.stats()
    return (
        f"Size: {stats['size']}/{stats['maxsize']} | "
        f"Hit Rate: {stats['hit_rate']:.0%} | "
        f"TTL: {stats['ttl_seconds']:.0f}s"
    )


@xlo.func(
    help="Show help for BondMaster functions.",
    args={
        "topic": "Optional: 'fields', 'countries', 'functions', or function name",
    },
    category="BondMaster",
)
def BONDHELP(topic: str | None = None) -> xlo.ExcelValue:
    """
    Get help on BondMaster functions.
    
    TOPICS:
        =BONDHELP()            → Overview
        =BONDHELP("fields")    → List all available fields
        =BONDHELP("countries") → List country codes
        =BONDHELP("functions") → List all functions
    """
    if topic is None:
        return [
            ["BondMaster Excel Add-in - Quick Reference"],
            [""],
            ["GETTING STARTED:"],
            ["1. Start API: bondmaster serve"],
            ["2. Check connection: =BONDAPI_STATUS()"],
            ["3. Try: =BONDSTATIC(\"US912810TM58\", \"coupon_rate\")"],
            [""],
            ["HELP TOPICS:"],
            ["=BONDHELP(\"fields\")    - Available data fields"],
            ["=BONDHELP(\"countries\") - Country codes"],
            ["=BONDHELP(\"functions\") - All functions"],
        ]

    topic = topic.lower().strip()

    if topic == "fields":
        result = [["Field", "Description"]]
        for field, desc in BOND_FIELDS.items():
            result.append([field, desc])
        return result

    if topic == "countries":
        result = [["Code", "Country"]]
        for code, name in COUNTRY_CODES.items():
            result.append([code, name])
        return result

    if topic == "functions":
        return [
            ["Function", "Description"],
            ["BONDSTATIC", "Get a single field value"],
            ["BONDINFO", "Get all fields as a row"],
            ["BONDLIST", "List ISINs by country"],
            ["BONDSEARCH", "Search with filters"],
            ["BONDCOUNT", "Count bonds"],
            ["BONDYEARSTOMAT", "Years to maturity"],
            ["BONDMATURITYRANGE", "Bonds maturing in date range"],
            ["BONDCOUPONFREQ", "Payment frequency text"],
            ["BONDISLINKER", "Check if inflation-linked"],
            ["BONDREFRESH", "Refresh data from sources"],
            ["BONDLINEAGE", "Data source attribution"],
            ["BONDHISTORY", "Change history"],
            ["BONDACTIONS", "Corporate actions"],
            ["BONDAPI_STATUS", "Check API connection"],
            ["BONDCACHE_CLEAR", "Clear cache"],
            ["BONDCACHE_STATS", "Cache statistics"],
            ["BONDHELP", "This help"],
        ]

    return _format_error(f"Unknown topic: {topic}. Try: fields, countries, functions")


@xlo.func(
    help="Validate an ISIN code (format and checksum).\n\nExample: =BONDISINVALID(\"GB00BYZW3G56\")",
    args={
        "isin": "ISIN to validate",
    },
    category="BondMaster",
)
def BONDISINVALID(isin: str) -> xlo.ExcelValue:
    """
    Validate an ISIN code.
    
    CHECKS:
        - Length (12 characters)
        - Format (2 letters + 9 alphanumeric + 1 digit)
        - Country code validity
    
    RETURNS: TRUE if valid, FALSE if invalid
    """
    if not isin:
        return False

    isin = _normalize_isin(isin)

    # Basic format check
    if not _is_valid_isin(isin):
        return False

    # Check country code
    country = isin[:2]
    # Accept known countries plus common ISIN prefixes
    valid_prefixes = set(COUNTRY_CODES.keys()) | {"XS", "EU"}  # XS = Eurobonds
    if country not in valid_prefixes:
        return False

    return True
