"""
BondMaster Excel UDFs (User Defined Functions).

These functions are exposed to Excel via xlOil.

Complexity Analysis:
    - BONDSTATIC: O(1) cache hit, O(1) API call on miss (single bond lookup)
    - BONDINFO: O(1) same as BONDSTATIC
    - BONDLIST: O(n) where n = bonds in country (streams from API)
    - BONDSEARCH: O(n) where n = matching bonds
    - BONDCOUNT: O(1) via pre-computed stats endpoint

Cache Strategy:
    TTL-based LRU cache (5 min default). Bond reference data is semi-static
    (changes infrequently), so short TTL balances freshness vs. performance.
    Cache auto-evicts stale entries on access. Manual clear via BONDCACHE_CLEAR().
"""

import os
import re
import threading
import time
from typing import NamedTuple

import httpx
import xloil as xlo

# Configuration (can be overridden via environment variable)
API_BASE_URL = os.environ.get("BONDMASTER_API_URL", "http://127.0.0.1:8000")
REQUEST_TIMEOUT = 10.0
MAX_RETRIES = 2
CACHE_TTL_SECONDS = float(os.environ.get("BONDMASTER_CACHE_TTL", "300"))  # 5 min default
CACHE_MAX_SIZE = 500

# ISIN validation pattern: 2 letters + 9 alphanumeric + 1 check digit
ISIN_PATTERN = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}[0-9]$")

# Module-level HTTP client (singleton)
_client: httpx.Client | None = None


def _get_client() -> httpx.Client:
    """Get or create HTTP client singleton."""
    global _client
    if _client is None:
        _client = httpx.Client(base_url=API_BASE_URL, timeout=REQUEST_TIMEOUT)
    return _client


def _close_client() -> None:
    """Close the HTTP client."""
    global _client
    if _client is not None:
        _client.close()
        _client = None


def _is_valid_isin(isin: str) -> bool:
    """Validate ISIN format. O(1) regex match."""
    return bool(ISIN_PATTERN.match(isin))


# =============================================================================
# TTL-aware LRU Cache
# =============================================================================

class _CacheEntry(NamedTuple):
    """Cache entry with expiration timestamp."""
    data: tuple  # Bond data as tuple of tuples (hashable)
    expires_at: float  # Unix timestamp


class _TTLCache:
    """
    Thread-safe TTL-aware LRU cache.

    Complexity:
        - get: O(1) average (dict lookup)
        - set: O(1) average, O(n) worst case during eviction
        - Eviction: LRU when full, plus TTL expiry on access
    """

    def __init__(self, maxsize: int = 500, ttl_seconds: float = 300.0):
        self._cache: dict[str, _CacheEntry] = {}
        self._maxsize = maxsize
        self._ttl = ttl_seconds
        self._lock = threading.Lock()
        self._hits = 0
        self._misses = 0

    def get(self, key: str) -> tuple | None:
        """Get value if exists and not expired. Returns None on miss/expiry."""
        with self._lock:
            entry = self._cache.get(key)
            if entry is None:
                self._misses += 1
                return None

            # Check TTL
            if time.time() > entry.expires_at:
                del self._cache[key]
                self._misses += 1
                return None

            # Move to end (most recently used) - Python 3.7+ dicts maintain order
            self._cache[key] = self._cache.pop(key)
            self._hits += 1
            return entry.data

    def set(self, key: str, value: tuple) -> None:
        """Store value with TTL. Evicts LRU if at capacity."""
        with self._lock:
            # Remove existing to update position
            if key in self._cache:
                del self._cache[key]
            # Evict oldest if at capacity
            elif len(self._cache) >= self._maxsize:
                oldest_key = next(iter(self._cache))
                del self._cache[oldest_key]

            self._cache[key] = _CacheEntry(
                data=value,
                expires_at=time.time() + self._ttl
            )

    def clear(self) -> None:
        """Clear all cached entries."""
        with self._lock:
            self._cache.clear()
            self._hits = 0
            self._misses = 0

    def stats(self) -> dict:
        """Return cache statistics."""
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


# Global cache instance
_bond_cache = _TTLCache(maxsize=CACHE_MAX_SIZE, ttl_seconds=CACHE_TTL_SECONDS)


def _fetch_bond_from_api(isin: str) -> tuple | None:
    """
    Fetch bond data from API with retry logic.

    Returns tuple of (key, value) pairs for hashability, or None.
    Complexity: O(1) - single HTTP request + JSON parse.
    """
    for attempt in range(MAX_RETRIES + 1):
        try:
            client = _get_client()
            response = client.get(f"/bonds/{isin}")

            if response.status_code == 200:
                data = response.json()
                return tuple(sorted(data.items()))
            elif response.status_code == 404:
                xlo.log(f"Bond not found: {isin}", level="DEBUG")
                return None
            else:
                xlo.log(
                    f"API error for {isin}: HTTP {response.status_code}",
                    level="WARNING"
                )
                return None

        except httpx.TimeoutException:
            xlo.log(f"Timeout fetching bond {isin} (attempt {attempt + 1})", level="WARNING")
            if attempt < MAX_RETRIES:
                time.sleep(0.1 * (2 ** attempt))  # Exponential backoff
                continue
            return None

        except httpx.RequestError as e:
            xlo.log(f"Request error for {isin}: {e}", level="ERROR")
            if attempt < MAX_RETRIES:
                time.sleep(0.1 * (2 ** attempt))
                continue
            return None

    return None


def _fetch_bond(isin: str) -> dict | None:
    """
    Fetch bond with TTL caching.

    Complexity: O(1) on cache hit, O(1) + network on cache miss.
    """
    isin = isin.upper().strip()

    if not _is_valid_isin(isin):
        return None

    # Check cache first
    cached = _bond_cache.get(isin)
    if cached is not None:
        return dict(cached)

    # Cache miss - fetch from API
    result = _fetch_bond_from_api(isin)
    if result is None:
        return None

    # Store in cache
    _bond_cache.set(isin, result)
    return dict(result)


def _clear_cache() -> None:
    """Clear the bond cache."""
    _bond_cache.clear()


# =============================================================================
# BONDSTATIC - Get a single field for a bond
# =============================================================================

@xlo.func(
    help="Get static reference data for a bond by ISIN",
    args={
        "isin": "The ISIN code of the bond (e.g., 'GB00BYZW3G56')",
        "field": "The field to retrieve (e.g., 'coupon_rate', 'maturity_date', 'issuer')",
    },
    category="BondMaster",
)
def BONDSTATIC(isin: str, field: str) -> xlo.ExcelValue:
    """
    Get a specific field from bond reference data.
    
    Available fields:
        isin, cusip, name, country, issuer, security_type, currency,
        coupon_rate (returns % e.g. 1.5 for 1.5%), coupon_frequency,
        day_count_convention, maturity_date, issue_date, first_coupon_date,
        original_tenor, outstanding_amount
    
    Examples:
        =BONDSTATIC("GB00BYZW3G56", "coupon_rate")     → 1.5 (%)
        =BONDSTATIC("GB00BYZW3G56", "maturity_date")   → 2026-07-22
        =BONDSTATIC("US912810TM58", "issuer")          → United States Treasury
    """
    if not isin or not field:
        return xlo.CellError.Value

    isin = isin.upper().strip()
    if not _is_valid_isin(isin):
        return xlo.CellError.Value

    bond = _fetch_bond(isin)
    if bond is None:
        return xlo.CellError.NA  # Bond not found

    field = field.lower().strip()

    # Handle special field aliases
    field_map = {
        "coupon": "coupon_rate",
        "maturity": "maturity_date",
        "issue": "issue_date",
        "type": "security_type",
        "freq": "coupon_frequency",
        "frequency": "coupon_frequency",
    }
    field = field_map.get(field, field)

    if field not in bond:
        return xlo.CellError.Name  # Unknown field

    value = bond.get(field)

    if value is None:
        return ""

    # Convert coupon rate to percentage for display
    if field == "coupon_rate" and isinstance(value, (int, float)):
        return value * 100  # Return as percentage (e.g., 1.5 for 1.5%)

    return value


# =============================================================================
# BONDINFO - Get all fields for a bond as a row
# =============================================================================

@xlo.func(
    help="Get all reference data for a bond as a row (spills)",
    args={
        "isin": "The ISIN code of the bond",
        "with_headers": "Include header row (default: False)",
    },
    category="BondMaster",
)
def BONDINFO(isin: str, with_headers: bool = False) -> xlo.ExcelValue:
    """
    Get all bond data as an array that spills across cells.
    
    Examples:
        =BONDINFO("GB00BYZW3G56")           → Single row of data
        =BONDINFO("GB00BYZW3G56", TRUE)     → Header row + data row
    """
    if not isin:
        return xlo.CellError.Value

    isin = isin.upper().strip()
    if not _is_valid_isin(isin):
        return xlo.CellError.Value

    bond = _fetch_bond(isin)
    if bond is None:
        return xlo.CellError.NA

    # Define column order
    columns = [
        "isin", "name", "country", "issuer", "security_type", "currency",
        "coupon_rate", "coupon_frequency", "maturity_date", "issue_date",
        "outstanding_amount"
    ]

    values = []
    for col in columns:
        val = bond.get(col, "")
        if col == "coupon_rate" and isinstance(val, (int, float)):
            val = val * 100  # Percentage
        values.append(val if val is not None else "")

    if with_headers:
        # Format headers nicely
        headers = [col.replace("_", " ").title() for col in columns]
        return [headers, values]

    return [values]


# =============================================================================
# BONDLIST - Get list of ISINs for a country/filter
# =============================================================================

@xlo.func(
    help="Get list of ISINs for a country, optionally filtered by security type",
    args={
        "country": "Country code (e.g., 'US', 'GB', 'DE')",
        "security_type": "Optional: 'NOMINAL' or 'INDEX_LINKED'",
    },
    category="BondMaster",
)
def BONDLIST(country: str, security_type: str | None = None) -> xlo.ExcelValue:
    """
    Get all ISINs for a country as a vertical array (spills down).
    
    Examples:
        =BONDLIST("GB")                    → All UK gilt ISINs
        =BONDLIST("US", "INDEX_LINKED")    → US TIPS only
    """
    if not country:
        return xlo.CellError.Value

    country = country.upper().strip()

    try:
        client = _get_client()
        params = {"country": country, "limit": 1000}
        if security_type:
            params["security_type"] = security_type.upper()

        response = client.get("/bonds", params=params)
        if response.status_code != 200:
            xlo.log(f"BONDLIST API error: HTTP {response.status_code}", level="WARNING")
            return xlo.CellError.NA

        data = response.json()

        # Handle both list and envelope responses
        bonds = data if isinstance(data, list) else data.get("data", [])

        if not bonds:
            return xlo.CellError.NA

        # Return as vertical array
        return [[b.get("isin", "")] for b in bonds]

    except httpx.TimeoutException:
        xlo.log(f"Timeout in BONDLIST for {country}", level="WARNING")
        return xlo.CellError.NA
    except httpx.RequestError as e:
        xlo.log(f"Request error in BONDLIST: {e}", level="ERROR")
        return xlo.CellError.NA


# =============================================================================
# BONDSEARCH - Search bonds with filters
# =============================================================================

@xlo.func(
    help="Search bonds with field/value filter pairs",
    args={
        "field1": "First filter field (e.g., 'country')",
        "value1": "First filter value (e.g., 'US')",
        "field2": "Optional second filter field",
        "value2": "Optional second filter value",
        "field3": "Optional third filter field",
        "value3": "Optional third filter value",
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
    Search for bonds matching filter criteria.
    Returns ISINs as a vertical array.
    
    Supported filter fields:
        country, security_type, currency, maturity_from, maturity_to
    
    Examples:
        =BONDSEARCH("country", "DE")
        =BONDSEARCH("country", "US", "security_type", "INDEX_LINKED")
    """
    params = {"limit": 1000}

    # Build filter params
    if field1 and value1:
        params[field1.lower()] = value1
    if field2 and value2:
        params[field2.lower()] = value2
    if field3 and value3:
        params[field3.lower()] = value3

    if len(params) == 1:  # Only limit, no filters
        return xlo.CellError.Value

    try:
        client = _get_client()
        response = client.get("/bonds", params=params)
        if response.status_code != 200:
            xlo.log(f"BONDSEARCH API error: HTTP {response.status_code}", level="WARNING")
            return xlo.CellError.NA

        data = response.json()
        bonds = data if isinstance(data, list) else data.get("data", [])

        if not bonds:
            return xlo.CellError.NA

        return [[b.get("isin", "")] for b in bonds]

    except httpx.TimeoutException:
        xlo.log("Timeout in BONDSEARCH", level="WARNING")
        return xlo.CellError.NA
    except httpx.RequestError as e:
        xlo.log(f"Request error in BONDSEARCH: {e}", level="ERROR")
        return xlo.CellError.NA


# =============================================================================
# BONDCOUNT - Count bonds
# =============================================================================

@xlo.func(
    help="Count bonds, optionally filtered by country",
    args={
        "country": "Optional country code to filter",
    },
    category="BondMaster",
)
def BONDCOUNT(country: str | None = None) -> xlo.ExcelValue:
    """
    Count total bonds in the database.
    
    Examples:
        =BONDCOUNT()       → Total bonds
        =BONDCOUNT("US")   → US bonds only
    """
    try:
        client = _get_client()
        response = client.get("/stats")
        if response.status_code != 200:
            xlo.log(f"BONDCOUNT API error: HTTP {response.status_code}", level="WARNING")
            return xlo.CellError.NA

        data = response.json()

        if country:
            country = country.upper()
            by_country = data.get("by_country", {})
            return by_country.get(country, 0)

        return data.get("total_bonds", 0)

    except httpx.TimeoutException:
        xlo.log("Timeout in BONDCOUNT", level="WARNING")
        return xlo.CellError.NA
    except httpx.RequestError as e:
        xlo.log(f"Request error in BONDCOUNT: {e}", level="ERROR")
        return xlo.CellError.NA


# =============================================================================
# BONDAPI_STATUS - Check API connectivity
# =============================================================================

@xlo.func(
    help="Check if BondMaster API is running",
    category="BondMaster",
    volatile=True,  # Always recalculate
)
def BONDAPI_STATUS() -> str:
    """
    Check BondMaster API connectivity.
    
    Returns "Connected" or "Disconnected".
    """
    try:
        client = _get_client()
        response = client.get("/health")
        if response.status_code == 200:
            return "Connected"
        return f"Error: {response.status_code}"
    except httpx.TimeoutException:
        return "Disconnected: Timeout"
    except httpx.RequestError as e:
        return f"Disconnected: {type(e).__name__}"


# =============================================================================
# BONDCACHE_CLEAR - Clear the cache
# =============================================================================

@xlo.func(
    help="Clear the bond data cache (forces refresh from API)",
    category="BondMaster",
    volatile=True,  # Always execute
)
def BONDCACHE_CLEAR() -> str:
    """
    Clear cached bond data. Call this after updating the database.
    """
    stats = _bond_cache.stats()
    _clear_cache()
    return f"Cleared {stats['size']} entries (was {stats['hit_rate']:.0%} hit rate)"


# =============================================================================
# BONDCACHE_STATS - Show cache statistics
# =============================================================================

@xlo.func(
    help="Show cache statistics (size, hit rate, TTL)",
    category="BondMaster",
    volatile=True,
)
def BONDCACHE_STATS() -> str:
    """
    Display cache performance statistics.

    Returns: "Size: N/500 | Hits: X% | TTL: 300s"
    """
    stats = _bond_cache.stats()
    return (
        f"Size: {stats['size']}/{stats['maxsize']} | "
        f"Hits: {stats['hit_rate']:.0%} | "
        f"TTL: {stats['ttl_seconds']:.0f}s"
    )
