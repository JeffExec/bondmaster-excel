"""
BondMaster Excel UDFs (User Defined Functions).

These functions are exposed to Excel via xlOil.
"""

import xloil as xlo
from typing import Optional
import httpx
from functools import lru_cache
from datetime import datetime

# Configuration
API_BASE_URL = "http://127.0.0.1:8000"
REQUEST_TIMEOUT = 10.0

# Cache for API responses (cleared on recalc)
_bond_cache: dict[str, dict] = {}


def _get_client() -> httpx.Client:
    """Get or create HTTP client."""
    return httpx.Client(base_url=API_BASE_URL, timeout=REQUEST_TIMEOUT)


def _fetch_bond(isin: str) -> dict | None:
    """Fetch bond data from API with caching."""
    isin = isin.upper().strip()
    
    if isin in _bond_cache:
        return _bond_cache[isin]
    
    try:
        with _get_client() as client:
            response = client.get(f"/bonds/{isin}")
            if response.status_code == 200:
                data = response.json()
                _bond_cache[isin] = data
                return data
            elif response.status_code == 404:
                return None
            else:
                return None
    except httpx.RequestError:
        return None


def _clear_cache():
    """Clear the bond cache."""
    global _bond_cache
    _bond_cache = {}


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
        coupon_rate, coupon_frequency, day_count_convention,
        maturity_date, issue_date, first_coupon_date, original_tenor,
        outstanding_amount
    
    Examples:
        =BONDSTATIC("GB00BYZW3G56", "coupon_rate")     → 0.015
        =BONDSTATIC("GB00BYZW3G56", "maturity_date")  → 2026-07-22
        =BONDSTATIC("US912810TM58", "issuer")         → United States Treasury
    """
    if not isin or not field:
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
def BONDLIST(country: str, security_type: Optional[str] = None) -> xlo.ExcelValue:
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
        with _get_client() as client:
            params = {"country": country, "limit": 1000}
            if security_type:
                params["security_type"] = security_type.upper()
            
            response = client.get("/bonds", params=params)
            if response.status_code != 200:
                return xlo.CellError.NA
            
            data = response.json()
            
            # Handle both list and envelope responses
            bonds = data if isinstance(data, list) else data.get("data", [])
            
            if not bonds:
                return xlo.CellError.NA
            
            # Return as vertical array
            return [[b.get("isin", "")] for b in bonds]
    
    except httpx.RequestError:
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
    field2: Optional[str] = None,
    value2: Optional[str] = None,
    field3: Optional[str] = None,
    value3: Optional[str] = None,
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
        with _get_client() as client:
            response = client.get("/bonds", params=params)
            if response.status_code != 200:
                return xlo.CellError.NA
            
            data = response.json()
            bonds = data if isinstance(data, list) else data.get("data", [])
            
            if not bonds:
                return xlo.CellError.NA
            
            return [[b.get("isin", "")] for b in bonds]
    
    except httpx.RequestError:
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
def BONDCOUNT(country: Optional[str] = None) -> xlo.ExcelValue:
    """
    Count total bonds in the database.
    
    Examples:
        =BONDCOUNT()       → Total bonds
        =BONDCOUNT("US")   → US bonds only
    """
    try:
        with _get_client() as client:
            response = client.get("/stats")
            if response.status_code != 200:
                return xlo.CellError.NA
            
            data = response.json()
            
            if country:
                country = country.upper()
                by_country = data.get("by_country", {})
                return by_country.get(country, 0)
            
            return data.get("total_bonds", 0)
    
    except httpx.RequestError:
        return xlo.CellError.NA


# =============================================================================
# BONDAPI_STATUS - Check API connectivity
# =============================================================================

@xlo.func(
    help="Check if BondMaster API is running",
    category="BondMaster",
)
def BONDAPI_STATUS() -> str:
    """
    Check BondMaster API connectivity.
    
    Returns "Connected" or "Disconnected".
    """
    try:
        with _get_client() as client:
            response = client.get("/health")
            if response.status_code == 200:
                return "Connected"
            return f"Error: {response.status_code}"
    except httpx.RequestError as e:
        return f"Disconnected: {e}"


# =============================================================================
# BONDCACHE_CLEAR - Clear the cache
# =============================================================================

@xlo.func(
    help="Clear the bond data cache (forces refresh from API)",
    category="BondMaster",
)
def BONDCACHE_CLEAR() -> str:
    """
    Clear cached bond data. Call this after updating the database.
    """
    _clear_cache()
    return f"Cache cleared at {datetime.now().strftime('%H:%M:%S')}"
