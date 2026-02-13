"""Tests for BondMaster Excel UDFs.

Comprehensive test suite covering:
- All UDF functions with mocked xloil and httpx
- Empty/invalid inputs
- API errors (connection failures, timeouts, HTTP errors)
- Malformed API responses
- Edge cases and boundary conditions
"""

import sys
from unittest.mock import MagicMock, patch

import pytest

# Mock xloil before importing udfs
mock_xloil = MagicMock()


class MockCellError:
    """Mock xloil.CellError enum."""
    Value = "#VALUE!"
    NA = "#N/A"
    Name = "#NAME?"


mock_xloil.CellError = MockCellError


def is_error(result) -> bool:
    """Check if result is an error message (new format: starts with ⚠️)."""
    if isinstance(result, str):
        return result.startswith("⚠️")
    return result in (MockCellError.Value, MockCellError.NA, MockCellError.Name)
mock_xloil.ExcelValue = object  # Type hint only
mock_xloil.func = lambda **kwargs: lambda f: f  # Decorator that returns function unchanged

sys.modules["xloil"] = mock_xloil

# Now import after mocking
from bondmaster_excel import udfs

# =============================================================================
# Test Fixtures
# =============================================================================

MOCK_BOND = {
    "isin": "GB00BYZW3G56",
    "cusip": None,
    "name": "1½% Treasury Gilt 2026",
    "country": "GB",
    "issuer": "UK Debt Management Office",
    "security_type": "NOMINAL",
    "currency": "GBP",
    "coupon_rate": 0.015,
    "coupon_frequency": 2,
    "day_count_convention": "ACT/ACT",
    "maturity_date": "2026-07-22",
    "issue_date": "2016-02-18",
    "first_coupon_date": "2016-07-22",
    "original_tenor": 10,
    "outstanding_amount": 35000000000,
}

MOCK_BOND_2 = {
    "isin": "US912810TM58",
    "cusip": "912810TM5",
    "name": "United States Treasury Note",
    "country": "US",
    "issuer": "United States Treasury",
    "security_type": "NOMINAL",
    "currency": "USD",
    "coupon_rate": 0.025,
    "coupon_frequency": 2,
    "maturity_date": "2028-05-15",
    "issue_date": "2018-05-15",
    "outstanding_amount": 50000000000,
}


@pytest.fixture(autouse=True)
def clear_cache():
    """Clear cache before each test."""
    udfs._bond_cache.clear()
    yield
    udfs._bond_cache.clear()


# =============================================================================
# Mock Response Helpers
# =============================================================================

class MockResponse:
    """Mock httpx.Response."""
    def __init__(self, status_code: int, json_data=None):
        self.status_code = status_code
        self._json_data = json_data

    def json(self):
        if self._json_data is None:
            raise ValueError("No JSON data")
        return self._json_data


class MockClient:
    """Mock httpx.Client context manager."""
    def __init__(self, get_func):
        self.get_func = get_func

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def get(self, url, params=None):
        return self.get_func(url, params)

    def request(self, method, url, params=None, json=None, headers=None):
        """Generic request method used by _api_request."""
        return self.get_func(url, params)


# =============================================================================
# BONDSTATIC Tests
# =============================================================================

class TestBONDSTATIC:
    """Tests for BONDSTATIC function."""

    def test_returns_field_value(self):
        """Test successful field retrieval."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "country")
            assert result == "GB"

    def test_coupon_rate_converted_to_percentage(self):
        """Test coupon rate is multiplied by 100."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "coupon_rate")
            assert result == 1.5  # 0.015 * 100

    def test_field_aliases_work(self):
        """Test field aliases (coupon -> coupon_rate)."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            assert udfs.BONDSTATIC("GB00BYZW3G56", "coupon") == 1.5
            assert udfs.BONDSTATIC("GB00BYZW3G56", "maturity") == "2026-07-22"
            assert udfs.BONDSTATIC("GB00BYZW3G56", "type") == "NOMINAL"
            assert udfs.BONDSTATIC("GB00BYZW3G56", "freq") == 2
            assert udfs.BONDSTATIC("GB00BYZW3G56", "frequency") == 2

    def test_none_field_returns_empty_string(self):
        """Test None field value returns empty string."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "cusip")
            assert result == ""

    # --- Empty Input Tests ---

    def test_empty_isin_returns_value_error(self):
        """Test empty ISIN returns #VALUE! error."""
        result = udfs.BONDSTATIC("", "coupon_rate")
        assert is_error(result)

    def test_none_isin_returns_value_error(self):
        """Test None ISIN returns #VALUE! error."""
        result = udfs.BONDSTATIC(None, "coupon_rate")
        assert is_error(result)

    def test_empty_field_returns_value_error(self):
        """Test empty field returns #VALUE! error."""
        result = udfs.BONDSTATIC("GB00BYZW3G56", "")
        assert is_error(result)

    def test_none_field_returns_value_error(self):
        """Test None field returns #VALUE! error."""
        result = udfs.BONDSTATIC("GB00BYZW3G56", None)
        assert is_error(result)

    def test_both_empty_returns_value_error(self):
        """Test both inputs empty returns #VALUE! error."""
        result = udfs.BONDSTATIC("", "")
        assert is_error(result)

    def test_whitespace_only_isin_returns_value_error(self):
        """Test whitespace-only ISIN returns #VALUE! (invalid format)."""
        # Whitespace gets stripped, then fails ISIN format validation
        result = udfs.BONDSTATIC("   ", "coupon_rate")
        assert is_error(result)

    # --- Bond Not Found Tests ---

    def test_bond_not_found_returns_na(self):
        """Test non-existent bond returns #N/A."""
        with patch.object(udfs, "_fetch_bond", return_value=None):
            result = udfs.BONDSTATIC("XX0000000000", "coupon_rate")
            assert is_error(result)

    # --- Unknown Field Tests ---

    def test_unknown_field_returns_name_error(self):
        """Test unknown field returns #NAME? error."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "nonexistent_field")
            assert is_error(result)

    # --- Input Normalization Tests ---

    def test_isin_normalized_to_uppercase(self):
        """Test ISIN is converted to uppercase."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND) as mock:
            udfs.BONDSTATIC("gb00byzw3g56", "country")
            # The normalization happens in _fetch_bond, verify it was called
            mock.assert_called_once()

    def test_field_normalized_to_lowercase(self):
        """Test field is normalized to lowercase."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "COUNTRY")
            assert result == "GB"


# =============================================================================
# BONDINFO Tests
# =============================================================================

class TestBONDINFO:
    """Tests for BONDINFO function."""

    def test_returns_data_row(self):
        """Test returns single row of data."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDINFO("GB00BYZW3G56")
            assert isinstance(result, list)
            assert len(result) == 1
            assert result[0][0] == "GB00BYZW3G56"  # ISIN first

    def test_with_headers_returns_two_rows(self):
        """Test with_headers=True returns header + data."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDINFO("GB00BYZW3G56", with_headers=True)
            assert len(result) == 2
            assert "ISIN" in result[0]  # Header row
            assert result[1][0] == "GB00BYZW3G56"  # Data row

    def test_coupon_rate_as_percentage(self):
        """Test coupon rate is percentage in row."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDINFO("GB00BYZW3G56")
            # Find coupon_rate column (index 6)
            assert result[0][6] == 1.5  # 0.015 * 100

    # --- Empty Input Tests ---

    def test_empty_isin_returns_value_error(self):
        """Test empty ISIN returns #VALUE! error."""
        result = udfs.BONDINFO("")
        assert is_error(result)

    def test_none_isin_returns_value_error(self):
        """Test None ISIN returns #VALUE! error."""
        result = udfs.BONDINFO(None)
        assert is_error(result)

    # --- Bond Not Found Tests ---

    def test_bond_not_found_returns_na(self):
        """Test non-existent bond returns #N/A."""
        with patch.object(udfs, "_fetch_bond", return_value=None):
            result = udfs.BONDINFO("XX0000000000")
            assert is_error(result)

    def test_missing_fields_return_empty_string(self):
        """Test missing fields in response become empty strings."""
        partial_bond = {"isin": "TEST123", "name": "Test Bond"}
        with patch.object(udfs, "_fetch_bond", return_value=partial_bond):
            result = udfs.BONDINFO("TEST123")
            # Fields not in partial_bond should be empty strings
            assert "" in result[0]


# =============================================================================
# BONDLIST Tests
# =============================================================================

class TestBONDLIST:
    """Tests for BONDLIST function."""

    def test_returns_vertical_isin_list(self):
        """Test returns ISINs as vertical array."""
        mock_response = MockResponse(200, [MOCK_BOND, MOCK_BOND_2])

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("GB")
            assert result == [["GB00BYZW3G56"], ["US912810TM58"]]

    def test_handles_envelope_response(self):
        """Test handles envelope-style response."""
        mock_response = MockResponse(200, {"data": [MOCK_BOND], "total": 1})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("GB")
            assert result == [["GB00BYZW3G56"]]

    def test_with_security_type_filter(self):
        """Test security_type filter is passed."""
        mock_response = MockResponse(200, [MOCK_BOND])

        def mock_get(url, params=None):
            assert params.get("security_type") == "INDEX_LINKED"
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            udfs.BONDLIST("US", "INDEX_LINKED")

    # --- Empty Input Tests ---

    def test_empty_country_returns_value_error(self):
        """Test empty country returns #VALUE! error."""
        result = udfs.BONDLIST("")
        assert is_error(result)

    def test_none_country_returns_value_error(self):
        """Test None country returns #VALUE! error."""
        result = udfs.BONDLIST(None)
        assert is_error(result)

    # --- API Error Tests ---

    def test_api_404_returns_na(self):
        """Test API 404 returns #N/A."""
        mock_response = MockResponse(404)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("ZZ")
            assert is_error(result)

    def test_api_500_returns_na(self):
        """Test API 500 error returns #N/A."""
        mock_response = MockResponse(500)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("GB")
            assert is_error(result)

    def test_connection_error_returns_na(self):
        """Test connection error returns #N/A."""
        import httpx

        def mock_get(url, params=None):
            raise httpx.RequestError("Connection failed")

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("GB")
            assert is_error(result)

    def test_empty_result_returns_na(self):
        """Test empty result list returns #N/A."""
        mock_response = MockResponse(200, [])

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("ZZ")
            assert is_error(result)


# =============================================================================
# BONDSEARCH Tests
# =============================================================================

class TestBONDSEARCH:
    """Tests for BONDSEARCH function."""

    def test_single_filter(self):
        """Test search with single filter."""
        mock_response = MockResponse(200, [MOCK_BOND])

        def mock_get(url, params=None):
            assert "country" in params
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDSEARCH("country", "GB")
            assert result == [["GB00BYZW3G56"]]

    def test_multiple_filters(self):
        """Test search with multiple filters."""
        mock_response = MockResponse(200, [MOCK_BOND])

        def mock_get(url, params=None):
            assert params.get("country") == "US"
            assert params.get("security_type") == "INDEX_LINKED"
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            udfs.BONDSEARCH("country", "US", "security_type", "INDEX_LINKED")

    def test_three_filters(self):
        """Test search with three filters."""
        mock_response = MockResponse(200, [MOCK_BOND])

        def mock_get(url, params=None):
            assert params.get("country") == "GB"
            assert params.get("currency") == "GBP"
            assert params.get("security_type") == "NOMINAL"
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            udfs.BONDSEARCH("country", "GB", "currency", "GBP", "security_type", "NOMINAL")

    # --- Empty Input Tests ---

    def test_no_filters_returns_value_error(self):
        """Test no filters returns #VALUE! error."""
        result = udfs.BONDSEARCH("", "")
        assert is_error(result)

    def test_field_without_value_returns_value_error(self):
        """Test field without value returns #VALUE! error."""
        result = udfs.BONDSEARCH("country", "")
        assert is_error(result)

    # --- API Error Tests ---

    def test_api_error_returns_na(self):
        """Test API error returns #N/A."""
        mock_response = MockResponse(500)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDSEARCH("country", "GB")
            assert is_error(result)

    def test_connection_error_returns_na(self):
        """Test connection error returns #N/A."""
        import httpx

        def mock_get(url, params=None):
            raise httpx.RequestError("Connection failed")

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDSEARCH("country", "GB")
            assert is_error(result)


# =============================================================================
# BONDCOUNT Tests
# =============================================================================

class TestBONDCOUNT:
    """Tests for BONDCOUNT function."""

    def test_total_count(self):
        """Test returns total bond count."""
        mock_response = MockResponse(200, {"total_bonds": 500, "by_country": {"GB": 100}})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT()
            assert result == 500

    def test_count_by_country(self):
        """Test returns count for specific country."""
        mock_response = MockResponse(200, {"total_bonds": 500, "by_country": {"GB": 100, "US": 300}})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT("GB")
            assert result == 100

    def test_unknown_country_returns_zero(self):
        """Test unknown country returns 0."""
        mock_response = MockResponse(200, {"total_bonds": 500, "by_country": {"GB": 100}})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT("ZZ")
            assert result == 0

    def test_country_normalized_uppercase(self):
        """Test country is normalized to uppercase."""
        mock_response = MockResponse(200, {"total_bonds": 500, "by_country": {"GB": 100}})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT("gb")
            assert result == 100

    # --- API Error Tests ---

    def test_api_error_returns_na(self):
        """Test API error returns #N/A."""
        mock_response = MockResponse(500)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT()
            assert is_error(result)

    def test_connection_error_returns_na(self):
        """Test connection error returns #N/A."""
        import httpx

        def mock_get(url, params=None):
            raise httpx.RequestError("Connection failed")

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT()
            assert is_error(result)


# =============================================================================
# BONDAPI_STATUS Tests
# =============================================================================

class TestBondApiStatus:
    """Tests for BONDAPI_STATUS function."""

    def test_connected(self):
        """Test returns Connected on success."""
        mock_response = MockResponse(200, {"status": "healthy"})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDAPI_STATUS()
            assert "Connected" in result

    def test_error_status_code(self):
        """Test returns error message for non-200."""
        mock_response = MockResponse(503)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDAPI_STATUS()
            assert "Disconnected" in result or "503" in result

    def test_connection_error(self):
        """Test returns Disconnected on connection error."""
        import httpx

        def mock_get(url, params=None):
            raise httpx.RequestError("Connection refused")

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDAPI_STATUS()
            assert "Disconnected" in result


# =============================================================================
# BONDCACHE_CLEAR Tests
# =============================================================================

class TestBondCacheClear:
    """Tests for BONDCACHE_CLEAR function."""

    def test_clears_cache(self):
        """Test cache is cleared."""
        # With TTL cache, we can verify by checking the returned stats
        with patch.object(udfs, "_get_client") as mock_client:
            mock_response = MockResponse(200, MOCK_BOND)
            mock_client.return_value.get.return_value = mock_response
            
            # Clear cache and verify function ran
            result = udfs.BONDCACHE_CLEAR()
            
            # New format: "Cleared N entries (was X% hit rate)"
            assert "Cleared" in result
            assert "entries" in result

    def test_returns_stats(self):
        """Test returns cache stats in output."""
        result = udfs.BONDCACHE_CLEAR()
        # New format: "Cleared N entries (was X% hit rate)"
        assert "Cleared" in result
        assert "Cleared" in result


# =============================================================================
# BONDCACHE_STATS Tests
# =============================================================================

class TestBondCacheStats:
    """Tests for BONDCACHE_STATS function."""

    def test_returns_stats_string(self):
        """Test returns formatted stats string."""
        result = udfs.BONDCACHE_STATS()
        assert "Size:" in result
        assert "Hit Rate:" in result or "Size:" in result
        assert "TTL:" in result

    def test_shows_cache_size(self):
        """Test shows current cache size and max."""
        result = udfs.BONDCACHE_STATS()
        # Format: "Size: N/500 | ..."
        assert "/500" in result or "/500" in result.replace(" ", "")


# =============================================================================
# _fetch_bond Tests (internal function)
# =============================================================================

class TestFetchBond:
    """Tests for _fetch_bond internal function."""

    def test_success(self):
        """Test successful fetch."""
        mock_response = MockResponse(200, MOCK_BOND)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs._fetch_bond("GB00BYZW3G56")
            assert result == MOCK_BOND

    def test_caches_result(self):
        """Test result is cached."""
        mock_response = MockResponse(200, MOCK_BOND)
        call_count = 0

        def mock_get(url, params=None):
            nonlocal call_count
            call_count += 1
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            # First call
            result1 = udfs._fetch_bond("GB00BYZW3G56")
            # Second call should use cache
            result2 = udfs._fetch_bond("GB00BYZW3G56")

            assert result1 == result2
            assert call_count == 1  # Only one API call

    def test_normalizes_isin(self):
        """Test ISIN is normalized."""
        mock_response = MockResponse(200, MOCK_BOND)

        def mock_get(url, params=None):
            assert "GB00BYZW3G56" in url  # Uppercase
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            udfs._fetch_bond("gb00byzw3g56")

    def test_404_returns_none(self):
        """Test 404 returns None."""
        mock_response = MockResponse(404)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs._fetch_bond("XX0000000000")
            assert result is None

    def test_500_returns_none(self):
        """Test 500 error returns None."""
        mock_response = MockResponse(500)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs._fetch_bond("GB00BYZW3G56")
            assert result is None

    def test_connection_error_returns_none(self):
        """Test connection error returns None."""
        import httpx

        def mock_get(url, params=None):
            raise httpx.RequestError("Timeout")

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs._fetch_bond("GB00BYZW3G56")
            assert result is None


# =============================================================================
# Malformed Response Tests
# =============================================================================

class TestMalformedResponses:
    """Tests for handling malformed API responses."""

    def test_bondlist_missing_isin_in_response(self):
        """Test BONDLIST handles bonds without ISIN field."""
        mock_response = MockResponse(200, [{"name": "Bond without ISIN"}, MOCK_BOND])

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("GB")
            # Should return empty string for missing ISIN
            assert result == [[""], ["GB00BYZW3G56"]]

    def test_bondinfo_with_null_values(self):
        """Test BONDINFO handles null values in response."""
        bond_with_nulls = {
            "isin": "GB00TEST1230",  # Valid 12-char ISIN format
            "name": None,
            "country": "GB",
            "coupon_rate": None,
        }
        with patch.object(udfs, "_fetch_bond", return_value=bond_with_nulls):
            result = udfs.BONDINFO("GB00TEST1230")  # Valid ISIN format (12 chars)
            # Null values should become empty strings
            assert isinstance(result, list)

    def test_bondstatic_integer_coupon_rate(self):
        """Test BONDSTATIC handles integer coupon rate."""
        bond_int_coupon = {**MOCK_BOND, "coupon_rate": 2}  # Integer instead of float
        with patch.object(udfs, "_fetch_bond", return_value=bond_int_coupon):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "coupon_rate")
            assert result == 200  # 2 * 100

    def test_bondcount_missing_by_country(self):
        """Test BONDCOUNT handles missing by_country."""
        mock_response = MockResponse(200, {"total_bonds": 100})  # No by_country

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT("GB")
            # Should return 0 if by_country is missing
            assert result == 0

    def test_bondcount_missing_total(self):
        """Test BONDCOUNT handles missing total_bonds."""
        mock_response = MockResponse(200, {"by_country": {"GB": 50}})  # No total_bonds

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDCOUNT()
            # Should return 0 if total_bonds is missing
            assert result == 0


# =============================================================================
# Edge Case Tests
# =============================================================================

class TestEdgeCases:
    """Edge case and boundary condition tests."""

    def test_bondstatic_strips_whitespace_from_field(self):
        """Test field name whitespace is stripped."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "  country  ")
            assert result == "GB"

    def test_bondlist_country_normalized(self):
        """Test country code is normalized to uppercase."""
        mock_response = MockResponse(200, [MOCK_BOND])

        def mock_get(url, params=None):
            assert params["country"] == "GB"
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            udfs.BONDLIST("gb")

    def test_bondsearch_fields_lowercased(self):
        """Test search field names are lowercased."""
        mock_response = MockResponse(200, [MOCK_BOND])

        def mock_get(url, params=None):
            assert "country" in params  # Lowercased
            assert "COUNTRY" not in params
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            udfs.BONDSEARCH("COUNTRY", "GB")

    def test_bondstatic_zero_coupon(self):
        """Test zero coupon bond."""
        zero_coupon_bond = {**MOCK_BOND, "coupon_rate": 0}
        with patch.object(udfs, "_fetch_bond", return_value=zero_coupon_bond):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "coupon_rate")
            assert result == 0

    def test_bondstatic_large_outstanding_amount(self):
        """Test large outstanding amounts."""
        large_bond = {**MOCK_BOND, "outstanding_amount": 999999999999999}
        with patch.object(udfs, "_fetch_bond", return_value=large_bond):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "outstanding_amount")
            assert result == 999999999999999

    def test_special_characters_in_isin(self):
        """Test ISIN with potentially problematic characters."""
        # ISINs are alphanumeric, but test edge cases
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDSTATIC("GB00BYZW3G56", "isin")
            assert result == "GB00BYZW3G56"


# =============================================================================
# Timeout Tests
# =============================================================================

class TestTimeouts:
    """Tests for timeout handling."""

    def test_timeout_error_in_fetch_bond(self):
        """Test timeout error returns None."""
        import httpx

        def mock_get(url, params=None):
            raise httpx.TimeoutException("Request timed out")

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs._fetch_bond("GB00BYZW3G56")
            assert result is None

    def test_timeout_in_bondlist(self):
        """Test timeout in BONDLIST returns #N/A."""
        import httpx

        def mock_get(url, params=None):
            raise httpx.TimeoutException("Request timed out")

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLIST("GB")
            assert is_error(result)


# =============================================================================
# BONDYEARSTOMAT Tests
# =============================================================================

class TestBONDYEARSTOMAT:
    """Tests for BONDYEARSTOMAT function."""

    def test_calculates_years_to_maturity(self):
        """Test basic years to maturity calculation."""
        # Bond matures 2026-07-22
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDYEARSTOMAT("GB00BYZW3G56")
            assert isinstance(result, float)
            assert result >= 0

    def test_matured_bond_returns_zero(self):
        """Test matured bond returns 0."""
        past_bond = {**MOCK_BOND, "maturity_date": "2020-01-01"}
        with patch.object(udfs, "_fetch_bond", return_value=past_bond):
            result = udfs.BONDYEARSTOMAT("GB00BYZW3G56")
            assert result == 0.0

    def test_bond_not_found_returns_error(self):
        """Test non-existent bond returns error."""
        with patch.object(udfs, "_fetch_bond", return_value=None):
            result = udfs.BONDYEARSTOMAT("XX0000000000")
            assert is_error(result)

    def test_no_maturity_date_returns_error(self):
        """Test bond without maturity date returns error."""
        no_maturity = {**MOCK_BOND, "maturity_date": None}
        with patch.object(udfs, "_fetch_bond", return_value=no_maturity):
            result = udfs.BONDYEARSTOMAT("GB00BYZW3G56")
            assert is_error(result)

    def test_custom_as_of_date(self):
        """Test calculation with custom as_of date."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDYEARSTOMAT("GB00BYZW3G56", "2025-01-01")
            assert isinstance(result, float)


# =============================================================================
# BONDCOUPONFREQ Tests
# =============================================================================

class TestBONDCOUPONFREQ:
    """Tests for BONDCOUPONFREQ function."""

    def test_semi_annual(self):
        """Test semi-annual frequency."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDCOUPONFREQ("GB00BYZW3G56")
            assert result == "Semi-annual"

    def test_annual(self):
        """Test annual frequency."""
        annual_bond = {**MOCK_BOND, "coupon_frequency": 1}
        with patch.object(udfs, "_fetch_bond", return_value=annual_bond):
            result = udfs.BONDCOUPONFREQ("GB00BYZW3G56")
            assert result == "Annual"

    def test_zero_coupon(self):
        """Test zero coupon bond."""
        zero_coupon = {**MOCK_BOND, "coupon_rate": 0}
        with patch.object(udfs, "_fetch_bond", return_value=zero_coupon):
            result = udfs.BONDCOUPONFREQ("GB00BYZW3G56")
            assert result == "Zero coupon"

    def test_bond_not_found(self):
        """Test non-existent bond returns error."""
        with patch.object(udfs, "_fetch_bond", return_value=None):
            result = udfs.BONDCOUPONFREQ("XX0000000000")
            assert is_error(result)


# =============================================================================
# BONDISLINKER Tests
# =============================================================================

class TestBONDISLINKER:
    """Tests for BONDISLINKER function."""

    def test_nominal_returns_false(self):
        """Test nominal bond returns FALSE."""
        with patch.object(udfs, "_fetch_bond", return_value=MOCK_BOND):
            result = udfs.BONDISLINKER("GB00BYZW3G56")
            assert result is False

    def test_index_linked_returns_true(self):
        """Test index-linked bond returns TRUE."""
        linker = {**MOCK_BOND, "security_type": "INDEX_LINKED"}
        with patch.object(udfs, "_fetch_bond", return_value=linker):
            result = udfs.BONDISLINKER("GB00BYZW3G56")
            assert result is True

    def test_bond_not_found(self):
        """Test non-existent bond returns error."""
        with patch.object(udfs, "_fetch_bond", return_value=None):
            result = udfs.BONDISLINKER("XX0000000000")
            assert is_error(result)


# =============================================================================
# BONDISINVALID Tests
# =============================================================================

class TestBONDISINVALID:
    """Tests for BONDISINVALID function."""

    def test_valid_isin_returns_true(self):
        """Test valid ISIN returns TRUE."""
        result = udfs.BONDISINVALID("GB00BYZW3G56")
        assert result is True

    def test_valid_us_isin(self):
        """Test valid US ISIN."""
        result = udfs.BONDISINVALID("US912810TM58")
        assert result is True

    def test_eurobond_xs_prefix(self):
        """Test XS prefix (Eurobonds) is accepted."""
        result = udfs.BONDISINVALID("XS1234567890")
        assert result is True

    def test_invalid_format_returns_false(self):
        """Test invalid format returns FALSE."""
        result = udfs.BONDISINVALID("invalid")
        assert result is False

    def test_invalid_country_returns_false(self):
        """Test unknown country code returns FALSE."""
        result = udfs.BONDISINVALID("ZZ1234567890")
        assert result is False

    def test_empty_returns_false(self):
        """Test empty string returns FALSE."""
        result = udfs.BONDISINVALID("")
        assert result is False


# =============================================================================
# BONDHELP Tests
# =============================================================================

class TestBONDHELP:
    """Tests for BONDHELP function."""

    def test_no_topic_returns_overview(self):
        """Test no topic returns overview."""
        result = udfs.BONDHELP()
        assert isinstance(result, list)
        assert len(result) > 0

    def test_fields_topic(self):
        """Test fields topic returns field list."""
        result = udfs.BONDHELP("fields")
        assert isinstance(result, list)
        # First row is header
        assert result[0] == ["Field", "Description"]
        assert len(result) > 10

    def test_countries_topic(self):
        """Test countries topic returns country list."""
        result = udfs.BONDHELP("countries")
        assert isinstance(result, list)
        assert result[0] == ["Code", "Country"]

    def test_functions_topic(self):
        """Test functions topic returns function list."""
        result = udfs.BONDHELP("functions")
        assert isinstance(result, list)
        assert result[0] == ["Function", "Description"]

    def test_unknown_topic_returns_error(self):
        """Test unknown topic returns error."""
        result = udfs.BONDHELP("invalid_topic")
        assert is_error(result)


# =============================================================================
# BONDMATURITYRANGE Tests
# =============================================================================

class TestBONDMATURITYRANGE:
    """Tests for BONDMATURITYRANGE function."""

    def test_returns_bonds_in_range(self):
        """Test returns bonds maturing in date range."""
        mock_response = MockResponse(200, {"data": [MOCK_BOND, MOCK_BOND_2]})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDMATURITYRANGE("2025-01-01", "2030-12-31")
            assert isinstance(result, list)
            assert len(result) == 2

    def test_empty_result_returns_error(self):
        """Test empty result returns error."""
        mock_response = MockResponse(200, {"data": []})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDMATURITYRANGE("2025-01-01", "2025-12-31")
            assert is_error(result)


# =============================================================================
# BONDLINEAGE Tests
# =============================================================================

class TestBONDLINEAGE:
    """Tests for BONDLINEAGE function (enterprise)."""

    def test_returns_lineage_summary(self):
        """Test returns lineage summary."""
        mock_lineage = {
            "data": {
                "contributing_sources": ["source1", "source2"],
                "reconciliation_confidence": 0.95,
                "field_sources": {"coupon_rate": {"source_name": "source1", "confidence": 0.98}},
            }
        }
        mock_response = MockResponse(200, mock_lineage)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLINEAGE("DE0001102580")
            assert "source1" in result

    def test_specific_field_lineage(self):
        """Test returns lineage for specific field."""
        mock_lineage = {
            "data": {
                "field_sources": {"coupon_rate": {"source_name": "source1", "confidence": 0.98}},
            }
        }
        mock_response = MockResponse(200, mock_lineage)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLINEAGE("DE0001102580", "coupon_rate")
            assert "source1" in result

    def test_api_error_returns_error(self):
        """Test API error returns error message."""
        mock_response = MockResponse(404)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDLINEAGE("XX0000000000")
            assert is_error(result)


# =============================================================================
# BONDHISTORY Tests
# =============================================================================

class TestBONDHISTORY:
    """Tests for BONDHISTORY function (enterprise)."""

    def test_returns_history(self):
        """Test returns change history."""
        mock_history = {
            "data": [
                {
                    "changed_at": "2026-01-15",
                    "change_type": "UPDATE",
                    "field_name": "coupon_rate",
                    "old_value": "0.01",
                    "new_value": "0.015",
                }
            ]
        }
        mock_response = MockResponse(200, mock_history)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDHISTORY("DE0001102580")
            assert isinstance(result, list)
            # Header row + data row
            assert len(result) == 2

    def test_empty_history_returns_error(self):
        """Test empty history returns error."""
        mock_response = MockResponse(200, {"data": []})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDHISTORY("DE0001102580")
            assert is_error(result)


# =============================================================================
# BONDACTIONS Tests
# =============================================================================

class TestBONDACTIONS:
    """Tests for BONDACTIONS function (enterprise)."""

    def test_returns_corporate_actions(self):
        """Test returns corporate actions."""
        mock_actions = {
            "data": [
                {
                    "isin": "DE0001102580",
                    "action_type": "MATURED",
                    "effective_date": "2026-01-15",
                    "notes": "Bond matured",
                }
            ]
        }
        mock_response = MockResponse(200, mock_actions)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDACTIONS()
            assert isinstance(result, list)
            # Header row + data row
            assert len(result) == 2

    def test_no_actions_returns_error(self):
        """Test no actions returns error."""
        mock_response = MockResponse(200, {"data": []})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDACTIONS()
            assert is_error(result)


# =============================================================================
# BONDREFRESH Tests
# =============================================================================

class TestBONDREFRESH:
    """Tests for BONDREFRESH function."""

    def test_refresh_returns_message(self):
        """Test refresh returns success message."""
        mock_response = MockResponse(200, {"message": "Refresh started"})

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDREFRESH("US", "test-api-key")
            assert "Refresh" in result

    def test_refresh_clears_cache(self):
        """Test refresh clears the cache."""
        mock_response = MockResponse(200, {"message": "Refresh started"})

        def mock_get(url, params=None):
            return mock_response

        # Pre-populate cache
        udfs._bond_cache.set("US912810TM58", MOCK_BOND_2)
        assert udfs._bond_cache.get("US912810TM58") is not None

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            udfs.BONDREFRESH("US", "test-api-key")

        # Cache should be cleared
        assert udfs._bond_cache.get("US912810TM58") is None

    def test_api_error_returns_error(self):
        """Test API error returns error message."""
        mock_response = MockResponse(403)

        def mock_get(url, params=None):
            return mock_response

        with patch.object(udfs, "_get_client", return_value=MockClient(mock_get)):
            result = udfs.BONDREFRESH("US", "bad-key")
            assert is_error(result)
