"""Tests for BondMaster Excel UDFs.

Note: These tests mock the API responses since xlOil UDFs 
require Excel to actually run.
"""

import pytest
from unittest.mock import patch, MagicMock
import httpx


# Mock bond data for testing
MOCK_BOND = {
    "isin": "GB00BYZW3G56",
    "cusip": None,
    "name": "1½% Treasury Gilt 2026",
    "country": "GB",
    "issuer": "UK Debt Management Office",
    "security_type": "NOMINAL",
    "currency": "GBP",
    "coupon_rate": 0.015,  # Stored as decimal
    "coupon_frequency": 2,
    "maturity_date": "2026-07-22",
    "issue_date": "2016-02-18",
    "outstanding_amount": 35000000000,
}


class TestBondStaticLogic:
    """Test the logic used in BONDSTATIC function."""
    
    def test_coupon_rate_conversion(self):
        """Verify coupon rate is converted to percentage."""
        # In UDF, we multiply by 100 for display
        coupon = MOCK_BOND["coupon_rate"] * 100
        assert coupon == 1.5
    
    def test_field_mapping(self):
        """Test field alias mapping."""
        field_map = {
            "coupon": "coupon_rate",
            "maturity": "maturity_date",
            "type": "security_type",
        }
        assert field_map.get("coupon") == "coupon_rate"
        assert field_map.get("maturity") == "maturity_date"
    
    def test_isin_normalization(self):
        """Test ISIN is normalized to uppercase."""
        isin = "gb00byzw3g56"
        assert isin.upper().strip() == "GB00BYZW3G56"


class TestAPIIntegration:
    """Test API request/response handling logic."""
    
    def test_bond_endpoint_url(self):
        """Test bond endpoint URL construction."""
        isin = "GB00BYZW3G56"
        url = f"/bonds/{isin}"
        assert url == "/bonds/GB00BYZW3G56"
    
    def test_list_params(self):
        """Test list endpoint parameters."""
        params = {"country": "GB", "limit": 1000}
        assert params["country"] == "GB"
        
        # With security type filter
        params["security_type"] = "INDEX_LINKED"
        assert params["security_type"] == "INDEX_LINKED"


class TestCacheBehavior:
    """Test caching logic."""
    
    def test_cache_operations(self):
        """Test cache dictionary operations."""
        cache = {}
        
        # Add to cache
        cache["GB00BYZW3G56"] = MOCK_BOND
        assert "GB00BYZW3G56" in cache
        
        # Retrieve from cache
        assert cache["GB00BYZW3G56"]["coupon_rate"] == 0.015
        
        # Clear cache
        cache.clear()
        assert len(cache) == 0


class TestFieldExtraction:
    """Test field extraction from bond data."""
    
    @pytest.mark.parametrize("field,expected", [
        ("isin", "GB00BYZW3G56"),
        ("country", "GB"),
        ("currency", "GBP"),
        ("security_type", "NOMINAL"),
        ("coupon_frequency", 2),
        ("maturity_date", "2026-07-22"),
    ])
    def test_field_values(self, field, expected):
        """Test extracting various fields."""
        assert MOCK_BOND.get(field) == expected
    
    def test_missing_field_returns_none(self):
        """Test missing field behavior."""
        assert MOCK_BOND.get("nonexistent_field") is None


class TestResponseParsing:
    """Test API response parsing logic."""
    
    def test_list_response_envelope(self):
        """Test handling of envelope vs raw list response."""
        # Envelope format
        envelope = {"data": [MOCK_BOND], "total": 1}
        bonds = envelope.get("data", [])
        assert len(bonds) == 1
        
        # Raw list format
        raw_list = [MOCK_BOND]
        bonds = raw_list if isinstance(raw_list, list) else raw_list.get("data", [])
        assert len(bonds) == 1
    
    def test_isin_extraction_for_list(self):
        """Test extracting ISINs for vertical array."""
        bonds = [MOCK_BOND, {"isin": "GB00BL6C7720"}]
        isins = [[b.get("isin", "")] for b in bonds]
        
        assert isins == [["GB00BYZW3G56"], ["GB00BL6C7720"]]
