# Changelog

All notable changes to bondmaster-excel will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

## [0.2.0] - 2026-02-17

### Added
- **bond-master v2.0 API support**:
  - Handle 202 responses (auto-lookup in progress)
  - Show "🔄 Looking up..." while background search runs
  - Show "⚠️ Not found after search" when lookup exhausted
- **BONDNAMESEARCH function**: Search bonds by name using full-text search
  - Example: `=BONDNAMESEARCH("OATEI 2030")` returns matching ISINs
  - Uses new `/v1/search` API endpoint
- **Auto-lookup UX**: BONDSTATIC and BONDINFO now trigger auto-lookup for unknown ISINs

### Changed
- `_api_request()` now handles 202 status codes
- `_fetch_bond()` returns status dict when lookup is in progress
- BONDHELP updated to include BONDNAMESEARCH

### Compatibility
- Requires bond-master v0.3.0+ for full auto-lookup functionality
- Backwards compatible with v0.2.x (no auto-lookup, returns not found)

### Fixed
- **xlOil 0.21 compatibility:** Removed `category` parameter from `@xlo.func` decorators (was causing TypeError on module load)
- **Installation docs:** Fixed pip install instructions (packages are on GitHub, not PyPI)
- **xlOil config:** Corrected LoadModules to use `bondmaster_excel.udfs` instead of `bondmaster_excel`

### Improved
- **Installation guide overhaul** (battle-tested on Windows Server 2025):
  - Added architecture matching warning (64-bit Excel ↔ 64-bit Python)
  - Added Option B for system-wide install (simpler for servers)
  - Added XLSTART folder creation step
  - Added xlOil DLL copy instructions for missing dependency errors
  - Added Visual C++ Redistributable requirement
  - Added Windows security unblock note for downloaded XLLs
  - Expanded troubleshooting with more error scenarios

### Added
- **Analytics Functions**: BONDYEARSTOMAT, BONDMATURITYRANGE, BONDCOUPONFREQ, BONDISLINKER
- **Enterprise Functions**: BONDLINEAGE, BONDHISTORY, BONDACTIONS, BONDREFRESH
- **Utility Functions**: BONDHELP (built-in help), BONDISINVALID (ISIN validation)
- **User-Friendly Errors**: Clear messages with ⚠️ prefix instead of #VALUE! errors
- **Field Shortcuts**: coupon→coupon_rate, maturity→maturity_date, type→security_type
- **Documentation**: Comprehensive README, GettingStarted.md tutorial, PortfolioTemplate.csv
- **CI/CD Pipeline**: GitHub Actions with test, lint, type-check, security jobs
- **Test Coverage**: 104 tests, 92% coverage

### Changed
- Error messages now explain the issue and suggest fixes
- Coupon rates displayed as percentages (1.5 instead of 0.015)
- Country validation with helpful error messages

### Fixed
- Type annotations for mypy compatibility (dict[str, Any] for mixed params)
- Trailing whitespace in docstrings
- Simplified return statements (SIM103)

## [0.1.0] - 2026-02-09

### Added
- Initial release
- Core UDFs: BONDSTATIC, BONDINFO, BONDLIST, BONDSEARCH, BONDCOUNT
- Utility UDFs: BONDAPI_STATUS, BONDCACHE_CLEAR, BONDCACHE_STATS
- TTL-based LRU cache (5 min default)
- Thread-safe HTTP client
- xlOil integration for native Excel XLL add-in
