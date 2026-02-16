# Changelog

All notable changes to bondmaster-excel will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

### Fixed
- **xlOil 0.21 compatibility:** Removed `category` parameter from `@xlo.func` decorators (was causing TypeError on module load)
- **Installation docs:** Fixed pip install instructions (packages are on GitHub, not PyPI)
- **xlOil config:** Corrected LoadModules to use `bondmaster_excel.udfs` instead of `bondmaster_excel`

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
