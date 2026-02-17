# Product Requirements Document
## bond-master + bondmaster-excel

**Version:** 1.0  
**Date:** 2026-02-17  
**Status:** Approved  
**Owner:** MapleCap Partners

> **Note:** This is a shared PRD covering both bond-master (backend) and bondmaster-excel (Excel frontend). The canonical version lives in [bond-master/docs/PRD.md](https://github.com/JeffExec/bond-master/blob/main/docs/PRD.md).

---

## 1. Problem Statement

MapleCap bond traders, quants, and analysts need static reference data for government bonds across 8 markets. Currently this requires:

- Manual lookups across multiple public sources
- Inconsistent data formats between sources
- No single source of truth
- Technical barrier for non-Python users

**Impact:** Time wasted on data gathering, risk of errors, not everyone can access the data they need.

**Solution:** A unified bond reference data system that:
- Pulls data from public sources automatically
- Caches in a database (local or shared)
- Provides easy access via Excel functions and REST API
- Handles missing/incorrect data gracefully

---

## 2. Users

| Persona | Tools | Primary Needs |
|---------|-------|---------------|
| Bond Trader | Excel | Quick ISIN lookups, bond lists by country |
| Quant | Python | Programmatic access via API/client |
| Analyst | Excel | Static data for reports, complete bond lists |

**Scope:** MapleCap internal users only.

---

## 3. Requirements

### 3.1 Functional Requirements (Must-Have)

| ID | Requirement | Acceptance Criteria |
|----|-------------|---------------------|
| F1 | List all bonds for a given country | `=BONDLIST("US")` returns all active US Treasuries |
| F2 | Get static data for any ISIN | `=BONDSTATIC("US912810TM58", "coupon_rate")` returns coupon |
| F3 | Support nominal AND inflation-linked bonds | OATEi, TIPS, Index-linked Gilts all included |
| F4 | Excel access (native functions) | No macro warnings, works like built-in functions |
| F5 | REST API access | OpenAPI-documented endpoints |
| F6 | Python client library | `from bondmaster import BondClient` |
| F7 | Local caching (offline mode) | Excel works fully offline after initial sync |
| F8 | Multi-source with fallbacks | If primary source fails, try secondary |
| F9 | Search by partial name | `=BONDSEARCH("OATEI 2030")` finds matching bonds |
| F10 | Shared database support | Users can run locally OR connect to central Postgres |
| F11 | On-demand fetch for missing bonds | Request unknown ISIN → system looks it up automatically |
| F12 | Rate-limited background sync | New bonds fetched without hammering sources |

### 3.2 Data Quality Requirements

| ID | Requirement | Acceptance Criteria |
|----|-------------|---------------------|
| D1 | Audit trail / data provenance | Every field tracks which source it came from and when |
| D2 | Data correction workflow | User can flag incorrect data; system replaces from alternative source (not free-text) |
| D3 | Data freshness indicator | User can see when data was last updated |

### 3.3 UX Requirements

| ID | Requirement | Acceptance Criteria |
|----|-------------|---------------------|
| U1 | Easy installation | Non-technical user can install in <30 minutes following README |
| U2 | "Looking up..." feedback | When bond not in cache, user sees status message, not just error |
| U3 | Human-readable errors | Clear messages explaining what went wrong and how to fix |

### 3.4 Non-Functional Requirements

| ID | Requirement | Target |
|----|-------------|--------|
| N1 | Lookup performance | <1 second for any cached ISIN |
| N2 | Public data only | No T&C violations; only use publicly available APIs |
| N3 | Rate limiting | Respectful request rates to avoid source blocking |

### 3.5 Markets (Must-Have)

| Flag | Country | Security Type | Primary Source |
|------|---------|---------------|----------------|
| 🇺🇸 | USA | Treasuries | Treasury Fiscal Data API |
| 🇬🇧 | UK | Gilts | UK DMO |
| 🇩🇪 | Germany | Bunds/Bobls/Schatz | Deutsche Finanzagentur |
| 🇫🇷 | France | OATs/OATi/OATei | Agence France Trésor |
| 🇮🇹 | Italy | BTPs/CCTs | Borsa Italiana / MEF |
| 🇪🇸 | Spain | Bonos/Obligaciones | MTS Data |
| 🇳🇱 | Netherlands | DSLs | DSTA |
| 🇯🇵 | Japan | JGBs | MOF Japan |

**Bond Types:** Nominal AND inflation-linked (linkers) for all markets.

---

## 4. Success Metrics

| Metric | Target | Measurement |
|--------|--------|-------------|
| Bond list completeness | 100% of active govt bonds per market | Compare against official issuance lists |
| Data retrieval speed | <1s for any cached ISIN | API response time monitoring |
| Installation success rate | 100% of users can install | Support ticket tracking |
| User adoption | All MapleCap bond data users | Usage analytics |
| Data accuracy | Zero user-reported errors sustained | Error flagging system |

---

## 5. Scope

### 5.1 In Scope (MVP — Delivered)

- ✅ 8 markets with multi-source collection
- ✅ SQLite local database
- ✅ CLI for fetch/refresh
- ✅ 18+ Excel functions via xlOil
- ✅ REST API with OpenAPI docs
- ✅ Python client library

### 5.2 In Scope (v2.0 — To Build)

- 🔲 Shared PostgreSQL database option
- 🔲 On-demand fetch for missing bonds
- 🔲 Rate-limited background sync for new issues
- 🔲 Audit trail / data provenance
- 🔲 Data correction workflow
- 🔲 Search by partial name (`BONDSEARCH`)
- 🔲 "Looking up..." async UX

### 5.3 Out of Scope

- Additional markets (AU, CA, CH, etc.) — future consideration
- Live pricing data — this is static reference data only
- One-click installer — user-friendly docs suffice for now
- Mobile app

---

## 6. Risks

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Data source website changes | High | High | Multi-source architecture with fallbacks (built) |
| T&C violation / IP blocking | Medium | High | Public APIs only, rate limiting, respectful scraping |
| Installation complexity | Medium | Medium | Detailed README with troubleshooting (PR #9) |
| Excel/xlOil compatibility | Low | Medium | Pin versions, test on Excel updates |
| Data accuracy issues | Medium | Medium | Audit trail + correction workflow (v2.0) |

---

## 7. Dependencies

- **xlOil:** Excel add-in framework (external dependency)
- **Public data sources:** Treasury Direct, UK DMO, etc. (external)
- **Python 3.11+:** Runtime requirement

---

## 8. Timeline

| Phase | Scope | Status |
|-------|-------|--------|
| MVP | Core functionality (8 markets, Excel, API) | ✅ Complete |
| v2.0 | Shared DB, on-demand fetch, audit trail | 🔲 Planning |

---

## Appendix A: Related Documents

- [bond-master README](https://github.com/JeffExec/bond-master)
- [bondmaster-excel README](../README.md)
- [ARCHITECTURE.md](../ARCHITECTURE.md)

---

*Document approved by Kamil Szynkarczuk, 2026-02-17*
