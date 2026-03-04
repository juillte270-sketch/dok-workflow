---
name: verify-auction-config
description: "Check ITEM_OVERRIDES in fill_auction_from_api.py match directive. Detects wrong unit conversions, missing filters, and note format errors."
---

# Verify Auction Config

Ensure fill_auction_from_api.py ITEM_OVERRIDES matches workflow_auction_prices.md.

## Files to Read
- `execution/fill_auction_from_api.py` → ITEM_OVERRIDES dict, ITEM_NOTES dict
- `directives/workflow_auction_prices.md` → ITEM_OVERRIDES section

## Checks

### 1. ITEM_OVERRIDES Completeness
- [ ] Every override in the directive exists in the script
- [ ] Every override in the script is documented in the directive
- [ ] No orphaned entries

### 2. Override Field Consistency (per item)
- [ ] `search_name` matches
- [ ] `grade_filter` / `unit_filter` matches
- [ ] `multiply` value matches (unit conversion)
- [ ] `filter_item_exact` flag matches where applicable
- [ ] `use_sheet_grade` flag matches where applicable

### 3. ITEM_NOTES
- [ ] Items with "(흙)" note format are correct
- [ ] Known items with notes: 깐양배추45, 깐양파, 깐대파
- [ ] number_format pattern is `#,##0"(흙)"`

### 4. API Configuration
- [ ] API URL is correct: `http://www.garak.co.kr/homepage/publicdata/dataXmlOpen.do`
- [ ] Encoding is UTF-8 (not euc-kr)
- [ ] Date logic: `--date` + 1 day for API call

### 5. Skip Logic
- [ ] 소분/재고 rows are skipped (Col A check)
- [ ] Items without API data are documented: 레몬(수입), 간마늘(가공), 라디치오(수입)

## Report Format
```
[auction-config] ITEM_OVERRIDES: ✅ 7 items consistent
[auction-config] ITEM_NOTES: ✅ 3 items verified
[auction-config] API config: ✅ URL/encoding correct
```
