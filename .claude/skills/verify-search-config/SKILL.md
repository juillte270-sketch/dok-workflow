---
name: verify-search-config
description: "Check SEARCH_CONFIG and MULTI_ROW_ITEMS in fill_sikbom_targeted.py match directive. Detects missing items, wrong suppliers, and unit conversion errors."
---

# Verify Search Config (Sikbom Prices)

Ensure fill_sikbom_targeted.py configuration matches workflow_sikbom_prices.md.

## Files to Read
- `execution/fill_sikbom_targeted.py` → SEARCH_CONFIG list, MULTI_ROW_ITEMS dict
- `directives/workflow_sikbom_prices.md` → SEARCH_CONFIG table, MULTI_ROW_ITEMS section

## Checks

### 1. SEARCH_CONFIG Completeness
- [ ] Every item in the directive's SEARCH_CONFIG table exists in the script
- [ ] Every item in the script's SEARCH_CONFIG exists in the directive
- [ ] No orphaned entries in either direction

### 2. Field Consistency (per item)
For each SEARCH_CONFIG entry, compare script vs directive:
- [ ] `search` keyword matches
- [ ] `suppliers` list matches (order matters = priority)
- [ ] `filter` value matches
- [ ] `exclude` value matches
- [ ] `unit_convert` value matches
- [ ] `match_grade` / `match_unit` values match

### 3. MULTI_ROW_ITEMS
- [ ] Script's MULTI_ROW_ITEMS dict matches directive's table
- [ ] Items with match_grade/match_unit in SEARCH_CONFIG are listed in MULTI_ROW_ITEMS
- [ ] Mode is correct: 'grade+unit' or 'grade'

### 4. Supplier Validity
- [ ] All supplier names are real 식봄 supplier names
- [ ] Known suppliers: CJ프레시웨이, 다봄푸드, 세현F&B, 케이에프피(강남), 농장에서바로, 쉐프의정원, 바름푸드

### 5. Unit Conversion Logic
- [ ] Items with unit_convert have correct multiplier
- [ ] Conversion documented in directive's 단위 환산 table
- [ ] Known conversions: 깐양배추45=×5, 깐양배추42=×4, 깐마늘20K=×20, 애호박특=×20

## Report Format
```
[search-config] Script: 24 items, Directive: 24 items
[search-config] Consistency: ✅ All match
[search-config] MULTI_ROW_ITEMS: ✅ 2 items (깐마늘 대서, 애호박)
[search-config] Unit conversions: ✅ 4 items verified
```
