---
name: verify-fixed-prices
description: "Check FIXED_PRICES in fill_prices.py matches MEMORY.md and directive. Detects price mismatches and missing items."
---

# Verify Fixed Prices

Ensure fixed purchase prices are consistent across all sources.

## Files to Read
- `execution/fill_prices.py` → `FIXED_PRICES` dict
- `MEMORY.md` (auto-memory) → "고정 매입단가 품목" section (search in memory dir too)
- `directives/workflow_price_filling.md` → FIXED_PRICES table

## Checks

### 1. Source Consistency
- [ ] FIXED_PRICES dict keys in fill_prices.py
- [ ] MEMORY.md 고정 매입단가 품목 list
- [ ] workflow_price_filling.md 고정 매입단가 table
- All three sources must have **identical** items and prices

### 2. Price Values
- [ ] All prices are positive integers
- [ ] No suspiciously low prices (< 500원) except 무순(대)=800
- [ ] No suspiciously high prices (> 50,000원)
- [ ] Paired items have consistent prices:
  - 곱슬이콩나물 == 일자콩나물
  - 아보카도 == 중숙아보카도

### 3. Coverage
- [ ] All items that are commonly ordered but NOT in 단가시트 should be in FIXED_PRICES
- [ ] Known fixed items: 숙주, 굵은숙주, 새싹(대), 무순(대), 중란, 대란, 콩나물(2종), 아보카도(2종), 파김치, 맛김치

### 4. Matching Logic
- [ ] `lookup_fixed_price()` does exact match first, then partial
- [ ] Partial match uses longest-key-first to avoid false matches
- [ ] Fixed price check happens BEFORE 단가시트 lookup in fill loop

## Report Format
```
[fixed-prices] fill_prices.py: 12 items
[fixed-prices] MEMORY.md: 12 items
[fixed-prices] directive: 12 items
[fixed-prices] Consistency: ✅ All match (or ❌ Mismatch: 파김치 19000 vs 18000)
```
