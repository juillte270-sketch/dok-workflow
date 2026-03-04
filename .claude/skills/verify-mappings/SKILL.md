---
name: verify-mappings
description: "Check mappings.json consistency: supplier aliases, inventory maps, page config, helo codes. Detects orphaned entries and missing mappings."
---

# Verify Mappings

Check `skills/order_processing/resources/mappings.json` for internal consistency.

## Files to Read
- `skills/order_processing/resources/mappings.json`
- `execution/create_price_sheet.py` (SUPPLIER_ALIASES)
- `execution/fill_inventory_sheet.py` (inventory_item_map usage)
- `execution/scrape_helo.py` (helo_code_map usage)

## Checks

### 1. item_supplier_map
- [ ] All supplier names are valid (no typos, consistent naming)
- [ ] No duplicate items across different suppliers (unless intentional)
- [ ] Key items present: 양파, 감자, 배추, 대파, 마늘, 오이, 당근, etc.

### 2. SUPPLIER_ALIASES consistency
- [ ] `create_price_sheet.py` SUPPLIER_ALIASES matches mappings
- [ ] Bidirectional aliases work (경향농산 ↔ 다모아버섯)

### 3. inventory_item_map
- [ ] All items in map are real inventory items
- [ ] No items with same target (collision check)
- [ ] Key items: 깐양파, 간마늘, 팽이버섯, 새송이, etc.

### 4. helo_code_map
- [ ] All mATNR codes have valid mappings (or explicit None)
- [ ] unit_type is one of: weight, count, multiply, skip
- [ ] Items with None mapping are documented (intentional skip)

### 5. page_config
- [ ] default and friday configs both exist
- [ ] All known stores appear in at least one page
- [ ] No duplicate stores within same page (except 동원1차 on pages 1 and 3)
- [ ] SUPPLIER_LIST covers all active suppliers

### 6. inventory_reorder_rules
- [ ] All inventory items have reorder rules
- [ ] type is one of: standard, special_box, manual, skip
- [ ] threshold and target values are reasonable (positive integers)

## Report Format
For each check, report PASS or list specific issues:
```
[mappings] item_supplier_map: ✅ 15 suppliers, 87 items
[mappings] SUPPLIER_ALIASES: ❌ 풍경 alias missing in mappings
[mappings] inventory_item_map: ✅ 25 items mapped
...
```
