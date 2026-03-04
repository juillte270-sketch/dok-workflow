---
name: naming-3step
description: "Ward Cunningham's 3-step method for intention-revealing names. Apply to functions, variables, files, and config keys. Good names ARE the essence."
argument_hint: "[code element to rename, e.g. 'the price lookup function' or 'review all names in fill_prices.py']"
---

# Naming 3-Step (의도 드러내는 네이밍)

From Ward Cunningham (wiki inventor, 50+ years programming):
Names are the essence of code. A good name eliminates the need for comments.

## The 3 Steps

### Step 1: Reveal Intent (의도 드러내기)
Make the name say what it DOES, even if it gets long.

```python
# Bad
def process(data):
def get_price(item):
def check(row):

# Good (long but clear)
def fill_purchase_and_selling_prices_from_price_sheet(master_path, date):
def lookup_purchase_price_from_price_map(price_map, item_name, unit):
def check_if_row_is_stock_or_subdivision(col1_value):
```

### Step 2: Hide Implementation (구현 숨기기)
Remove HOW it works from the name. Only WHAT it achieves matters.

```python
# Step 1 result (too much implementation detail)
def iterate_price_map_and_find_best_similarity_match(price_map, name):

# Step 2 result (implementation hidden)
def lookup_purchase_price(price_map, item_name, unit):
```

### Step 3: Hint at Return Type (반환 타입 힌트)
The name should imply what you get back.

```python
# No hint about what's returned
def lookup(item):

# Clear return hint
def lookup_purchase_price(...)  →  returns int (price)
def lookup_fixed_price(...)     →  returns (int, str|None) tuple
def find_drive_folder(...)      →  returns str (folder path)
def is_stock_row(...)           →  returns bool
def get_store_list(...)         →  returns list
```

## Application to Our Project

### Script Names (`execution/*.py`)
| Current | Issue | Suggested |
|---------|-------|-----------|
| `fill_prices.py` | OK — clear intent | Keep |
| `process_orders.py` | "process" is vague | `parse_and_classify_orders.py` |
| `simple_docx_converter.py` | What's "simple"? | `delivery_list_docx_builder.py` |

### Function Names
| Current | Improved |
|---------|----------|
| `_has_any_price(entries)` | OK — boolean hint via `has_` |
| `_best_from_entries(entries)` | `_select_best_price_from_entries()` |
| `_name_similarity(a, b)` | `_calculate_name_overlap_ratio()` |

### Config Keys (mappings.json)
| Current | Assessment |
|---------|------------|
| `item_supplier_map` | Good — {item} → {supplier} |
| `page_config` | Good — page layout config |
| `helo_code_map` | Good — HELO mATNR code → item |
| `inventory_reorder_rules` | Good — reorder rules per item |

## When to Apply
- Creating a new function or script
- During code review
- When a name causes confusion (had to read code to understand)
- When refactoring after a bug fix

## Output Format
```
[naming-3step] Element: {function/variable/file}
[naming-3step] Current: {current name}
[naming-3step] Step 1 (intent): {long but clear name}
[naming-3step] Step 2 (hide impl): {simplified}
[naming-3step] Step 3 (return hint): {final name}
[naming-3step] Recommendation: {rename / keep current}
```
