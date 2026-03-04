---
name: verify-directives
description: "Check execution scripts have matching directives and vice versa. Detects orphaned scripts, outdated directives, and missing documentation."
---

# Verify Directives

Ensure execution scripts and directives are in sync.

## Files to Read
- All `directives/workflow_*.md` files
- All `execution/*.py` files (primary/active ones only)
- `MEMORY.md` → 디렉티브 목록 section

## Active Scripts (primary, not debug/test/one-off)

| Script | Expected Directive |
|--------|--------------------|
| process_orders.py | workflow_01_order_processing.md |
| simple_docx_converter.py | workflow_01_order_processing.md |
| update_order_sheet.py | workflow_02_order_entry.md |
| generate_invoices.py | workflow_generate_invoices.md |
| create_price_sheet.py | workflow_create_price_sheet.md |
| fill_sikbom_targeted.py | workflow_sikbom_prices.md |
| fill_auction_from_api.py | workflow_auction_prices.md |
| fill_prices.py | workflow_price_filling.md |
| fill_inventory_sheet.py | workflow_inventory_shipment.md |
| manage_inventory_date.py | workflow_inventory_date.md |
| scrape_helo.py | workflow_helo_scraping.md |
| send_daily_summary.py | workflow_daily_summary.md |
| generate_final_invoices.py | workflow_final_invoices.md |
| batch_generate_ledgers_v3.py | workflow_batch_ledgers.md |
| kakao_order_manager.py | workflow_kakao_orders.md |

## Checks

### 1. Script → Directive Coverage
- [ ] Every active script has a corresponding directive
- [ ] Directive references the correct script filename

### 2. Directive → Script Coverage
- [ ] Every directive's "실행 스크립트" field points to an existing file
- [ ] No directives referencing deleted/renamed scripts

### 3. MEMORY.md 디렉티브 목록
- [ ] All directives listed in MEMORY.md exist
- [ ] No unlisted directives (except new ones awaiting update)

### 4. Directive Freshness
For each directive, check if "학습 기록" section has recent entries:
- [ ] Last entry date within 30 days suggests active maintenance
- [ ] Stale directives (>30 days) flagged for review

### 5. Workflow Pipeline Order
- [ ] workflow_full_process.md references all pipeline steps in correct order
- [ ] Step numbering matches MEMORY.md workflow reminder

## Report Format
```
[directives] Active scripts: 15, Directives: 15
[directives] Coverage: ✅ All scripts have directives
[directives] MEMORY.md: ✅ All directives listed
[directives] Stale: ⚠️ workflow_daily_routine.md (last update >30 days)
```
