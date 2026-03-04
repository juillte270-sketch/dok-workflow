---
name: verify-implementation
description: "Run ALL verify-* skills in parallel to check project-wide consistency. Generates a unified report with issues and recommendations. Use after making changes, before committing, or when something feels broken."
argument_hint: "[optional: specific area to focus on, e.g. 'prices only' or 'mappings']"
---

# Verify Implementation - Unified Project Verification

## Registered Verify Skills

| Skill | Domain | Checks |
|-------|--------|--------|
| verify-mappings | mappings.json | Supplier aliases, inventory maps, page config |
| verify-fixed-prices | fill_prices.py | FIXED_PRICES dict vs MEMORY.md |
| verify-search-config | fill_sikbom_targeted.py | SEARCH_CONFIG + MULTI_ROW_ITEMS vs directive |
| verify-auction-config | fill_auction_from_api.py | ITEM_OVERRIDES vs directive |
| verify-directives | directives/ ↔ execution/ | Script-directive alignment |
| verify-column-refs | All scripts | Master file column index consistency |

## Execution

1. Run ALL registered verify skills **in parallel** using the Task tool (subagent_type=Bash or general-purpose)
2. Each skill reads the relevant files and checks for inconsistencies
3. Collect results from all skills

## Report Format

After all checks complete, output a unified report:

```
=== VERIFY IMPLEMENTATION REPORT ===

[verify-mappings]       ✅ PASS (or ❌ N issues)
[verify-fixed-prices]   ✅ PASS (or ❌ N issues)
[verify-search-config]  ✅ PASS (or ❌ N issues)
[verify-auction-config] ✅ PASS (or ❌ N issues)
[verify-directives]     ✅ PASS (or ❌ N issues)
[verify-column-refs]    ✅ PASS (or ❌ N issues)

--- Issues ---
1. [skill-name] Description of issue
2. [skill-name] Description of issue
...

--- Recommendations ---
- Fix priority items first
- Update directives if code changed
- Update MEMORY.md if new patterns discovered
```

## Rules
- If user specifies a focus area, only run relevant skills
- Always show PASS/FAIL for each skill even if no issues
- For each issue, suggest whether to fix code or update documentation
- Do NOT auto-fix — report only, let user decide
