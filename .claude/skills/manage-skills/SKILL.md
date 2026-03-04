---
name: manage-skills
description: "Create, update, or audit verify-* skills. Use after discovering new patterns, fixing bugs, or adding new scripts/directives. Keeps all verify skills in sync with the codebase."
argument_hint: "[action: 'audit' | 'create <name>' | 'update <name>' | 'sync']"
---

# Manage Skills - Verify Skill Lifecycle Manager

## Actions

### audit (default)
Scan the codebase for:
- Changed execution scripts not covered by any verify skill
- New directives without corresponding verify checks
- Verify skills referencing deleted/moved files
- Patterns in MEMORY.md that should be verified automatically

### create <name>
Create a new verify skill:
1. Ask user what domain/checks it should cover
2. Generate `.claude/commands/verify-<name>.md` with proper front matter
3. Add entry to verify-implementation.md registered skills table
4. Front matter description must be under 100 tokens
5. Markdown body must be under 5,000 tokens
6. Reference external files (Level 3) for detailed specs

### update <name>
Update an existing verify skill:
1. Read the current skill file
2. Compare with current codebase state
3. Add new checks, remove obsolete ones
4. Update verify-implementation.md table if needed

### sync
Run audit + auto-fix:
1. Audit all skills
2. For each issue found, propose a fix
3. Ask user for confirmation before applying

## Registered Verify Skills

| Skill | File | Domain |
|-------|------|--------|
| verify-mappings | `.claude/commands/verify-mappings.md` | mappings.json consistency |
| verify-fixed-prices | `.claude/commands/verify-fixed-prices.md` | FIXED_PRICES in fill_prices.py |
| verify-search-config | `.claude/commands/verify-search-config.md` | SEARCH_CONFIG in fill_sikbom_targeted.py |
| verify-auction-config | `.claude/commands/verify-auction-config.md` | ITEM_OVERRIDES in fill_auction_from_api.py |
| verify-directives | `.claude/commands/verify-directives.md` | Script ↔ directive alignment |
| verify-column-refs | `.claude/commands/verify-column-refs.md` | Column index consistency |
| dev-plan | `.claude/skills/dev-plan/SKILL.md` | 4-phase development workflow |
| design-guide | `.claude/skills/design-guide/SKILL.md` | Dashboard UI palette & patterns |

## Skill Creation Rules

### Front Matter (Level 1)
```yaml
---
name: verify-<domain>
description: "One-line description of what this skill verifies. Max 100 tokens."
---
```

### Body (Level 2, max 5000 tokens)
- What to check (checklist format)
- Which files to read
- Expected values / patterns
- How to report issues

### External References (Level 3+)
- Link to directives for detailed specs: `See directives/workflow_<name>.md`
- Link to MEMORY.md for known patterns
- Link to execution scripts for current implementation

## When to Run
- After fixing a bug (new pattern discovered → new check needed)
- After adding a new execution script
- After updating directives
- After a session where user provided feedback on incorrect behavior
