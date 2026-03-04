---
name: assetize
description: "Meta-skill: turn experiences, fixes, and patterns into reusable project assets. Orchestrates extract-essence, updates directives/memory, and maintains the asset pipeline. Use after completing a task or at end of day."
argument_hint: "[optional: specific session or date to review, e.g. '2026-02-25 session' or 'today's bug fixes']"
---

# Assetize (자산화)

The master process for turning work into lasting value.
"90% of skills became $0. The remaining 10% became 1000x." — Kent Beck

## Philosophy

Our 3-layer architecture IS an assetization system:
- **Directives** = assetized workflows (에센스 of how we work)
- **Skills** = assetized patterns (에센스 of how we solve problems)
- **Scripts** = assetized execution (에센스 of deterministic work)
- **MEMORY.md** = assetized experience (에센스 of what we learned)

The goal: every session leaves the system stronger than before.

## When to Assetize

| Trigger | Action |
|---------|--------|
| Bug fixed | Run debug-5step → extract-essence → update directive + memory |
| New edge case handled | Update relevant directive's edge case section |
| User correction received | Update mappings/config + add to learning records |
| Workflow ran successfully with changes | Verify directive matches reality |
| End of work session | Review what was learned, update assets |
| New script created | Create/update corresponding directive |

## The Assetization Pipeline

### 1. Collect Raw Material
Review the session:
- What scripts were modified?
- What bugs were fixed?
- What user corrections were received?
- What new patterns were discovered?

### 2. Filter: Is This Worth Assetizing?

**YES — Assetize if:**
- Pattern will recur (same bug class, similar items)
- Multiple people/sessions would benefit
- Configuration that could drift
- Business rule that's not obvious from code

**NO — Don't assetize if:**
- One-time fix for unique situation
- Already documented elsewhere
- Speculative/unverified (wait for confirmation)
- Session-specific temporary state

### 3. Route to Correct Asset Type

```
Pattern recurs across scripts     → Skill (.claude/skills/)
Workflow step changed             → Directive (directives/)
Config/mapping value changed      → Config (mappings.json / script constants)
Bug fix pattern for future ref    → Memory (MEMORY.md)
Validation rule discovered        → Verify Skill (.claude/skills/verify-*)
```

### 4. Write with Essence

Apply the essence principle:
- Strip noise, keep core pattern
- Use intention-revealing names (naming-3step)
- Include one concrete example from the actual case
- Keep concise — shorter assets get used more

### 5. Connect to Existing Assets

- Link from MEMORY.md if relevant
- Update related directives
- Check for conflicts with existing assets
- Ensure a fresh session would find this asset

## Kent Beck's 1000x Framework Applied

### Ambitious Vision (야심찬 비전)
Our system should eventually handle the full daily pipeline with minimal intervention:
orders → classification → delivery list → order sheet → price sheet → invoices → reports

### Milestones (이정표)
Each workflow step becoming more automated is a milestone.
Track in `directives/workflow_full_process.md`.

### Complexity Management (복잡성 관리)
- Each script handles ONE concern
- Directives define the connections between scripts
- Skills handle cross-cutting patterns (debugging, verification)
- Self-annealing: errors make the system stronger

## Output Format
```
=== ASSETIZE SESSION REPORT ===

Session: {date/description}

Assets created/updated:
1. [directive] {path} — {what changed}
2. [memory] MEMORY.md — {what added}
3. [config] mappings.json — {what updated}
4. [skill] {path} — {new skill created}

Pending (needs user confirmation):
- {item requiring user decision}

System improvement: {1-sentence summary of how system got stronger}
```

## The Virtuous Cycle
```
Work → Learn → Assetize → Better Work → Learn More → Assetize More
  ↑                                                        ↓
  └────────────── System gets stronger ←──────────────────┘
```

This IS the self-annealing loop from CLAUDE.md, formalized.
