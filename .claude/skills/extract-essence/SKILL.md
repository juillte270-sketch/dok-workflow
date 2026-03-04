---
name: extract-essence
description: "Extract the 'essence' (핵심 패턴) from a successful workflow run, script fix, or repeated pattern. Produces a reusable asset: skill, directive update, or memory entry."
argument_hint: "[what to extract from, e.g. 'today's price filling' or 'the 배추 bug fix']"
---

# Extract Essence (에센스 추출)

Turn experience into reusable assets. Inspired by 배달의민족 design assetization:
strip away noise, keep only the core pattern that produces consistent results.

## When to Use
- After fixing a tricky bug (the fix pattern may recur)
- After a workflow runs successfully with new edge cases handled
- When you notice the same manual step repeated across sessions
- When a user correction reveals a pattern not yet documented

## Process

### Step 1: Identify the Raw Material
Read the relevant files/logs to understand what happened:
- What was the problem or task?
- What was the solution or process?
- What made it work (or fail initially)?

### Step 2: Strip to Essence
Remove context-specific details. Ask:
- Would this apply to other items/stores/dates?
- What's the minimum information needed to reproduce this result?
- What's the "sculpy" — the core shape before colors/details?

### Step 3: Choose Asset Type
| Signal | Asset Type | Location |
|--------|-----------|----------|
| Recurring validation need | Verify Skill | `.claude/skills/verify-*/SKILL.md` |
| New workflow step discovered | Directive update | `directives/workflow_*.md` |
| Bug fix pattern | Memory entry | `MEMORY.md` or topic file |
| New mapping/config | Config update | `mappings.json` or script constant |
| Reusable multi-step process | New Skill | `.claude/skills/*/SKILL.md` |

### Step 4: Write the Asset
- Use intention-revealing names (Ward Cunningham 3-step)
- Include the "why" not just the "what"
- Add examples from the actual case
- Keep it concise — essence, not encyclopedia

### Step 5: Verify Integration
- Does the new asset conflict with existing ones?
- Is it referenced from MEMORY.md or relevant directive?
- Would a fresh session find and use this asset?

## Output Format
```
[extract-essence] Source: {what was analyzed}
[extract-essence] Essence: {1-2 sentence core pattern}
[extract-essence] Asset type: {skill/directive/memory/config}
[extract-essence] Location: {file path}
[extract-essence] Action: Created / Updated {description}
```

## Anti-patterns
- Don't assetize one-off fixes that won't recur
- Don't duplicate what's already in a directive
- Don't over-abstract — keep the concrete example alongside the pattern
- Don't save session-specific state as permanent memory
