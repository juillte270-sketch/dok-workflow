---
name: debug-5step
description: "Structured 5-step debugging methodology. Prevents emotional debugging ('just fix it') by following a systematic process: define → correct behavior → reproduce → hypothesize → verify."
argument_hint: "[bug description, e.g. '배추 price shows 15000 instead of 12000']"
---

# Debug 5-Step (디버깅 5단계)

When a bug occurs, follow these 5 steps IN ORDER.
Do not skip steps. Do not jump to fixing before Step 4.

Based on: Kent Beck's TDD philosophy + 휘동's debugging interviews + 배민 자산화 approach.

## Step 1: Define the Problem in One Sentence (문제 한 문장 정의)

Bad: "가격이 이상해요"
Good: "배추 매입가가 12,000원이어야 하는데 15,000원으로 입력됨"

**Template:** `{what} should be {expected} but is {actual}`

Why this matters:
- Forces clarity on what's actually wrong
- Prevents solving the wrong problem
- Gives LLM a precise target

## Step 2: Define Correct Behavior — Given-When-Then (올바른 동작 정의)

```
Given: 단가시트에 배추 12,000원/통, 양배추 15,000원/개 존재
When:  fill_prices.py가 발주시트의 "배추" 행을 처리할 때
Then:  매입가 컬럼에 12,000 입력 (양배추 15,000이 아님)
```

Why this matters:
- Establishes the "lamp" (Kent Beck) — immutable test criteria
- Both human and AI agree on what "fixed" means
- Prevents the AI genie from "creatively" reinterpreting the goal

## Step 3: Build Minimal Reproduction (최소 재현환경 구축)

Reduce scope to isolate the bug:

**For our project:**
- Which script? (`execution/*.py`)
- Which function? (narrow to specific function)
- Which input? (specific row/item/date)
- Can we test with `--dry-run` or print statements?

**Actions:**
1. Read the script's relevant function
2. Trace the data flow for the specific failing case
3. Add targeted print/log if needed (not shotgun debugging)
4. If complex, create a minimal test script

**Key question:** "What is the SMALLEST piece of code that still shows the bug?"

## Step 4: List Cause Hypotheses (원인 가설 나열)

Generate multiple hypotheses BEFORE fixing anything:

```
Hypothesis 1: Partial matching picks 양배추 before 배추 (iteration order)
Hypothesis 2: Exact match fails due to unit mismatch (kg vs 통)
Hypothesis 3: price_map doesn't contain 배추 entry at all
Hypothesis 4: Row index offset causes wrong item lookup
```

**Sources for hypotheses:**
- Error message / stack trace
- MEMORY.md bug fix patterns (similar past bugs)
- Script logic review (Step 3)
- Ask: "What assumption could be wrong?"

**Rules:**
- List at least 3 hypotheses
- Order by likelihood (most likely first)
- Include at least one "unlikely but possible" hypothesis
- AI-generated hypotheses + your domain knowledge

## Step 5: Verify Hypotheses One by One (가설 하나씩 검증)

Test each hypothesis independently:

```
[H1] Partial matching order → Add print(price_item) in matching loop
     Result: ✅ Confirmed — '깐양배추45' matched before '배추'
[H2] Unit mismatch → Check _best_from_entries('배추', ...)
     Result: ✅ Also confirmed — kg vs 통 returns 0
[H3] price_map missing → print('배추' in price_map)
     Result: ❌ Ruled out — 배추 exists
```

**After finding root cause:**
1. Fix the code
2. Verify the Given-When-Then from Step 2 passes
3. Check for similar patterns elsewhere (same bug class)
4. Run extract-essence skill to assetize the learning

## Output Format
```
[debug-5step] Problem: {one sentence}
[debug-5step] Correct: Given {X} When {Y} Then {Z}
[debug-5step] Scope: {script}:{function} with input {item}
[debug-5step] Hypotheses:
  H1: {description} → {confirmed/ruled out}
  H2: {description} → {confirmed/ruled out}
  H3: {description} → {confirmed/ruled out}
[debug-5step] Root cause: {H#} — {explanation}
[debug-5step] Fix: {what was changed}
[debug-5step] Assetize: {directive/memory update if pattern is reusable}
```

## Self-Annealing Connection
After fixing, this feeds back into the system:
- Step 5 fix → update script (Layer 3)
- Root cause pattern → update directive (Layer 1)
- Reusable insight → update MEMORY.md or create verify skill
- This IS the self-annealing loop from CLAUDE.md
