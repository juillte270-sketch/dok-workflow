---
name: pre-prompting
description: "Before writing any new script, workflow, or complex prompt: define success criteria, set up a testable environment, and draft a good first attempt. Based on Anthropic's official 3 prerequisites."
argument_hint: "[task description, e.g. 'new invoice generation script' or 'fill auction prices']"
---

# Pre-Prompting (프리프롬프팅)

Do these 3 things BEFORE writing code or prompts. From Anthropic's official documentation:
"If these aren't established, we strongly recommend investing time in setting them up first."

## The 3 Prerequisites

### 1. Define Success Criteria (성공 기준 정의)
Answer clearly: "How will I know this worked?"

**For scripts:**
- What are the expected outputs? (file format, location, content)
- What edge cases must be handled?
- What does failure look like? (specific error conditions)

**For workflows:**
- What state should the system be in after completion?
- What should the user see/receive?
- What should NOT happen? (negative criteria)

**Template:**
```
SUCCESS when:
- [ ] {specific measurable outcome 1}
- [ ] {specific measurable outcome 2}
- [ ] {edge case handled}

FAILURE if:
- [ ] {specific failure condition}
```

### 2. Set Up Testable Environment (테스트 가능한 환경)
Make results reproducible and verifiable.

**For scripts:**
- `--dry-run` flag for non-destructive testing
- Sample input data (not production) for first run
- Log output that shows what would happen
- Version the prompt/config so you can compare iterations

**For workflows:**
- Identify which step to test in isolation
- Prepare rollback plan (backup before modifying)
- Define the "minimum reproduction" scope

**Key question:** "Can I run this again with the same inputs and get the same results?"

### 3. Good First Draft (좋은 첫 번째 초안)
A mediocre first attempt wastes more iterations than taking 10 extra minutes upfront.

**Before coding, gather:**
- Existing directive for this workflow (`directives/workflow_*.md`)
- Similar existing scripts in `execution/`
- Known edge cases from MEMORY.md
- Relevant mappings from `mappings.json`

**The draft should include:**
- Clear function/variable names (intention-revealing)
- Handling for the top 3 known edge cases
- Comments on non-obvious business logic
- Connection points to existing infrastructure

## Application to Our Project

| Task | Success Criteria | Test Environment | First Draft Source |
|------|-----------------|------------------|-------------------|
| New script | Expected output file | `--dry-run` + sample data | Existing similar script |
| Bug fix | Specific failing case passes | Isolated test case | Stack trace + directive |
| Directive update | No conflicts with other directives | verify-directives skill | Current directive + learnings |
| Config change | All dependent scripts still work | verify-implementation skill | Current config + new data |

## Output Format
```
[pre-prompting] Task: {description}
[pre-prompting] Success criteria:
  - {criterion 1}
  - {criterion 2}
[pre-prompting] Test environment: {how to verify}
[pre-prompting] First draft basis: {what existing assets to build on}
[pre-prompting] Ready to proceed: ✅ / ❌ {missing prerequisite}
```
