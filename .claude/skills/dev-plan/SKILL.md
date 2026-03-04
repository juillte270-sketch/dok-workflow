---
name: dev-plan
description: "4-phase development workflow: Research → Plan → Annotate → Implement. Creates plan files in plans/ for user review before any code changes."
argument_hint: "[task description or 'review' to check pending plans]"
---

# Dev Plan - 시스템 개발 4단계 프로세스

## Quick Start

사용자가 `/dev-plan {작업 설명}` 호출 시:

1. 적용 대상인지 확인 (일상 운영이면 "이 작업은 기존 워크플로우로 처리 가능합니다" 안내)
2. Phase 1 (Research) 시작

`/dev-plan review` 호출 시:
- 기존 plan.md에서 `[수정]`/`[질문]` 마커 검색
- 마커별 대응 후 plan.md 업데이트

## Phase 체크리스트

### Phase 1: Research
- [ ] 관련 execution 스크립트 읽기
- [ ] 관련 directives 읽기
- [ ] MEMORY.md에서 관련 항목 확인
- [ ] mappings.json 관련 부분 확인 (해당 시)
- [ ] `plans/{date}_{name}_research.md` 작성
- [ ] 열린 질문 → 사용자에게 확인

### Phase 2: Plan
- [ ] 성공 기준 정의 (검증 가능)
- [ ] Step별 구현 계획 작성 (변경 전/후 코드 포함)
- [ ] 영향 범위 파일 목록 작성
- [ ] 테스트 계획 작성
- [ ] `plans/{date}_{name}_plan.md` 작성
- [ ] 사용자에게 "plan.md 검토해 주세요" 요청

### Phase 3: Annotate
- [ ] plan.md에서 검토 마커 검색
- [ ] `[수정]` 마커 → 계획 수정 후 업데이트
- [ ] `[질문]` 마커 → 답변 확인 후 업데이트
- [ ] `[삭제]` 마커 → 해당 항목 제거
- [ ] `[추가]` 마커 → 새 항목 추가
- [ ] 모든 마커 해소 확인
- [ ] 사용자 "구현해" 대기

### Phase 4: Implement
- [ ] plan.md Step 순서대로 실행
- [ ] 각 Step 완료 후 결과 확인
- [ ] 계획에 없는 변경 필요 시 → 즉시 중단, Phase 3 복귀
- [ ] 테스트 실행
- [ ] 성공 기준 체크

## 핵심 규칙

1. **코드 작성 금지**: `[수정]`/`[질문]` 마커가 하나라도 남아있으면 구현하지 않는다
2. **승인 필수**: 사용자가 "구현해" (또는 동등한 승인)를 명시적으로 말해야 Phase 4 진행
3. **범위 제한**: plan.md에 명시된 파일과 변경만 적용. 추가 변경은 중단 후 승인
4. **되돌리기 우선**: 잘못된 방향이면 패치하지 말고 git reset 후 범위 축소

## Output Format

각 Phase 완료 시 사용자에게 보고:

```
## Phase {N} 완료: {Phase명}

**산출물**: `plans/{파일명}`
**다음 단계**: {Phase N+1 안내 또는 "구현해"를 기다립니다}
```

## 관련 참고

- 상세 절차: `directives/workflow_dev_process.md`
- 템플릿: `plans/_templates/`
- 검토 마커 설명: `plans/README.md`
