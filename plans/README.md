# plans/

시스템 개발 시 Research → Plan → Annotate → Implement 프로세스의 산출물 저장소.

## 파일명 규칙

```
{YYYYMMDD}_{task_name}_research.md   ← 리서치 결과
{YYYYMMDD}_{task_name}_plan.md       ← 구현 계획
```

## 디렉토리 구조

```
plans/
  _templates/          ← 템플릿 (복사해서 사용)
    research_template.md
    plan_template.md
  README.md            ← 이 파일
  20260227_xxx_research.md  ← 실제 산출물
  20260227_xxx_plan.md
```

## 검토 마커

plan.md에 인라인 메모를 추가할 때 사용:

| 마커 | 의미 | 예시 |
|------|------|------|
| `[OK]` | 승인 | `[OK] Step 1 좋습니다` |
| `[수정]` | 수정 필요 | `[수정] 여기는 dict 대신 list 사용` |
| `[질문]` | 확인 필요 | `[질문] 이 API는 rate limit 있나요?` |
| `[삭제]` | 제거 | `[삭제] 이 단계 불필요` |
| `[추가]` | 추가 필요 | `[추가] 에러 핸들링 빠졌음` |

## 언제 사용?

- 새 스크립트 작성 / 기존 로직 수정
- 새 디렉티브 / 스킬 생성
- mappings.json 구조 변경
- 복잡한 버그 수정 (2개+ 파일 연쇄 변경)

일상 운영(발주 처리, 기존 스크립트 실행 등)에는 불필요.

상세: `directives/workflow_dev_process.md`
