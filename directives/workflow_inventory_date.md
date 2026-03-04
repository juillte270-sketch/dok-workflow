# 재고시트 날짜 그룹 관리 워크플로우

## 개요
재고 시트에 새 날짜 컬럼 그룹(6컬럼)을 생성하고,
mappings.json의 재고 보충 규칙에 따라 입고량을 자동 계산한다.

## 실행 스크립트
`execution/manage_inventory_date.py`

## 사용법
```bash
python execution/manage_inventory_date.py \
  --prev-date 2026-02-23 --new-date 2026-02-24 \
  --master "G:/내 드라이브/.../도크발주관리데이터.xlsx"

# 옵션
python execution/manage_inventory_date.py --prev-date 2026-02-23 --new-date 2026-02-24 --dry-run
python execution/manage_inventory_date.py --prev-date 2026-02-23 --new-date 2026-02-24 --skip-copy    # 이미 존재할 때
python execution/manage_inventory_date.py --prev-date 2026-02-23 --new-date 2026-02-24 --skip-incoming # 입고계산 생략
```

### 파라미터
| 파라미터 | 필수 | 설명 |
|---------|------|------|
| `--prev-date` | O | 이전 날짜 (YYYY-MM-DD) |
| `--new-date` | O | 새 날짜 (YYYY-MM-DD) |
| `--master` | X | 마스터 파일 경로 |
| `--dry-run` | X | 미리보기 |
| `--skip-copy` | X | 날짜 그룹 복사 생략 |
| `--skip-incoming` | X | 입고량 계산 생략 |

## 재고시트 6컬럼 그룹 구조
| 오프셋 | 내용 | 수식/값 |
|--------|------|--------|
| +0 | 품목 | 이전 그룹에서 복사 |
| +1 | 재고 | `=이전_현재고` (수식) |
| +2 | 1차/직납 | 빈칸 (fill_inventory_sheet.py가 채움) |
| +3 | 2차출고 | 빈칸 (scrape_helo.py가 채움) |
| +4 | 입고 | 자동 계산 (이 스크립트) |
| +5 | 현재고 | `=재고+입고-2차출고-1차/직납` (수식) |

## 데이터 행 범위
| 섹션 | 행 범위 |
|------|--------|
| 채소 | R4-R31 |
| 과일 서브헤더 | R33 |
| 과일 | R34-R36 |
| 공산품 서브헤더 | R38 |
| 공산품 | R39-R43 |

## Step A: 날짜 그룹 복사
- 이전 날짜 그룹 옆(+6컬럼)에 새 그룹 생성
- 헤더 "M/DD 재고현황", 서브헤더, 섹션 구분 행 포맷 모두 복사
- 이미 존재하면 복사 건너뜀

## Step B: 입고량 계산
`mappings.json` → `inventory_reorder_rules`의 규칙 타입별 처리:

| 타입 | 로직 | 예시 |
|------|------|------|
| `range` | 재고 < min이면 max-재고 입고 | 깐양파: min=3, max=7 |
| `threshold` | 재고 ≤ threshold이면 order_qty 입고 | 김치: threshold=3, order=4 |
| `tiered` | threshold 구간별 다른 order_qty | 배추: ≤5→10, ≤10→5 |
| `target` | 재고 ≤ threshold이면 target-재고 입고 | 감자: target=30 |
| `target_unit` | target까지 unit 단위 올림 | 고추: target=10, unit=0.5 |
| `special_box` | 재고<0이면 box 2차출고+입고 | 팽이: per_box=34 |
| `skip` | 무시 (깐양파1kg 등) | |
| `manual` | 수동 판단 필요 | 참나물 |

### 요일별 규칙
일~수 vs 목~금으로 min/max가 다른 품목 지원 (`min_weekday`, `max_weekday`):
```json
{"min_weekday": {"sun_wed": 3, "thu_fri": 4}}
```

### 공산품 알림 (inventory_alert_rules)
쌀, 마카로니, 아카시아청, 고춧가루 등 재고 부족 시 알림 메시지 생성

## 전체 파이프라인에서의 위치
재고 관련 보조 스크립트. 새 날짜의 재고 컬럼 생성 시 사용.
```
manage_inventory_date.py (날짜 그룹 생성 + 입고 계산)
  → fill_inventory_sheet.py (1차/직납 입력) [1단계]
  → scrape_helo.py (2차출고 입력) [4단계]
```
