---
description: DoKH 일일 업무 전체 워크플로우 (시간대별 3단계 운영)
---

# DoKH 일일 업무 전체 워크플로우

사용자가 발주리스트/공급처리스트를 전달하면 시간대별로 작업을 수행한다.

## 전체 파이프라인

```
=== 오전 (발주 데이터 수신 시) ===
[1] 배송리스트 생성 + 재고시트 1차/직납 입력
[2] 발주시트 입력 & 가명세서 생성
[3] 단가표 생성 → [3-1] 식봄 가격 입력

=== 오후 10시 (동원 마감 후) ===
[4] 재고시트 2차출고 입력 (HELO)

=== 새벽 2~3시 (경매 완료 후) ===
[5] 단가표 경매가 입력
[6] (수동) 매입가 입력 — 사용자가 직접
[7] 발주시트 매입가/판매단가 입력
[8] 일일 업무 현황 카톡 전송

=== 별도 요청 시 ===
[송부용] 최종 거래명세서 + 거래원장 (함께 생성)
[리포트] 일일/주간/월간 리포트 (마스터 파일 확인 후)
```

## 마스터 파일 경로
```
G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx
```

---

## 오전 작업 (발주 데이터 수신 시)

### [1] 배송리스트 생성 + 재고시트 1차/직납 입력

**목표**: 원본 발주 데이터를 파싱하여 배송리스트 생성 + 재고시트에 1차/직납 수량 입력

**스크립트**:
- `execution/process_orders.py` + `execution/simple_docx_converter.py` (배송리스트)
- `execution/fill_inventory_sheet.py` (재고시트 1차/직납)

```bash
# 배송리스트 생성 (--master-file 전달 시 재고 기반 STOCK/MARKET 자동 분류)
python execution/process_orders.py --input "data/inputs/orders_YYYYMMDD.txt" --date "YYYY-MM-DD" --master-file "G:/.../도크발주관리데이터.xlsx"

# 재고시트 1차/직납 입력
python execution/fill_inventory_sheet.py --date "YYYY-MM-DD"
```

**재고 기반 분류**: `--master-file`을 전달하면 재고시트 현재고(전일)를 조회하여, STOCK 품목의 총 수요량이 현재고를 초과하면 자동으로 MARKET(시장구매)으로 전환한다.

**입력**: `data/inputs/orders_YYYYMMDD.txt` (카카오톡 원본 발주 데이터)
**출력**:
- `data/inputs/processed_orders_YYYYMMDD.txt` (분류된 발주 데이터)
- `data/outputs/배송리스트_YYYYMMDD.docx` + Google Drive 업로드
- 도크발주관리데이터.xlsx 재고시트 1차/직납 컬럼 업데이트

**디렉티브**: `workflow_01_order_processing.md`, `workflow_inventory_shipment.md`

---

### [2] 발주시트 입력 & 가명세서 생성

**목표**: 발주 시트에 데이터 입력 + 매장별 가명세서(가격 미포함) 동시 생성

**스크립트**:
- `execution/update_order_sheet.py` (발주시트)
- `execution/generate_invoices.py` (가명세서)

```bash
# 발주시트 입력
python execution/update_order_sheet.py --date "YYYY-MM-DD"

# 가명세서 생성
python execution/generate_invoices.py --date "YYYY-MM-DD"
```

**입력**: `processed_orders_YYYYMMDD.txt`
**출력**:
- 도크발주관리데이터.xlsx 발주 시트 업데이트
- `data/outputs/invoices_YYYYMMDD/` (매장별 .xlsx)
- `data/outputs/거래명세서_YYYYMMDD.zip`

**디렉티브**: `workflow_02_order_entry.md`, `workflow_generate_invoices.md`

---

### [3] 단가표 생성

**목표**: 당일 발주 품목의 단가 시트 행 생성 (가격 컬럼은 비어있음)

**스크립트**: `execution/create_price_sheet.py`

```bash
python execution/create_price_sheet.py \
  --input "data/inputs/processed_orders_YYYYMMDD.txt" \
  --master "G:/내 드라이브/.../도크발주관리데이터.xlsx" \
  --date "YYYY-MM-DD"
```

**디렉티브**: `workflow_create_price_sheet.md`

### [3-1] 식봄 가격 입력 (Col H)

**목표**: 단가표 생성 직후, 식봄 사이트에서 가격 검색하여 Col H에 입력

**스크립트**: `execution/fill_sikbom_targeted.py`

```bash
python execution/fill_sikbom_targeted.py \
  --master "G:/내 드라이브/.../도크발주관리데이터.xlsx" \
  --date "YYYY-MM-DD"
```

- 식봄(foodspring.co.kr)에서 Selenium으로 가격 검색
- 주요 업체 우선순위: CJ프레시웨이 > 다봄푸드 > 세현F&B
- 품목별 정확한 검색어 매핑 (깻잎→"찹큰", 배추→"배추52" 등)

**디렉티브**: `workflow_sikbom_prices.md`

---

## 오후 10시 작업 (동원 마감 후)

### [4] 재고시트 2차출고 입력

**목표**: 동원 HELO 사이트에서 구매오더 스크래핑 → 재고시트 2차출고 컬럼 입력

**스크립트**: `execution/scrape_helo.py`

```bash
python execution/scrape_helo.py --date "YYYY-MM-DD"
```

- 동원 HELO 사이트 (nexacro 프레임워크) 구매오더 데이터 추출
- 데이터셋 `dsEaiOrder`에서 품목/수량/규격 파싱
- 재고시트 2차출고 컬럼에 자동 입력

**디렉티브**: `workflow_helo_scraping.md`

---

## 새벽 2~3시 작업 (경매 완료 후)

### [5] 단가표 경매가 입력 (Col F, G)

**목표**: 가락시장 경매 최고가/평균가를 단가 시트에 입력

**스크립트**: `execution/fill_auction_from_api.py`

```bash
python execution/fill_auction_from_api.py \
  --master "G:/내 드라이브/.../도크발주관리데이터.xlsx" \
  --date "YYYY-MM-DD"
```

- 가락시장 공공데이터 API에서 경매 최고가/평균가 조회
- 날짜 규칙: 발주일+1일로 API 조회 (경매는 익일 새벽 완료)
- ITEM_OVERRIDES로 품목별 단위 변환
- ITEM_NOTES로 흙 표시 (깐양배추45, 깐양파, 깐대파)

**디렉티브**: `workflow_auction_prices.md`

### [6] (수동) 매입가 입력

사용자가 경매가/식봄가를 참고하여 단가시트 Col I(매입가)를 직접 입력한다.

### [7] 발주시트 매입가/판매단가 입력

**목표**: 단가시트의 매입가를 발주시트에 반영, 판매단가 결정

**스크립트**: `execution/fill_prices.py`

```bash
python execution/fill_prices.py \
  --date "YYYY-MM-DD" \
  --master-file "G:/내 드라이브/.../도크발주관리데이터.xlsx"
```

- 단가시트 매입가(Col I) → 발주시트 반영
- 판매단가는 과거 거래 이력 기반 결정
- 마진율 검증 (0% 미만 재계산, 10% 미만 경고)

**디렉티브**: `workflow_price_filling.md`

### [8] 일일 업무 현황 카톡 전송

**목표**: 당일 업무 현황을 카카오톡으로 전송

**스크립트**: `execution/send_daily_summary.py`

```bash
python execution/send_daily_summary.py --date "YYYY-MM-DD"
```

**디렉티브**: `workflow_daily_summary.md`

---

## 별도 요청 시

### [송부용] 최종 거래명세서 + 거래원장

사용자가 요청하면 **반드시 함께** 생성한다.

**스크립트**:
- `execution/generate_final_invoices.py` (최종 명세서)
- `execution/batch_generate_ledgers_v3.py` (거래원장)

```bash
# 최종 거래명세서 (단가 반영 + Drive 업로드)
python execution/generate_final_invoices.py --date "YYYY-MM-DD"

# 거래원장 업데이트
python execution/batch_generate_ledgers_v3.py \
  --master "G:/내 드라이브/.../도크발주관리데이터.xlsx" \
  --start "MM-01" --end "MM-DD"
```

**주의**: 가격 없이 명세서를 보내면 안됨. 반드시 매입가 입력 완료 후 생성.

**디렉티브**: `workflow_final_invoices.md`, `workflow_batch_ledgers.md`

### [리포트] 일일/주간/월간 리포트

사용자가 요청하면 마스터 파일(도크발주관리데이터.xlsx) 확인 후 생성한다.

---

## 단가 시트 컬럼 구조

| Col | 내용 | 채우는 단계 | 스크립트 |
|-----|------|-----------|---------|
| A | 구분 | [3] | create_price_sheet |
| B | 날짜 | [3] | create_price_sheet |
| C | 품목 | [3] | create_price_sheet |
| D | 품위(등급) | [3] | create_price_sheet |
| E | 단위 | [3] | create_price_sheet |
| **F** | **경매최고가** | **[5] 새벽** | **fill_auction_from_api** |
| **G** | **경매평균가** | **[5] 새벽** | **fill_auction_from_api** |
| **H** | **식봄 가격** | **[3-1] 오전** | **fill_sikbom_targeted** |
| I | 매입가 | [6] 수동 | 사용자 직접 입력 |
| J | 거래 업체 | [6] 수동 | 사용자/fill_prices |

## 요약: 시간대별 실행 순서

```bash
# === 오전 (발주 수신 시) ===

# 1. 배송리스트 생성
python execution/process_orders.py --input "data/inputs/orders_YYYYMMDD.txt" --date "YYYY-MM-DD"
# 1. 재고시트 1차/직납 입력
python execution/fill_inventory_sheet.py --date "YYYY-MM-DD"

# 2. 발주시트 입력
python execution/update_order_sheet.py --date "YYYY-MM-DD"
# 2. 가명세서 생성
python execution/generate_invoices.py --date "YYYY-MM-DD"

# 3. 단가표 생성
python execution/create_price_sheet.py --input "data/inputs/processed_orders_YYYYMMDD.txt" --master "MASTER" --date "YYYY-MM-DD"
# 3-1. 식봄 가격 입력
python execution/fill_sikbom_targeted.py --master "MASTER" --date "YYYY-MM-DD"

# === 오후 10시 (동원 마감 후) ===

# 4. 재고 2차출고 입력
python execution/scrape_helo.py --date "YYYY-MM-DD"

# === 새벽 2~3시 (경매 완료 후) ===

# 5. 경매가 입력
python execution/fill_auction_from_api.py --master "MASTER" --date "YYYY-MM-DD"
# 6. (수동) 매입가 입력 — 사용자
# 7. 매입가/판매단가 반영
python execution/fill_prices.py --date "YYYY-MM-DD" --master-file "MASTER"
# 8. 일일 업무 현황 카톡 전송
python execution/send_daily_summary.py --date "YYYY-MM-DD"

# === 별도 요청 시 ===

# 송부용 거래명세서 + 거래원장
python execution/generate_final_invoices.py --date "YYYY-MM-DD"
python execution/batch_generate_ledgers_v3.py --master "MASTER" --start "MM-01" --end "MM-DD"
```

## 주의사항

1. **날짜 형식**: 반드시 `YYYY-MM-DD`
2. **파일 잠금**: 마스터 파일이 열려있으면 저장 실패 — 닫고 재실행
3. **가격 없이 명세서 전달 금지**: 송부용 명세서는 반드시 매입가 입력 완료 후 생성
4. **송부용 명세서 요청 시 거래원장도 함께**: 절대 누락하지 말 것
5. **식봄가는 오전, 경매가는 새벽**: 식봄가(3-1)가 경매가(5)보다 먼저 입력됨

## 학습 기록
- (2026-02-05) 초기 워크플로우 (1~4단계)
- (2026-02-14) 최종 거래명세서 단계 추가 (generate_final_invoices)
- (2026-02-15) 거래원장 단계 추가 (batch_generate_ledgers_v3)
- (2026-02-24) 단가 시트 가격 채우기 파이프라인 추가 (경매가 → 식봄가 → 매입가 3단계)
- (2026-02-24) 전체 파이프라인 재구성: 시간대별 3단계 운영 (오전/오후10시/새벽), 송부용 명세서+거래원장은 별도 요청 시
