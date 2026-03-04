# 영수증 자동 매입가 입력

## 목적
공급처에서 카카오톡으로 보내는 영수증(URL/사진/텍스트)을 자동 파싱하여
단가시트 Col I(매입가)에 입력한다.

## 전체 흐름

### 방법 1: 자동 감시 데몬 (권장)
```
receipt_watcher.py 상주 실행
  → 트리거 1: 카카오톡 다운로드 폴더에 이미지 생성 감지 (watchdog)
  → 트리거 2: 클립보드에 itanet/marketbom URL 복사 감지 (Win32 API)
  → 트리거 3: 클립보드에 다중행 텍스트 (공급처 키워드 포함) 복사 감지
  → parse_receipt.py: 유형 자동 감지 → 공급처 식별 → 품목+단가 추출
  → fill_receipt_prices.py: 단가시트 행 매칭 → Col I 입력
  → Windows 토스트 알림 + 카카오톡 나에게보내기
```

**사용자 동작**: 이미지 클릭(확인 차) 또는 URL 복사 — 1초 소요

### 방법 2: 수동 CLI
```
공급처 영수증 수신 (카톡)
  → 사용자가 챗봇에 전달 (URL/이미지/텍스트)
  → parse_receipt.py: 유형 자동 감지 → 공급처 식별 → 품목+단가 추출
  → fill_receipt_prices.py: 단가시트 행 매칭 → Col I 입력
  → 운영자에게 확인 요약 카톡 전송
```

## 스크립트

| 스크립트 | 역할 |
|---------|------|
| `execution/receipt_watcher.py` | **자동 감시 데몬** (폴더+클립보드 감시, 시스템 트레이) |
| `execution/parse_receipt.py` | 영수증 파싱 엔진 (URL/이미지/텍스트) |
| `execution/fill_receipt_prices.py` | 파싱 결과 → 단가시트 매입가 입력 |
| `execution/kakao_order_server.py` | 챗봇 연동 (POST /kakao/receipt) |

## 공급처별 유형 (Phase별)

### Phase 1: ERP URL
- **건영농산**: itanet.co.kr — HTML 테이블 스크래핑
- **오복상회**: marketbom.com — HTML 테이블 스크래핑

### Phase 2: POS 이미지
- **POS 인쇄**: 가야웰빙, 경북상회, 명진농산, 세연농산, 우신농산, 풍경농산, 다모아버섯
- **POS 화면**: 나물향기, 정운, 초원농산, 현진상회 (현진상회: 단가×10)

### Phase 3: 손글씨/텍스트
- **손글씨**: 영운농산, 명화농산, 태현상회, 이모유통 (확인 필수, 발주리스트 교차매칭 자동 적용)
- **카톡 텍스트**: 명화농산 (대안)

### 대상 제외
- **정복상회**: 월말 정산 (월 1회)

## CLI 사용법

### 영수증 파싱만
```bash
# URL 파싱
python execution/parse_receipt.py --url "https://www.itanet.co.kr/..." --map

# 이미지 파싱
python execution/parse_receipt.py --image receipt.jpg --supplier 경북상회 --map

# 텍스트 파싱
python execution/parse_receipt.py --text "감자 62000\n당근 35000" --supplier 명화농산 --map
```

### 파싱 + 단가시트 입력
```bash
# URL에서 직접
python execution/fill_receipt_prices.py \
  --master "G:/내 드라이브/.../도크발주관리데이터.xlsx" \
  --date 2026-03-01 \
  --url "https://www.itanet.co.kr/..."

# dry-run 미리보기
python execution/fill_receipt_prices.py \
  --master "..." --date 2026-03-01 --url "..." --dry-run

# JSON으로 다건 입력
python execution/fill_receipt_prices.py \
  --master "..." --date 2026-03-01 \
  --receipt-json '[{"supplier":"건영농산","items":[{"name":"감자/왕왕","unit_price":62000}]}]'
```

## 리소스 파일
- `skills/order_processing/resources/receipt_item_map.json`: 공급처별 품목 매핑
  - `supplier_detection`: URL 패턴 + 텍스트 키워드 → 공급처 자동 식별
  - `supplier_receipt_type`: 공급처별 영수증 유형
  - `supplier_price_multiplier`: 가격 배수 (현진상회: ×10)
  - `item_map`: 공급처별 영수증 품목명 → 단가시트 품목명 변환

## 발주리스트 교차매칭 (손글씨 영수증 보정)

**문제**: 손글씨 영수증의 OCR은 품목명이 부정확하다 (예: "슈퍼(L)하스" = 귤).
**해결**: 공급처 발주리스트의 **품목 순서**와 영수증의 **품목 순서**가 동일한 점을 활용.

```
processed_orders_{date}.txt에서 공급처별 발주 품목 로드
  → 발주리스트 품목명을 item_map으로 단가시트명에 매핑
  → OCR 결과와 순서 기반으로 1:1 매칭
  → OCR 가격 + 발주리스트 품목명 = 정확한 입력
```

- **자동 적용 조건**: `supplier_receipt_type`이 `handwritten`인 공급처
- **대상**: 영운농산, 명화농산, 태현상회, 이모유통
- **함수**: `parse_receipt.py` > `load_supplier_order_items()`, `map_to_danga_items(order_items=...)`
- 발주 품목 수 <= OCR 품목 수일 때만 활성화 (초과 OCR 라인은 NOISE로 표시)
- 발주리스트에 없는 추가 라인은 기존 OCR 기반 매핑으로 폴백

## 안전장치
1. **FIXED_PRICES 보호**: 숙주/콩나물/계란/김치 등 고정가 품목은 절대 덮어쓰지 않음
2. **가격 이상치 검증**: 경매가/식봄가 대비 +-50% 벗어나면 경고 (입력은 진행)
3. **손글씨 확인 필수**: Phase 3 결과는 confidence=low, 사용자 확인 권장
4. **발주리스트 교차매칭**: 손글씨 공급처는 자동으로 발주리스트와 순서 대조
5. **dry-run 모드**: `--dry-run`으로 실제 입력 없이 결과만 확인

## 챗봇 연동

### 엔드포인트
- `POST /kakao/receipt`: 카카오 오픈빌더 영수증 스킬
- `POST /receipt/parse`: REST API (파싱만)
- `POST /receipt/fill`: REST API (파싱 + 입력)

### 플로우
1. 사용자가 "영수증" 스킬 블록에서 URL/이미지/텍스트 전송
2. 서버가 백그라운드에서 파싱 + 입력
3. 사용자가 "확인" 재전송 시 결과 반환
4. 운영자에게 카카오톡 나에게보내기로 요약 알림

## 학습/개선
- 새 공급처 추가 시: `receipt_item_map.json`에 매핑 추가
- 품목명 불일치 발견 시: `item_map`에 추가
- 가격 배수 필요 시: `supplier_price_multiplier`에 추가

## 자동 감시 데몬 (receipt_watcher.py)

### 실행 방법
```bash
# 기본 실행 (시스템 트레이 상주)
python execution/receipt_watcher.py

# 옵션
python execution/receipt_watcher.py --dry-run          # 파일 변경 없이 결과만
python execution/receipt_watcher.py --console           # 트레이 없이 콘솔 모드
python execution/receipt_watcher.py --date 2026-03-01   # 특정 날짜
python execution/receipt_watcher.py --watch-dir "D:\카카오톡"  # 폴더 경로 지정
```

### 트리거별 커버리지

| 유형 | 트리거 | 사용자 동작 | 대상 공급처 |
|------|--------|-----------|-----------|
| 이미지 (POS/손글씨) | 폴더 감시 (watchdog) | 카톡에서 이미지 클릭 | ~15개 전체 |
| URL (ERP) | 클립보드 감시 (Win32) | URL 복사 | 건영농산, 오복상회 |
| 텍스트 | 클립보드 감시 (Win32) | 텍스트 복사 | 명화농산 등 |

### 카카오톡 다운로드 폴더
자동 탐지 순서:
1. `%USERPROFILE%\Documents\카카오톡 받은 파일\`
2. `%USERPROFILE%\Documents\KakaoTalk Downloads\`
3. `.env`의 `KAKAO_DOWNLOAD_DIR`
4. `--watch-dir` CLI 옵션

### 중복 방지
- 파일: MD5 해시 비교
- URL/텍스트: 문자열 해시 비교
- 세션 내 `processed_hashes` set으로 관리

### 시스템 트레이 메뉴
- 상태 보기 (오늘 처리 건수)
- 처리 내역 보기
- Dry-run 토글
- 종료

### 필요 패키지
```bash
pip install watchdog pywin32 pystray Pillow winotify
```

## 주의사항
- **MASTER_FILE_PATH**: .env에 마스터 파일 경로 설정 필요
- **GEMINI_API_KEY**: 이미지 OCR에 필요 (.env)
- **KAKAO_DOWNLOAD_DIR**: 비워두면 자동 탐지
- Phase 1(URL)은 Gemini 불필요, Phase 2/3에서만 사용
- 단가시트 첫 번째 연속 블록만 대상 (기존 패턴 동일)
- 최근 5분 이내 이미지만 처리 (오래된 파일 무시)
- 손글씨 영수증(confidence=low): 토스트에 "확인 필요" 경고 표시
