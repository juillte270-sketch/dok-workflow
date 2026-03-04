# 프로젝트 작업 내역

## 1. 환경 설정 및 기초 작업
- [x] 프로젝트 디렉토리 구조 생성 (`directives`, `execution`, `.tmp`)
- [x] 필수 파일 생성 (`.env`, `.gitignore`)
- [x] 에이전트 지침서(`GEMINI.md` 등) 미러링
- [x] Git 저장소 초기화 및 커밋

---

## 2. 쓰레드(Threads) 리포트 워크플로우 ✅ 완료
Threads 플랫폼 게시물 자동 수집 → AI 요약 → 이메일/카카오톡 발송 자동화

- [x] Threads 데이터 수집 (Playwright)
- [x] AI 요약 (Gemini 1.5 Flash)
- [x] 이메일 발송 (Gmail API)
- [x] 카카오톡 발송 (카카오 API)
- [x] 자동화 설정 (매일 19:00 실행)

**📄 상세 내역**: [TASKS_THREADS.md](TASKS_THREADS.md)

---

## 3. DoKH 일일 업무 자동화

### 📋 완료된 워크플로우 (1~3단계 + 가명세서) - 일~금 저녁 9시 실행

#### 1단계: 발주 처리 및 배송리스트 생성
- [x] `directives/workflow_01_order_processing.md` 지침서
- [x] `execution/process_orders.py` 구현
    - [x] 발주 데이터 파싱 및 분류
    - [x] 공급처 수기 입력 우선 적용
    - [x] 매장별/공급업체별 자동 분류
    - [x] 재고/소분, 시장구매 규칙 적용
- [x] `execution/simple_docx_converter.py` 구현
    - [x] 배송리스트 DOCX 생성
    - [x] 폰트 11pt 조정, Argparse 적용

#### 2단계: 발주 시트 입력
- [x] `directives/workflow_02_order_entry.md` 지침서
- [x] `execution/update_order_sheet.py` 구현
    - [x] Google Drive 마스터 파일 연동
    - [x] 자동 데이터 입력 및 업데이트
    - [x] 매장별 그룹화 및 정렬
    - [x] 이익률 자동 계산 공식 적용
    - [x] 안전 업데이트 로직

#### 3단계: 단가 시트 작업
- **3-1단계: 단가 시트 생성 (매입단가용)**
    - [x] `directives/workflow_03_price_sheet.md` 지침서 (V7 알고리즘)
    - [x] `execution/create_price_sheet.py` 구현
        - [x] 기존 날짜 데이터 자동 삭제 후 재생성
        - [x] 단가양식 템플릿 복사
        - [x] 공급업체 별칭 매칭 (경향농산↔다모아버섯 등)
        - [x] 양방향 품목 매칭 (대파↔깐대파 등)
        - [x] 중복 품목 처리 (주문 횟수만큼 행 유지)
        - [x] 보호 섹션 유지 (소분/재고, 동원1배치)
        - [x] 발주 리스트 기반 필터링
        - [x] 공급업체 병합 해제 및 스타일 유지
        - [x] 누락 품목 검증 리포트

- **3-2단계: 가명세서 생성 (단가 입력 전)**
    - [x] `execution/generate_invoices.py` 구현
    - [x] **0원 단가**로 품목/수량만 채워진 엑셀 파일 생성
    - [x] 생성된 파일은 `data/outputs/거래명세서_YYYYMMDD.zip`으로 제공
    - [x] 사용자가 단가 확인 및 입력을 위한 기초 자료로 활용

### ✅ [3.5단계] 매입/판매단가 자동 입력 (신규 완료)
- [x] `directives/workflow_price_filling.md` 지침서
- [x] `execution/fill_prices.py` 구현 (652줄)
    - [x] 단가시트에서 매입가 자동 조회 및 발주시트 입력
    - [x] 단위 변환 로직 (`parse_price_unit`) - kg/박스/팩/알 등
    - [x] 소분/재고 vs 직납/재고 가격 자동 분류
    - [x] 공급업체 자동 결정 (`item_supplier_map` 참조)
    - [x] 판매단가 자동 결정 (거래처별 차등 로직)
        - [x] 샤브야키(동탄 제외): 통일가 (최빈값)
        - [x] 샤브야키 동탄: 최근 거래가 또는 신규(25% 마진)
        - [x] 부엉이산장: 통일가 (최빈값)
        - [x] 기타 거래처: 최근 거래가 또는 신규(20% 마진)
    - [x] 수식 자동 입력 (J=총매입, L=총판매, M=마진율, N=이익률)
    - [x] 마진율 검증 (음수→재계산, 저마진/고마진 경고)
    - [x] 퍼지 매칭 (`normalize_item`) 지원

### ✅ [4단계] 송부용 거래명세서 (Excel/PDF) 생성 (별도 요청 시 실행)
- **최종 단가 확인 후 별도 실행 요청**
- [x] **가이드 기반 생성 방식** 구현 (`generate_invoices_from_guide.py`)
  - [x] 마스터 데이터(`도크발주관리데이터.xlsx`)에서 직접 품목/수량/판매가 로드
  - [x] 가명세서를 템플릿으로 활용 (서식 보존)
  - [x] 당월금액, 당일금액, 총사용액 자동 계산
  - [x] 병합셀 안전 처리 (`safe_set_value` 함수)
- [x] **최종 명세서 생성** 구현 (`generate_final_invoices.py`, 509줄)
  - [x] Fuzzy Matching 기반 품목 매칭 (`normalize_key`)
  - [x] 마스터 파일 판매가 자동 조회 및 입력
  - [x] 당월/당일/총사용액 자동 계산
  - [x] 검증 리포트 자동 생성 (`ValidationReport` 클래스)
  - [x] `directives/workflow_final_invoices.md` 지침서
- [x] **Excel → PDF 자동 변환** (win32com)
- [x] **Google Drive 자동 업로드**
  - [x] 매장별 폴더 매핑 설정 파일 (`resources/drive_folder_mapping.json`)
  - [x] 월별 폴더 자동 생성 (`2026년 2월`)
- [x] **안전 규칙 적용**
  - [x] ⚠️ 드라이브 파일/폴더 자동 삭제 금지
  - [x] 폴더 구조 변경 시 사용자 확인 필수
  - [x] 덮어쓰기 방식으로 업데이트 (삭제 X)

### ✅ [4-1단계] 거래원장 자동 입력 (구현 완료)
- [x] `execution/update_transaction_ledger.py` 구현 (194줄)
    - [x] 거래원장 템플릿 복사 및 월별 폴더 생성
    - [x] 마스터 데이터(`총 판매금액`) 집계 및 자동 입력
    - [x] 상/하단 합계 금액 자동 계산
    - [x] 날짜 범위 (start_date ~ end_date) 기반 처리
    - [x] argparse 기반 CLI

### 🔗 통합 워크플로우 (1~3단계 자동화)
- [x] `execution/run_full_workflow.py` - **1~3-2단계** 순차 실행
    - [x] `directives/workflow_daily_routine.md` 문서화
    - [x] `directives/workflow_full_process.md` 전체 프로세스 가이드
    - [x] 순차 실행 자동화 (1단계 → 2단계 → 3-1단계 → 3-2단계)
    - [x] 오류 처리 및 진행 상황 표시
    - [x] **자동 학습 기능 통합** (입력↔출력 비교)
    - [x] `--skip-invoice`, `--skip-learning` 옵션 지원
- [ ] 3.5단계(fill_prices) 통합 워크플로우 연동
- [ ] 4단계 + 4-1단계 통합 실행 스크립트 분리

### 🤖 AI 학습 시스템
- [x] **매핑 학습 엔진**
    - [x] `execution/learn_mappings.py` - 매핑 자동 학습
    - [x] `directives/mapping_learning_guide.md` - 학습 가이드
    - [x] 품목-공급업체 매핑 자동 학습
    - [x] 재고/소분 규칙 자동 학습
    - [x] 시장구매 규칙 자동 학습
- [x] **분류 및 학습 통합**
    - [x] `execution/classify_and_learn.py` - 분류+학습 통합
    - [x] `directives/classify_and_learn_workflow.md` - 워크플로우 가이드
    - [x] 대화형 모드 지원 (복사-붙여넣기)
    - [x] 자동 학습 사이클 구축
- [x] **전체 워크플로우 자동 학습**
    - [x] `run_full_workflow.py`에 자동 학습 통합
    - [x] 입력과 출력 자동 비교
    - [x] 매핑 자동 업데이트
    - [x] 지속적 개선 사이클

### 💰 식봄(Foodspring) 단가 스크래핑
- [x] `execution/fetch_foodspring_prices.py` 구현
    - [x] 로그인 및 모바일 검색 우회 로직
    - [x] 이메일 로그인 지원
    - [x] 가격 데이터 자동 추출
- [x] `execution/update_market_prices.py` 구현
    - [x] 엑셀 단가표 자동 업데이트

### 📈 경매가 자동화 🚧 진행 중
- [/] `execution/update_auction_prices.py` 구현 (233줄)
    - [x] 가락시장 경매 API 연동 (서울청과 법인)
    - [x] 70+ 품목 매핑 (Excel 품목명 → API 검색어)
    - [x] 경매 최고가/평균가 자동 조회
    - [x] Excel 날짜 +1일 경매 데이터 조회 (저녁 경매 → 다음날 새벽 종료)
    - [x] 마스터 파일 단가시트 자동 업데이트 (F/G열)
    - [x] CSV 데이터 자동 저장 (`auction_data_YYYYMMDD.csv`)
    - [x] 병합셀 안전 처리
    - [ ] 운영 안정화 및 검증 완료 필요

### 📰 일일 시황 리포트
- [x] `execution/daily_market_report.py` 구현
    - [x] 시장 데이터 수집 (yfinance)
    - [x] 뉴스 스크래핑 (feedparser + newspaper)
    - [x] AI 리포트 생성 (Gemini API)
    - [x] 카카오톡 자동 발송

---

## 4. 향후 개발 계획

### 📈 경매가 자동화 완성 (우선순위: 높음)
- [ ] **경매가 운영 안정화**
    - [ ] API 응답 검증 및 오류 처리 강화
    - [ ] 일일 워크플로우 연동
    - [ ] 경매가 데이터 정확도 검증

### 📊 마진 및 단가 계산 고도화 (우선순위: 중간)
- [ ] **자동 단가 계산 고도화**
    - [ ] 경매가 기반 매입 단가 추천 로직
    - [ ] 마진율 기반 판매 단가 자동 조정
    - [ ] 이상 단가 감지 및 알림

### 🔄 통합 자동화 확장 (우선순위: 높음)
- [ ] `fill_prices.py`를 `run_full_workflow.py`에 통합
- [ ] 4단계 + 4-1단계 통합 실행 스크립트 (`run_final_step.py`)
- [ ] 전체 파이프라인 원클릭 실행 (1~4-1단계)

### 🔍 데이터 분석 및 리포팅 (우선순위: 낮음)
- [ ] **매출 분석 대시보드**
    - [ ] 매장별 매출 추이 분석
    - [ ] 품목별 수익성 분석
    - [ ] 공급업체별 거래량 분석
- [ ] **자동 리포트 생성**
    - [ ] 일일/주간/월간 리포트
    - [ ] 이상 패턴 감지 및 알림

### 🌐 웹 인터페이스 (우선순위: 낮음)
- [ ] **관리자 대시보드**
    - [ ] 발주 현황 실시간 조회
    - [ ] 거래명세서 온라인 조회
    - [ ] 매핑 규칙 웹 편집기
- [ ] **모바일 앱**
    - [ ] 발주 입력 모바일 앱
    - [ ] 실시간 알림

---

## 📊 프로젝트 현황

### ✅ 완료된 기능
- 발주 처리 자동화 (1~3-2단계 통합)
- 매입/판매단가 자동 입력 (3.5단계)
- 송부용 거래명세서 생성 (4단계)
- 거래원장 자동 입력 (4-1단계)
- AI 학습 시스템
- 식봄 단가 스크래핑
- 일일 시황 리포트 (카카오톡)
- Google Drive 자동 업로드

### 🚧 진행 중
- 경매가 자동화 안정화 및 검증
- 통합 워크플로우 확장 (fill_prices 연동)
- 시스템 안정화 및 최적화

### 📅 다음 단계
1. fill_prices → 통합 워크플로우 연동 (우선순위 1)
2. 4단계 + 4-1단계 통합 스크립트 (우선순위 2)
3. 마진/단가 계산 고도화 (우선순위 3)

---

**최종 업데이트**: 2026-02-14  
**버전**: 2.0  
**상태**: 매입/판매단가 자동입력·거래원장 자동입력 완료, 경매가 자동화 진행 중, 통합 워크플로우 확장 진행 중
