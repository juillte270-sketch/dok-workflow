# 프로젝트 작업 내역 (완료됨)

## 1. 환경 설정 및 기초 작업
- [x] 프로젝트 디렉토리 구조 생성 (`directives`, `execution`, `.tmp`)
- [x] 필수 파일 생성 (`.env`, `.gitignore`)
- [x] 에이전트 지침서(`GEMINI.md` 등) 미러링

## 2. 쓰레드(Threads) 리포트 워크플로우 개발
- [x] **기획 및 요구사항 분석**
    - [x] 구현 계획 수립, 도구 선정
- [x] **실행 스크립트 개발**
    - [x] `analyze_and_report.py`: 데이터 수집, AI 요약, 이메일 발송 통합
    - [x] `kakao_auth.py`: 카카오톡 OAuth 인증 및 토큰 발급

- [x] **기능 구현**
    - [x] **쓰레드 데이터 수집**: Playwright 활용 (@choi.openai)
    - [x] **AI 요약**: Gemini 1.5 Flash 모델 활용하여 게시물 요약
    - [x] **이메일 발송**: Gmail API 연동 (전체 리포트)
    - [x] **카카오톡 발송**: 카카오 API 연동 (요약본 + 이메일 링크)

## 3. DoKH 일일 업무 자동화 (진행 중)
- [x] **환경 설정 및 분석**
    - [x] 디렉토리 구조 생성 (resources, inputs, outputs)
    - [x] 의존성 파일 위치 정의 및 **구글 드라이브 API 연동** (Sheets/Drive API, OAuth 인증 완료)
    - [x] Git 저장소 초기화 및 커밋
- [x] **1단계: 발주 처리 및 배송 리스트** (워크플로우 1)
    - [x] `directives/workflow_01_order_processing.md` 지침서 생성
    - [x] `execution/process_orders.py` 스크립트 구현
- [x] **2단계: 발주 시트 입력** (워크플로우 2)
    - [x] `directives/workflow_02_order_entry.md` 생성
    - [x] 마스터 파일(`도크발주관리데이터.xlsx`) 구조 분석
    - [x] `execution/update_order_sheet.py` 구현 및 실행 완료
    - [x] **거래명세서 생성**
        - [x] 템플릿 설정 (구글 드라이브 연동)
        - [x] `execution/generate_invoices.py` 구현
- [ ] **3단계: 단가 시트 생성** (워크플로우 3)
    - [ ] `directives/workflow_03_price_sheet.md` 생성
    - [ ] `execution/create_price_sheet.py` 구현
- [ ] **4단계: 경매가 업데이트** (워크플로우 4)
    - [ ] `directives/workflow_04_auction_prices.md` 생성
    - [ ] `execution/update_auction_prices.py` 구현
- [ ] **5단계: 마진 및 단가 계산** (워크플로우 5)
    - [ ] `directives/workflow_05_calculate_prices.md` 생성
    - [ ] `execution/calculate_margins.py` 구현
- [ ] **6단계: 최종 명세서 생성** (워크플로우 6)
    - [ ] `directives/workflow_06_generate_statements.md` 생성
    - [ ] `execution/generate_statements.py` 구현
- [ ] **통합 및 검증**
    - [ ] 전체 마스터 워크플로우 생성
    - [ ] 예시 데이터 기반 통합 테스트 (내일 진행 예정)

## 4. 배포 및 자동화 (기존)
- [x] **검증 테스트**: 전체 프로세스(수집-요약-발송) 정상 작동 확인
- [x] **자동화 설정**: 윈도우 작업 스케줄러 등록
    - 매일 **오후 7시** (19:00) 자동 실행
