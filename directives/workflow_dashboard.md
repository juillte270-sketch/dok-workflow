# Streamlit 운영 대시보드

## 개요
`dashboard/` 폴더에 있는 Streamlit 기반 운영 대시보드.
`execution/` 스크립트를 직접 호출하여 일일 업무를 UI에서 관리합니다.

## 실행 방법
```bash
cd "도크 워크플로우"
streamlit run dashboard/app.py
# 브라우저: http://localhost:8501
```

## 구조
```
dashboard/
├── .streamlit/config.toml    # 다크 테마, 서버 설정
├── app.py                     # 메인 엔트리 (사이드바 + 페이지 라우터)
├── utils.py                   # 공용 유틸리티
│   ├── 날짜 유틸             # today_str, date_to_compact, date_to_iso
│   ├── 설정 관리             # load/save_settings, get_timeout
│   ├── 상태 관리             # load/save_progress, mark_stage, check_prerequisites
│   ├── 실행 이력             # load_history, _append_history
│   ├── 로깅                  # append_log, read_log
│   ├── 스크립트 실행         # run_script, run_script_streaming
│   └── 파일 헬퍼             # check_master_file, list_output_files
├── views/                     # ⚠️ pages/ 아닌 views/ (Streamlit 멀티페이지 충돌 방지)
│   ├── p0_main.py            # 메인 대시보드 + 일괄 실행
│   ├── p1_orders.py          # 발주 처리 (텍스트 입력 → process_orders.py)
│   ├── p2_sheet_invoice.py   # 발주시트 & 가명세서
│   ├── p3_price_sheet.py     # 단가표 관리 (생성 + 식봄 + 경매가)
│   ├── p4_purchase_price.py  # 매입가 직접 입력 + 판매가 계산
│   ├── p5_final_invoice.py   # 최종 명세서 + 거래원장
│   ├── p6_auction.py         # 경매가 실시간 모니터링 + 7일 추이
│   ├── p7_settings.py        # 설정 (경로, 토큰, 타임아웃, 매핑)
│   └── p8_history.py         # 실행 이력 타임라인
├── state/                     # 일별 진행 상태 JSON
├── logs/                      # 일별 실행 로그
└── history/                   # 실행 이력 JSONL
```

## 주요 기능

### 날짜 동기화
- 사이드바 글로벌 날짜 피커 → 모든 페이지에서 동일 날짜 사용
- `st.session_state.target_date`로 관리

### 의존성 그래프
단계 간 선후관계가 강제됨:
```
stage1 (발주 처리)
  ├→ stage2_order (발주시트)
  ├→ stage2_invoice (가명세서)
  └→ stage3_price (단가표)
       ├→ stage3_sikbom (식봄)
       └→ stage5_auction (경매가)
stage2_order + stage3_price → stage7_fill (판매가)
stage7_fill + stage2_invoice → stage8_final (최종 명세서)
```
선행 단계가 완료되지 않으면 버튼이 비활성화됨.

### 일괄 실행
메인 대시보드에서 "오전 일괄 실행" → stage1→2a→2b→3 순차 실행.
이미 완료된 단계는 자동 건너뜀.

### 마스터 파일 잠금 감지
- Windows Excel 잠금 파일 `~$` 체크
- `r+b` 모드로 쓰기 잠금 확인
- 잠금 감지 시 "엑셀을 닫아주세요" 안내

### 실행 이력
- 모든 스크립트 실행이 JSONL 파일에 기록됨
- 타임라인 뷰 + 로그 필터링/검색
- 7일 이전 로그 자동 정리 기능

### 단가표 캐싱
- `@st.cache_data(ttl=30)` 적용으로 30초 내 중복 로딩 방지
- 새로고침 버튼으로 수동 캐시 무효화

## 에러 처리 패턴
1. 마스터 파일 열려 있으면 → `check_master_before_run()` → st.error
2. 스크립트 실패 → stderr 표시 + "재시도" 버튼
3. 타임아웃 → 설정에서 조정 가능 (기본 300초, 식봄 600초)
4. 선행 단계 미완료 → 버튼 비활성화 + 안내 메시지

## 설정 (dashboard/state/settings.json)
```json
{
  "master_file": "G:\\내 드라이브\\...\\도크발주관리데이터.xlsx",
  "drive_folder_id": "1rD5u5OwwmawOSy_3Iu169ltGQwQV1DJW",
  "timeouts": { "default": 300, "sikbom": 600, "auction": 120, "ledger": 600 }
}
```

### 동시 실행 방지 (Execution Lock)
- `run_script_safe()`: 실행 전 잠금 파일 생성, 완료 후 해제
- 15분 이상 된 잠금은 자동 해제 (stale lock)
- 잠금 중 다른 스크립트 실행 시도 시 안내 메시지

### 마스터 파일 자동 백업
- 마스터 파일을 수정하는 스크립트 실행 전 자동 백업
- 최근 5개 백업만 유지, 나머지 자동 삭제
- 설정 > 백업 탭에서 수동 백업 및 다운로드 가능

### 카톡 요약 전송
- 메인 대시보드에서 "카톡 요약 전송" 버튼
- `send_daily_summary.py --results` 형태로 현재 진행 상태 전달

### 마진 분석
- 매입가/판매가 페이지에서 매장별 마진율 차트
- 저마진 품목 자동 감지 (15% 미만)
- 매장별 총매입/평균마진 요약 테이블

### 경매가 다중 품목 비교
- 최대 5개 품목 동시 조회 (병렬 API 호출)
- 비교 차트 + 데이터 테이블

### 단가표 완성도 게이지
- 컬럼별 입력 현황 표시
- 전체 완성도 프로그레스 바
- CSV 내보내기 기능

## 에러 처리: 마스터 파일 잠금 시 동작
- p3, p4: 페이지 UI는 표시하되 실행 버튼만 비활성화 (early return 금지)
- p5: 선행 단계 체크로 가격 미입력 시 명세서 생성 차단

## 발견된 이슈와 해결 (2026-02-27)
- CORS 경고: `enableXsrfProtection = false` 추가로 해결
- 사이드바 진행률: `load_progress(date_str)` 매 렌더링 호출로 실시간 반영
- 날짜 불일치: 글로벌 `st.session_state.target_date` 도입
- 7일 경매 추이: `concurrent.futures.ThreadPoolExecutor` 병렬 조회 (4 workers)
- 매입가 저장: row_map을 파일로 persist하여 페이지 전환 후에도 유지
- 동시 실행 방지: LOCK_FILE 기반 mutex, 15분 stale timeout
- 마스터 백업: 수정 스크립트 실행 전 자동, 최근 5개 유지
- `pages/` → `views/` 변경: Streamlit 멀티페이지 자동 감지와 수동 라우팅 충돌 해결
- 배치 실행 tuple unpacking 버그: `check_master_before_run()` 반환값 처리 수정
