# DoKH 일일 업무 자동화 시스템

DoKH의 일일 발주 처리부터 거래명세서 생성까지 전체 과정을 자동화하는 시스템입니다.

> 📊 **[프로젝트 대시보드 보기](DASHBOARD.md)** - 전체 현황을 한눈에!

## 🚀 빠른 시작

### 1단계부터 4단계까지 한 번에 실행 (권장)

```bash
python execution/run_full_workflow.py --date "2026-02-05"
```

이 명령 하나로 다음 모든 작업이 자동으로 실행됩니다:
- ✅ 발주 처리 및 배송리스트 생성
- ✅ 발주 시트 입력 (Google Drive)
- ✅ 단가 시트 생성
- ✅ 거래명세서 생성 (16개 매장)
- ✅ **AI 자동 학습** (입력과 출력 비교하여 매핑 자동 업데이트) ⭐ NEW!

## 🤖 AI 분류 및 학습 (NEW!)

발주 리스트를 제공하면 AI가 자동으로 공급업체별로 분류하고, 결과를 검토하여 시스템을 학습시킬 수 있습니다.

### 대화형 분류 및 학습

```bash
python execution/classify_and_learn.py --interactive
```

**워크플로우:**
1. 발주 리스트 입력 (복사-붙여넣기)
2. AI 자동 분류 결과 확인
3. 정확하면 완료, 수정 필요하면 수정 후 자동 학습
4. 다음부터 더 정확하게 분류!

**예시:**
```
📋 발주 리스트 입력:
-샤브야키 도봉점
무순(대) 1팩
숙주 4박스

📦 AI 분류 결과:
-정복상회
무순(대) 1팩

-오복상회
숙주 4박스

✅ 정확! → 완료
또는
❌ 수정 → 자동 학습 → 다음부터 개선!
```

자세한 내용: [분류 및 학습 가이드](directives/classify_and_learn_workflow.md)

## 📋 시스템 구성

### 자동화 단계

1. **발주 처리 및 배송리스트 생성**
   - 원본 발주 데이터 파싱
   - 매장별/공급업체별 분류
   - 4페이지 배송리스트 DOCX 생성

2. **발주 시트 입력**
   - Google Drive 마스터 엑셀 파일 업데이트
   - 매장별 그룹화 및 정렬
   - 이익률 자동 계산

3. **단가 시트 생성**
   - 단가양식 템플릿 복사
   - 발주 품목 기반 필터링
   - 공급업체 매칭 및 중복 처리

4. **거래명세서 생성**
   - 각 매장별 엑셀 거래명세서 생성
   - Google Drive 월별 폴더 자동 저장
   - ZIP 압축 파일 생성

## 📂 파일 구조

```
프로젝트/
├── execution/
│   ├── run_full_workflow.py      # 통합 실행 스크립트 (1~4단계)
│   ├── process_orders.py          # 1단계: 발주 처리
│   ├── update_order_sheet.py      # 2단계: 발주 시트 입력
│   ├── create_price_sheet.py      # 3단계: 단가 시트 생성
│   └── generate_invoices.py       # 4단계: 거래명세서 생성
├── directives/
│   ├── workflow_full_process.md   # 전체 워크플로우 문서
│   ├── workflow_01_order_processing.md
│   ├── workflow_02_order_entry.md
│   └── workflow_03_price_sheet.md
├── data/
│   ├── inputs/
│   │   ├── orders_YYYYMMDD.txt    # 입력: 발주 데이터
│   │   └── processed_orders_YYYYMMDD.txt  # 처리된 발주 데이터
│   └── outputs/
│       ├── 배송리스트_YYYYMMDD.docx
│       ├── invoices_YYYYMMDD/     # 개별 거래명세서
│       └── 거래명세서_YYYYMMDD.zip
└── skills/
    └── order_processing/
        └── resources/
            └── mappings.json       # 품목-공급업체 매핑
```

## 💡 사용 예시

### 기본 사용 (오늘 날짜)

```bash
# 입력 파일: data/inputs/orders_20260205.txt
python execution/run_full_workflow.py --date "2026-02-05"
```

### 커스텀 입력 파일

```bash
python execution/run_full_workflow.py --input "data/inputs/orders_today.txt" --date "2026-02-05"
```

### 거래명세서 생성 건너뛰기

```bash
python execution/run_full_workflow.py --date "2026-02-05" --skip-invoice
```

### 개별 단계 실행

필요시 각 단계를 개별적으로 실행할 수 있습니다:

```bash
# 1단계만
python execution/process_orders.py --input "data/inputs/orders_20260205.txt" --date "2026-02-05"

# 2단계만
python execution/update_order_sheet.py --date "2026-02-05"

# 3단계만
python execution/create_price_sheet.py --input "data/inputs/processed_orders_20260205.txt" --master "G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx" --date "2026-02-05"

# 4단계만
python execution/generate_invoices.py --date "2026-02-05"
```

## 📤 출력 파일

### 로컬 파일
- `data/outputs/배송리스트_YYYYMMDD.docx` - 배송리스트 문서
- `data/outputs/invoices_YYYYMMDD/` - 개별 거래명세서 폴더 (16개 파일)
- `data/outputs/거래명세서_YYYYMMDD.zip` - 거래명세서 압축 파일

### Google Drive
- `도크발주관리데이터.xlsx` - 발주 시트 및 단가 시트 업데이트
- 각 매장 폴더의 `YYYY년 M월/` - 거래명세서 자동 저장

## 🎯 처리 대상 매장

- **샤브야키** (5개점): 도봉점, 주안점, 분당점, 평택점, 동탄점
- **부엉이산장** (3개점): 강남점, 구월점, 마곡점
- **오레노이키루미치** (2개점): 하남점, 압구정점
- **파라이** (2개점): 송파점, 역삼점
- **기타**: 고른햇살, 봄날, 선데이버거클럽, 소유 프루트

## ⚙️ 자동화 설정

Windows 작업 스케줄러를 사용하여 매일 자동 실행:

1. `run_daily.bat` 파일 생성:
```batch
@echo off
cd "C:\Users\DoKH_D\OneDrive\Desktop\안티그래비티 프로젝트 (3)"
python execution/run_full_workflow.py --date "%date:~0,4%-%date:~5,2%-%date:~8,2%"
```

2. Windows 작업 스케줄러에서 매일 특정 시간에 실행하도록 설정

## 🔧 문제 해결

### "입력 파일을 찾을 수 없습니다"
→ `data/inputs/orders_YYYYMMDD.txt` 파일이 있는지 확인

### "마스터 파일을 찾을 수 없습니다"
→ Google Drive가 동기화되어 있는지 확인

### "거래명세서 템플릿을 찾을 수 없습니다"
→ 해당 매장의 Google Drive 폴더에 템플릿이 있는지 확인

## 📚 추가 문서

- [전체 워크플로우 상세 문서](directives/workflow_full_process.md)
- [1단계: 발주 처리](directives/workflow_01_order_processing.md)
- [2단계: 발주 시트 입력](directives/workflow_02_order_entry.md)
- [3단계: 단가 시트 생성](directives/workflow_03_price_sheet.md)
- [분류 및 학습 워크플로우](directives/classify_and_learn_workflow.md)
- [매핑 학습 가이드](directives/mapping_learning_guide.md)
- **[작업 내역 (DoKH)](TASKS_KR.md)** - DoKH 프로젝트 전체 작업 내역
- **[작업 내역 (Threads)](TASKS_THREADS.md)** - Threads 리포트 워크플로우 상세

## 🎉 완료된 기능

- ✅ 발주 처리 및 배송리스트 생성
- ✅ 발주 시트 자동 입력
- ✅ 단가 시트 자동 생성
- ✅ 거래명세서 자동 생성
- ✅ 1~4단계 통합 워크플로우
- ✅ Google Drive 자동 업로드
- ✅ 식봄(Foodspring) 단가 스크래핑
- ✅ 경매가 자동 업데이트

---

**개발**: Google Deepmind Advanced Agentic Coding Team  
**버전**: 1.0  
**최종 업데이트**: 2026-02-05
