# 재고시트 2차출고 입력 (HELO 스크래핑) 워크플로우

## 개요
동원 HELO 사이트에서 구매오더(물동량 조회) 데이터를 스크래핑하여
도크발주관리데이터.xlsx 재고 시트의 2차출고 컬럼에 자동 입력한다.

## 실행 스크립트
`execution/scrape_helo.py`

## 사용법
```bash
python execution/scrape_helo.py --date 2026-02-23
python execution/scrape_helo.py --date 2026-02-23 --dry-run       # 미리보기
python execution/scrape_helo.py --date 2026-02-23 --extract-only  # JSON 추출만
```

### 파라미터
| 파라미터 | 필수 | 설명 |
|---------|------|------|
| `--master` | X | 마스터 파일 경로 (기본값: G드라이브) |
| `--date` | O | 대상일 (YYYY-MM-DD) |
| `--dry-run` | X | 미리보기만 (저장 안 함) |
| `--extract-only` | X | 데이터 추출만 (시트 입력 안 함, JSON 저장) |

## 입력
- HELO 사이트 (https://helo.dongwon.com) 구매오더 데이터
- `도크발주관리데이터.xlsx` → "재고" 시트

## 출력
- 재고시트 2차출고 컬럼에 품목별 수량 입력
- `--extract-only` 시: `helo_extracted_data.json`

## 실행 시간
- **오후 10시** (동원 마감 후) — HELO 데이터는 당일 오전에 등록, 이른 시간에는 0건일 수 있음

## HELO 사이트 기술 세부사항

### 프레임워크
nexacro (TOBESOFT) — 표준 HTML이 아닌 JS 컴포넌트 기반

### 접속 순서
1. **페이지 로딩**: Selenium 비-headless (nexacro는 headless 불가) → 최대 60초 대기
2. **nexacro 대기**: `app.mainframe.childframe.form.loginDiv` 존재 확인 (1초 간격 폴링)
3. **로그인**: JS API `loginDiv.form.loginFormDiv.form.{USER_ID, USER_PASSWORD, loginButton}`
   - ID/PW: `1728602461`
4. **비밀번호 팝업 닫기**: `[id*=closeButton]`에 MouseEvent(mousedown→mouseup→click) dispatch
5. **메뉴 이동**:
   - btn1396 (공급처전용 탭) → gridrow_3 ([일배] 발주정보) → gridrow_7 (구매오더)
   - ActionChains 물리클릭 필수
6. **날짜 설정**: `win.divSearch.form.calSGiday.set_value("YYYYMMDD")`
7. **조회**: `win.buttonDiv.form.searchButton.click()` → 데이터 없으면 JS alert → dismiss
8. **데이터 추출**: `dsEaiOrder` 데이터셋에서 JS로 직접 그룹핑

### 데이터셋 구조 (dsEaiOrder)
- ZTOT 행 (`kUNNR="ZTOT"`): 소계 행, `mENGE`(수량)만 유효
- 일반 행: `mAKTX`(품목명), `gROES`(규격), `mATNR`(자재코드) 포함
- JS에서 자재코드별 그룹핑 → total_qty 합산

## 품목 매핑 (우선순위)

### 1. HELO_CODE_MAP (자재코드 직접 매핑 — 최우선)
| 코드 | 품목 | 재고시트 | 변환 |
|------|------|---------|------|
| 334491 | 고추(청양) 500G/EA | 고추(KG) | weight |
| 334467 | 깐양파 5KG/망 | 깐양파5kg | count |
| 335163 | 깐양파 1KG*10EA/BOX | 깐양파1kg | count |
| 334492 | 대파(깐) 1KG/단 | 깐대파(단) | count |
| 334525 | 양상추 300G*12EA/BOX | 양상추(12수/통) | count |
| 334526 | 새송이 2KG/BOX | 새송이(KG) | weight |
| 334594 | 팽이버섯 120G*34EA/BOX | 팽이버섯(팩) | count |
| 334844 | 못난이감자 3KG/BOX | 감자(KG) | weight |
| 166989 | 맛김치 10KG/BOX | 김치(박스) | count |
| ... | (총 25개 코드) | | |

### 2. HELO_KEYWORD_MAP (mAKTX 키워드 매칭 — 폴백)
코드 매핑에 없는 품목은 품목명 키워드로 재고시트 매칭

### 단위 변환
- `weight`: qty × parse_spec_kg(gROES) → KG 환산
  - 예: "500G/EA" → 0.5kg, "1KG*10EA/BOX" → 10kg
- `count`: qty 그대로
- `multiply`: qty × factor

## 전체 파이프라인에서의 위치
```
[4] 재고시트 2차출고 입력 (오후 10시, 동원 마감 후) ← 이 워크플로우
```
