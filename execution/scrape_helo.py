"""
동원 HELO 사이트에서 구매오더(물동량 조회) 데이터를 스크래핑하여
도크발주관리데이터.xlsx 재고 시트의 2차출고 컬럼에 자동 입력.

HELO는 nexacro (TOBESOFT) 프레임워크 — 표준 HTML이 아닌 JS 컴포넌트 기반.
검증된 접근 순서:
  1. Playwright headless → HELO 페이지 로딩 (20초 대기)
  2. nexacro JS API → loginDiv.loginFormDiv.{USER_ID, USER_PASSWORD, loginButton}
  3. 비밀번호 변경 팝업 → MouseEvent dispatch로 닫기
  4. 공급처전용 탭 (btn1396) → [일배] 발주정보 (gridrow_3) → 구매오더 (gridrow_7)
  5. 조회 (searchButton) → dsEaiOrder 데이터셋 추출 (ZTOT 소계 행만)

Usage:
  python execution/scrape_helo.py --date 2026-02-23
  python execution/scrape_helo.py --date 2026-02-23 --extract-only
  python execution/scrape_helo.py --date 2026-02-23 --dry-run
"""
import argparse
import json
import os
import re
import time
from datetime import datetime

import openpyxl

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_MASTER = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
MAPPINGS_PATH = os.path.join(
    project_root, "skills", "order_processing", "resources", "mappings.json"
)

HELO_URL = "https://helo.dongwon.com/HELO/helo/index.html"
HELO_ID = "1728602461"
HELO_PW = "1728602461"

# HELO 자재코드(mATNR) → 재고시트 매핑 (가장 안정적: 코드는 변하지 않음)
# "convert": "count" = mENGE 그대로, "weight" = mENGE * spec_kg
# 재고시트에 없는 품목은 sheet_name=None (스킵)
HELO_CODE_MAP = {
    # 채소류
    "334491": {"sheet_name": "고추(KG)", "convert": "weight"},        # 고추(청양) 500G/EA
    "334467": {"sheet_name": "깐양파5kg", "convert": "count"},       # 깐양파 5KG/망
    "335163": {"sheet_name": "깐양파1kg", "convert": "count"},       # 깐양파 1KG*10EA/BOX (개별 1kg짜리)
    "334492": {"sheet_name": "깐대파(단)", "convert": "count"},      # 대파(깐) 1KG/단
    "334525": {"sheet_name": "양상추(12수/통)", "convert": "count"}, # 양상추 300G*12EA/BOX
    "334536": {"sheet_name": "알배기(통)", "convert": "count"},      # 알배기 500G*12EA/BOX
    "334900": {"sheet_name": "청경채(KG)", "convert": "weight"},     # 청경채 500G/EA
    "334521": {"sheet_name": None, "convert": "count"},              # 숙주 3.5KG/BOX (재고시트 확인 필요)
    "334523": {"sheet_name": None, "convert": "count"},              # 콩나물(곱슬) 3.5KG/BOX
    "334518": {"sheet_name": None, "convert": "count"},              # 부추 500G*10EA/BOX
    "334539": {"sheet_name": None, "convert": "count"},              # 깻잎 10G*100EA/BOX
    "334899": {"sheet_name": None, "convert": "count"},              # 피망(청) 300G/EA
    "334586": {"sheet_name": None, "convert": "count"},              # 베이비채소 50G/EA
    "334585": {"sheet_name": None, "convert": "count"},              # 베이비순 500G/EA
    "334590": {"sheet_name": None, "convert": "count"},              # 마늘(50G/EA) — 소량, 재고 불필요
    "334597": {"sheet_name": "마늘500g", "convert": "count"},        # 간마늘 500G/EA → 1개=1단위
    "334598": {"sheet_name": "깐쪽파(단)", "convert": "multiply", "factor": 0.5},  # 깐쪽파 500G=반단
    "334538": {"sheet_name": "감자(KG)", "convert": "weight"},       # 감자 1KG/EA
    "334530": {"sheet_name": "맛느타리(개)", "convert": "multiply", "factor": 3},  # 느타리 500G/EA (200g×3개 환산)
    "334528": {"sheet_name": "맛느타리(개)", "convert": "count"},    # 느타리 200G/EA → 1EA=1개
    "334494": {"sheet_name": "양배추(통)", "convert": "multiply", "factor": 3},   # 양배추 8KG/망 → 1망=3통
    # 버섯/과일
    "334526": {"sheet_name": "새송이(KG)", "convert": "weight"},     # 새송이 2KG/BOX
    "334594": {"sheet_name": "팽이버섯(팩)", "convert": "count"},    # 팽이버섯 120G*34EA/BOX
    "334532": {"sheet_name": "레몬(개)", "convert": "count"},        # 레몬 100G*140EA/BOX
    "334844": {"sheet_name": "감자(KG)", "convert": "weight"},       # 못난이감자(왕왕) 3KG/BOX
    "334577": {"sheet_name": None, "convert": "count"},              # 바나나
    # 가공식품
    "166989": {"sheet_name": "김치(박스)", "convert": "count"},      # 맛김치 10KG/BOX
    "167383": {"sheet_name": "마카로니", "convert": "count"},         # 종합마카로니 2.5KG/BOX
    "180449": {"sheet_name": "아카시아청", "convert": "count"},      # 아카시아청 1.2KG*8EA/BOX
    # 기타
    "335145": {"sheet_name": None, "convert": "count"},              # 라임
    "334466": {"sheet_name": "깐양파10kg", "convert": "count"},     # 깐양파 10KG (별도 코드)
}

# HELO mAKTX 키워드 → 재고시트 자동매칭 규칙
# (keyword_in_mAKTX, sheet_item_keyword, convert_mode)
# convert_mode: "kg" = qty * unit_weight_kg, "count" = qty 그대로
HELO_KEYWORD_MAP = [
    # 청양고추: 500G/EA → KG 변환
    ("고추(청양", "고추", "kg"),
    ("깐양파", "깐양파", "count"),
    ("대파", "깐대파", "count"),
    ("양배추", "양배추", "count"),
    ("감자", "감자", "kg"),
    ("당근", "당근", "kg"),
    ("팽이", "팽이버섯", "count"),
    ("새송이", "새송이", "kg"),
    ("표고", "표고버섯", "kg"),
    ("양상추", "양상추", "count"),
    ("적채", "적채", "count"),
    ("취청오이", "취청오이", "kg"),
    ("청경채", "청경채", "kg"),
    ("알배기", "알배기", "count"),
    ("무(", "무", "count"),
    ("배추", "배추", "count"),
    ("숙주", "숙주", "count"),
    ("콩나물", "콩나물", "count"),
    ("김치", "김치", "count"),
    ("깻잎", "깻잎", "count"),
    ("애호박", "애호박", "count"),
    ("쪽파", "깐쪽파", "count"),
    ("부추", "부추", "count"),
    ("마늘", "마늘", "kg"),
    ("레몬", "레몬", "count"),
    ("쌀", "쌀", "count"),
]


def load_helo_code_map():
    """mappings.json + 스크립트 내장 HELO_CODE_MAP 병합.
    mappings.json의 값이 우선 (코드 수정 없이 매핑 업데이트 가능).
    """
    code_map = dict(HELO_CODE_MAP)
    try:
        with open(MAPPINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "helo_code_map" in data:
            code_map.update(data["helo_code_map"])
    except Exception:
        pass
    return code_map


# ─── Playwright / HELO 접속 ────────────────────────────────────

def init_browser(headless=True):
    """Playwright 브라우저 초기화"""
    import sys as _sys
    _sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from _browser import ensure_browser, create_browser
    ensure_browser()
    pw, browser, page = create_browser(headless=headless)
    # dialog(alert) 자동 처리
    page.on("dialog", lambda d: d.accept())
    return pw, browser, page


def wait_for_nexacro(page, timeout=30):
    """nexacro 프레임워크 + 로그인 폼 로딩 대기 (1초 간격 폴링)"""
    print("Waiting for nexacro to load...", flush=True)
    app_ready = False
    for i in range(timeout):
        try:
            result = page.evaluate(
                """
                () => {
                    try {
                        var app = nexacro.getApplication();
                        if (!app) return 'no_app';
                        if (!app.mainframe) return 'no_mainframe';
                        if (!app.mainframe.childframe) return 'no_childframe';
                        var form = app.mainframe.childframe.form;
                        if (!form) return 'no_form';
                        if (form.loginDiv) return 'login_ready';
                        if (form.workDiv) return 'main_ready';
                        return 'form_loading';
                    } catch(e) {
                        return 'not_loaded';
                    }
                }
                """
            )
            if result in ("login_ready", "main_ready"):
                print(f"  nexacro {result} ({i + 1}s)", flush=True)
                return True
            elif not app_ready and result not in ("not_loaded", "no_app"):
                app_ready = True
                print(f"  nexacro loading... ({result})", flush=True)
        except Exception:
            pass
        time.sleep(1)
    print(f"  nexacro not loaded after {timeout}s", flush=True)
    return False


def login_helo(page):
    """HELO 로그인 (nexacro JS API — 검증된 경로)"""
    print(f"Navigating to {HELO_URL}...", flush=True)
    page.goto(HELO_URL, wait_until="domcontentloaded", timeout=60000)

    if not wait_for_nexacro(page, 60):
        return False

    # 로그인 폼 접근: loginDiv → loginFormDiv → USER_ID / USER_PASSWORD / loginButton
    print("Logging in via nexacro API...", flush=True)
    result = page.evaluate(
        f"""
        () => {{
            try {{
                var app = nexacro.getApplication();
                var form = app.mainframe.childframe.form;
                var loginForm = form.loginDiv.form.loginFormDiv.form;
                loginForm.USER_ID.set_value("{HELO_ID}");
                loginForm.USER_PASSWORD.set_value("{HELO_PW}");
                loginForm.loginButton.click();
                return "ok";
            }} catch(e) {{
                return "error: " + e.message;
            }}
        }}
    """
    )
    print(f"  Login result: {result}", flush=True)

    if result != "ok":
        return False

    time.sleep(5)

    # 비밀번호 변경 팝업 닫기
    dismiss_password_popup(page)
    time.sleep(3)

    # mainForm 로딩 확인
    check = page.evaluate(
        """
        () => {
            try {
                var app = nexacro.getApplication();
                var form = app.mainframe.childframe.form;
                if (form.workDiv || form.mainDiv) return "main_loaded";
                return "unknown_state";
            } catch(e) {
                return "error: " + e.message;
            }
        }
    """
    )
    print(f"  Post-login state: {check}", flush=True)
    return "main_loaded" in check or "unknown" in check


def dismiss_password_popup(page):
    """비밀번호 변경 유효기간 팝업 닫기.

    nexacro 팝업의 closeButton은 DOM .click()이 안 먹으므로
    MouseEvent(mousedown→mouseup→click) 시퀀스를 dispatch.
    """
    print("  Checking for password popup...", flush=True)
    try:
        result = page.evaluate(
            """
            () => {
                var btn = document.querySelector("[id*=closeButton]");
                if (btn) {
                    ['mousedown', 'mouseup', 'click'].forEach(function(evtType) {
                        var evt = new MouseEvent(evtType, {
                            bubbles: true, cancelable: true, view: window
                        });
                        btn.dispatchEvent(evt);
                    });
                    return "dismissed";
                }
                return "no_popup";
            }
        """
        )
        print(f"  Popup: {result}", flush=True)
    except Exception as e:
        print(f"  Popup check error: {e}", flush=True)


def navigate_to_purchase_order(page):
    """구매오더(물동량조회) 메뉴로 이동.

    경로: 공급처전용 탭 → [일배] 발주정보 트리 → 구매오더(물동량조회)
    Playwright의 locator.click()은 물리 클릭과 동등.
    """
    # Step 1: 공급처전용 탭 (btn1396)
    print("Step 1: Clicking 공급처전용 tab...", flush=True)
    try:
        result = page.evaluate(
            """
            () => {
                var btn = document.querySelector("[id*=btn1396]");
                if (btn) {
                    ['mousedown', 'mouseup', 'click'].forEach(function(t) {
                        btn.dispatchEvent(new MouseEvent(t, {bubbles:true, cancelable:true, view:window}));
                    });
                    return "clicked";
                }
                return "not_found";
            }
        """
        )
        print(f"  btn1396: {result}", flush=True)
    except Exception as e:
        print(f"  Tab click error: {e}", flush=True)
        return False

    time.sleep(3)

    # Step 2: [일배] 발주정보 트리 확장 (gridrow_3)
    print("Step 2: Expanding [일배] 발주정보...", flush=True)
    try:
        page.locator("[id*=gridrow_3]").first.click()
        print("  Expanded gridrow_3", flush=True)
    except Exception as e:
        print(f"  Tree expand error: {e}", flush=True)
        # Fallback: try other gridrow indices
        for idx in [2, 4, 5]:
            try:
                page.locator(f"[id*=gridrow_{idx}]").first.click()
                print(f"  Fallback: expanded gridrow_{idx}", flush=True)
                break
            except Exception:
                continue

    time.sleep(2)

    # Step 3: 구매오더(물동량조회) 클릭 (gridrow_7)
    print("Step 3: Clicking 구매오더(물동량조회)...", flush=True)
    try:
        page.locator("[id*=gridrow_7]").first.click()
        print("  Clicked gridrow_7", flush=True)
    except Exception as e:
        print(f"  Menu click error: {e}", flush=True)
        # Fallback: try adjacent indices
        for idx in [6, 8, 9]:
            try:
                page.locator(f"[id*=gridrow_{idx}]").first.click()
                print(f"  Fallback: clicked gridrow_{idx}", flush=True)
                break
            except Exception:
                continue

    time.sleep(5)
    return True


def set_date_and_search(page, target_date):
    """날짜 설정 + 조회 버튼 클릭.

    target_date: "YYYY-MM-DD" 형식. None이면 기본 날짜(오늘)로 조회.
    """
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    date_val = dt.strftime("%Y%m%d")

    print(f"Setting date to {date_val} and searching...", flush=True)
    result = page.evaluate(
        f"""
        () => {{
            try {{
                var app = nexacro.getApplication();
                var form = app.mainframe.childframe.form;
                var win = form.workDiv.form.WIN1410.form;
                win.divSearch.form.calSGiday.set_value("{date_val}");
                win.buttonDiv.form.searchButton.click();
                return "ok";
            }} catch(e) {{
                return "error: " + e.message;
            }}
        }}
    """
    )
    print(f"  Search: {result}", flush=True)
    time.sleep(6)
    # dialog handler auto-accepts alerts
    time.sleep(1)
    return result == "ok"


# ─── 데이터 추출 / 파싱 ──────────────────────────────────────────

def extract_and_aggregate(page):
    """dsEaiOrder에서 자재코드별 집계 데이터 추출.

    HELO 구조: 각 자재에 식당별 행 + ZTOT(소계) 행.
    ZTOT 행은 mAKTX/gROES가 비어있으므로 일반 행에서 이름/규격을 가져옴.

    JS에서 직접 그룹핑하여 데이터 전송량 최소화.

    Returns: list of {mATNR, mAKTX, gROES, mEINS, total_qty, store_count}
    """
    print("Extracting and aggregating dsEaiOrder...", flush=True)
    result = page.evaluate(
        """
        () => {
            try {
                var ds = nexacro.getApplication().mainframe.childframe.form
                         .workDiv.form.WIN1410.form.dsEaiOrder;
                var count = ds.getRowCount();
                if (count === 0) return JSON.stringify({count: 0, items: {}});

                var items = {};
                for (var r = 0; r < count; r++) {
                    var code = String(ds.getColumn(r, 'mATNR') || '');
                    var kunnr = String(ds.getColumn(r, 'kUNNR') || '');

                    if (!items[code]) {
                        items[code] = {
                            mATNR: code, mAKTX: '', gROES: '', mEINS: '',
                            total_qty: 0, store_count: 0
                        };
                    }

                    if (kunnr === 'ZTOT') {
                        items[code].total_qty += parseFloat(
                            String(ds.getColumn(r, 'mENGE') || '0')
                        );
                    } else {
                        items[code].store_count++;
                        if (!items[code].mAKTX) {
                            items[code].mAKTX = String(ds.getColumn(r, 'mAKTX') || '');
                            items[code].gROES = String(ds.getColumn(r, 'gROES') || '');
                            items[code].mEINS = String(ds.getColumn(r, 'mEINS') || '');
                        }
                    }
                }
                return JSON.stringify({count: count, items: items});
            } catch(e) {
                return JSON.stringify({error: e.message});
            }
        }
    """
    )

    data = json.loads(result)
    if "error" in data:
        print(f"  Error: {data['error']}", flush=True)
        return []

    total = data.get("count", 0)
    items = data.get("items", {})
    print(f"  Total rows: {total}, Unique items: {len(items)}", flush=True)

    return list(items.values())


def parse_spec_kg(groes):
    """gROES (규격) 문자열에서 EA당 중량(kg) 파싱.

    HELO 수량(mENGE)은 EA 단위이므로, 복합 규격에서 EA당 중량만 반환.

    Examples:
        "500G/EA" → 0.5
        "5KG/망" → 5.0
        "1KG/단" → 1.0
        "1KG*10EA/BOX" → 1.0  (EA당 1kg, BOX당 10kg 아님)
        "1KG*4EA/BOX" → 1.0   (EA당 1kg)
        "10G*100EA/BOX" → 0.01 (EA당 10g)
    Returns: float or None
    """
    g = groes.upper().strip()

    # Compound: NUMunit * NUMea — EA당 중량만 추출 (곱하지 않음)
    m = re.match(r"(\d+(?:\.\d+)?)\s*(KG|G)\s*\*\s*(\d+)", g)
    if m:
        val, unit = float(m.group(1)), m.group(2)
        return (val / 1000) if unit == "G" else val

    # Simple: NUMkg or NUMg
    m = re.match(r"(\d+(?:\.\d+)?)\s*(KG|G)", g)
    if m:
        val, unit = float(m.group(1)), m.group(2)
        return val / 1000 if unit == "G" else val

    return None


def match_helo_to_sheet(ztot_items, item_rows, helo_code_map):
    """HELO 품목 → 재고시트 매핑.

    우선순위:
      1. helo_code_map (자재코드 직접 매핑 — 가장 정확)
      2. HELO_KEYWORD_MAP (mAKTX 키워드 → 재고시트 키워드 매칭)

    Returns: list of (sheet_item, row, final_qty, helo_name, matched_by)
             sheet_item=None이면 매칭 실패
    """
    # 재고시트 품목 키워드 인덱스 (base keyword → full item name, row)
    sheet_index = {}
    for sheet_item, row in item_rows.items():
        base = re.split(r"[\(\d]", sheet_item)[0].strip()
        if base and len(base) >= 1:
            sheet_index[base] = (sheet_item, row)

    results = []

    for item in ztot_items:
        code = item["mATNR"]
        raw_name = item["mAKTX"]
        groes = item["gROES"]
        qty = item["total_qty"]

        # Clean: "[C]고추(청양/국내산/도크/500g/..)" → "고추(청양/국내산/도크/500g/..)"
        clean = re.sub(r"^\[.*?\]", "", raw_name).strip()

        matched = False

        # 1) Code map (direct mATNR mapping)
        if code in helo_code_map:
            info = helo_code_map[code]
            sn = info.get("sheet_name", "")
            conv = info.get("convert", "count")

            if sn:
                # Exact match first
                target_row = item_rows.get(sn)

                # Fuzzy match: base keyword (e.g., "팽이버섯(팩)" ↔ "팽이버섯(개)")
                if target_row is None:
                    sn_base = re.split(r"[\(\d]", sn)[0].strip()
                    if sn_base:
                        for si, sr in item_rows.items():
                            si_base = re.split(r"[\(\d]", si)[0].strip()
                            if sn_base == si_base:
                                sn = si  # Use actual sheet name
                                target_row = sr
                                break

                if target_row is not None:
                    if conv == "weight":
                        spec_kg = parse_spec_kg(groes)
                        fq = qty * spec_kg if spec_kg else qty
                    elif conv == "multiply":
                        fq = qty * info.get("factor", 1)
                    else:
                        fq = qty
                    results.append((sn, target_row, fq, raw_name, "code"))
                    matched = True

        # 2) Keyword auto-match
        if not matched:
            for helo_kw, sheet_kw, conv_mode in HELO_KEYWORD_MAP:
                if helo_kw not in clean:
                    continue

                # Find matching sheet item
                target = None
                for sbase, (sitem, srow) in sheet_index.items():
                    if sheet_kw in sbase or sbase in sheet_kw:
                        target = (sitem, srow)
                        break

                if not target:
                    # Try direct lookup in item_rows
                    for sitem, srow in item_rows.items():
                        if sheet_kw in sitem:
                            target = (sitem, srow)
                            break

                if target:
                    sitem, srow = target
                    if conv_mode == "kg":
                        spec_kg = parse_spec_kg(groes)
                        fq = qty * spec_kg if spec_kg else qty
                    else:
                        fq = qty

                    results.append((sitem, srow, fq, raw_name, "keyword"))
                    matched = True
                    break

        if not matched:
            results.append((None, None, qty, raw_name, "none"))

    return results


# ─── 재고시트 입력 ────────────────────────────────────────────────

def find_date_columns(ws, target_date):
    """재고 시트에서 target_date의 2차출고 컬럼 찾기.

    Row 1-5 헤더에서 "M/DD 재고현황" 검색.
    날짜 그룹: 품목(+0) | 재고(+1) | 1차/직납(+2) | 2차출고(+3) | 입고(+4) | 현재고(+5)
    """
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    month, day = dt.month, dt.day

    # Exact date match: "2/23"이 "12/23" 안에서 매칭되는 것 방지
    date_prefix = f"{month}/{day}"

    for row_idx in range(1, 6):
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=row_idx, column=col_idx).value
            if cell_val is None:
                continue
            cell_str = str(cell_val).strip()
            if "재고" not in cell_str:
                continue
            # Exact match: date_prefix must be at start or preceded by non-digit
            idx = cell_str.find(date_prefix)
            if idx < 0:
                continue
            if idx > 0 and cell_str[idx - 1].isdigit():
                continue  # e.g., "12/23" when looking for "2/23"
            # 품목=col, 2차출고=col+3
            return col_idx, col_idx + 3

    return None, None


def find_item_rows(ws, item_col):
    """재고시트 품목 컬럼 → {품목명: 행번호} 매핑"""
    row_map = {}
    for r in range(4, min(ws.max_row + 1, 200)):
        val = ws.cell(row=r, column=item_col).value
        if val:
            row_map[str(val).strip()] = r
    return row_map


def fill_helo_data(master_file, target_date, match_results, dry_run=False):
    """재고시트 2차출고 컬럼에 HELO 데이터 입력"""
    print(f"\nOpening: {master_file}", flush=True)
    wb = openpyxl.load_workbook(master_file)
    ws = wb["재고"]

    item_col, outgoing_col = find_date_columns(ws, target_date)
    if item_col is None:
        print(f"Error: Could not find columns for {target_date}", flush=True)
        wb.close()
        return

    print(f"Columns: 품목=C{item_col}, 2차출고=C{outgoing_col}", flush=True)

    updates = 0
    for sheet_item, row, qty, helo_name, match_by in match_results:
        if not sheet_item or not row:
            continue

        cell = ws.cell(row=row, column=outgoing_col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            print(f"  MERGED: {sheet_item} (R{row}), skipped", flush=True)
            continue

        # qty를 정수화 (소수점 불필요 시)
        display_qty = int(qty) if qty == int(qty) else round(qty, 1)

        # 같은 셀에 여러 HELO 품목이 매핑될 수 있음 (예: 청경채 1kg + 500g)
        # 기존 값이 있으면 누적
        existing = cell.value
        if existing and isinstance(existing, (int, float)):
            final_qty = existing + qty
        else:
            final_qty = qty
        display_final = int(final_qty) if final_qty == int(final_qty) else round(final_qty, 1)

        if dry_run:
            if existing:
                print(
                    f"  {sheet_item} (R{row}): {existing}+{display_qty}={display_final} [DRY RUN] ← {helo_name[:40]}",
                    flush=True,
                )
            else:
                print(
                    f"  {sheet_item} (R{row}): {display_final} [DRY RUN] ← {helo_name[:40]}",
                    flush=True,
                )
        else:
            cell.value = display_final
            if existing:
                print(
                    f"  {sheet_item} (R{row}): {existing}+{display_qty}={display_final} ← {helo_name[:40]}",
                    flush=True,
                )
            else:
                print(
                    f"  {sheet_item} (R{row}): {display_final} ← {helo_name[:40]}",
                    flush=True,
                )
        updates += 1

    print(f"\nUpdated: {updates}", flush=True)

    if updates > 0 and not dry_run:
        for attempt in range(5):
            try:
                print(f"Saving... (attempt {attempt + 1})", flush=True)
                wb.save(master_file)
                print("Saved.", flush=True)
                break
            except PermissionError:
                if attempt < 4:
                    print("  File locked, retrying in 3s...", flush=True)
                    time.sleep(3)
                else:
                    import tempfile

                    tmp = os.path.join(
                        tempfile.gettempdir(), "도크발주관리데이터_helo.xlsx"
                    )
                    wb.save(tmp)
                    print(f"  Saved to temp: {tmp}", flush=True)
    elif dry_run:
        print("[DRY RUN] No changes saved.", flush=True)

    wb.close()


# ─── 메인 ────────────────────────────────────────────────────────

def scrape_helo(target_date, extract_only=False, dry_run=False, master_file=DEFAULT_MASTER):
    """메인: HELO 스크래핑 → 재고시트 2차출고 입력"""
    # Streamlit Cloud에서는 HELO 접속 불가 (내부망/IP 제한)
    if os.path.exists("/mount/src") or os.environ.get("STREAMLIT_SHARING_MODE") == "1":
        print("HELO 스크래핑은 로컬 환경에서만 실행 가능합니다. "
              "(동원 HELO는 내부망/IP 제한으로 Cloud 접속 불가)", flush=True)
        return None

    pw = None
    browser = None
    page = None
    try:
        pw, browser, page = init_browser()

        # 1. 로그인
        if not login_helo(page):
            print("HELO login failed.", flush=True)
            return None

        # 2. 구매오더 메뉴 이동
        navigate_to_purchase_order(page)

        # 3. 날짜 설정 + 조회
        if not set_date_and_search(page, target_date):
            print("Search failed.", flush=True)
            return None

        # 4. 데이터 추출 + 집계 (JS에서 그룹핑)
        agg_items = extract_and_aggregate(page)
        if not agg_items:
            print("No data for this date.", flush=True)
            ss_path = os.path.join(project_root, "debug_helo_screenshot.png")
            try:
                page.screenshot(path=ss_path)
                print(f"Screenshot: {ss_path}", flush=True)
            except Exception:
                pass
            return None

        # 5. 결과 출력
        print(f"\n=== HELO Items ({len(agg_items)}) ===", flush=True)
        for item in agg_items:
            qty = item["total_qty"]
            qty_d = int(qty) if qty == int(qty) else qty
            print(
                f"  [{item['mATNR']}] {item['mAKTX'][:50]}  "
                f"spec={item['gROES']}  qty={qty_d}  stores={item.get('store_count', 0)}",
                flush=True,
            )

        if extract_only:
            # Save to JSON for debugging
            out_path = os.path.join(project_root, "helo_extracted_data.json")
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(agg_items, f, ensure_ascii=False, indent=2)
            print(f"\n[EXTRACT ONLY] Saved to {out_path}", flush=True)
            return agg_items

        # 6. 재고시트 매핑
        print(f"\nOpening {master_file} for mapping...", flush=True)
        wb_tmp = openpyxl.load_workbook(master_file, read_only=True)
        ws_tmp = wb_tmp["재고"]
        item_col, _ = find_date_columns(ws_tmp, target_date)
        if item_col is None:
            print(f"Error: Cannot find columns for {target_date}", flush=True)
            wb_tmp.close()
            return agg_items
        item_rows = find_item_rows(ws_tmp, item_col)
        wb_tmp.close()

        helo_code_map = load_helo_code_map()
        match_results = match_helo_to_sheet(agg_items, item_rows, helo_code_map)

        # 결과 출력
        print(f"\n=== Matching Results ===", flush=True)
        matched_count = 0
        no_match_list = []
        for sheet_item, row, qty, helo_name, match_by in match_results:
            if sheet_item:
                qty_d = int(qty) if qty == int(qty) else round(qty, 1)
                print(f"  OK [{match_by}] {helo_name[:40]} -> {sheet_item} = {qty_d}", flush=True)
                matched_count += 1
            else:
                no_match_list.append(helo_name[:50])
                print(f"  NO MATCH: {helo_name[:50]}  (qty={qty})", flush=True)

        print(f"\nMatched: {matched_count}/{len(match_results)}", flush=True)
        if no_match_list:
            print(f"Unmatched: {len(no_match_list)}", flush=True)

        # 7. 재고시트 입력
        if matched_count > 0:
            fill_helo_data(master_file, target_date, match_results, dry_run)
        else:
            print("No items matched. Skipping sheet fill.", flush=True)

        return agg_items

    except Exception as e:
        print(f"HELO scraping error: {e}", flush=True)
        import traceback
        traceback.print_exc()
        return None

    finally:
        if browser:
            print("\nClosing browser...", flush=True)
            time.sleep(2)
            browser.close()
        if pw:
            pw.stop()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="HELO 구매오더 스크래핑 → 재고시트 2차출고"
    )
    parser.add_argument("--master", default=DEFAULT_MASTER)
    parser.add_argument("--date", required=True, help="YYYY-MM-DD")
    parser.add_argument(
        "--extract-only", action="store_true", help="데이터 추출만 (시트 입력 안함)"
    )
    parser.add_argument("--dry-run", action="store_true", help="미리보기 (저장 안함)")
    args = parser.parse_args()

    import sys
    result = scrape_helo(args.date, args.extract_only, args.dry_run, args.master)
    if result is not None:
        from _notify import notify
        notify("stage4_helo", date_str=args.date)
    else:
        print("FAILED: HELO scraping returned no data.", flush=True)
        sys.exit(1)
