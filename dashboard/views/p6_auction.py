"""Page 6: Auction Price Monitoring - 가락시장 경매가 조회 (Altair UI)."""
import streamlit as st
import altair as alt
import pandas as pd
import requests
import re
import datetime
import concurrent.futures
import urllib.parse

# ---------------------------------------------------------------------------
# API Config (fill_auction_from_api.py와 동일)
# ---------------------------------------------------------------------------
API_BASE = "http://www.garak.co.kr/homepage/publicdata/dataXmlOpen.do"
API_ID = "6335"
API_PASSWD = "tfe7c1p4!!"
API_DATAID = "data68"

# ---------------------------------------------------------------------------
# 품목 리스트 — fill_auction_from_api.py에서 ITEM_SEARCH_MAP/ITEM_OVERRIDES 로드
# ---------------------------------------------------------------------------
_ITEM_SEARCH_MAP = None
_ITEM_OVERRIDES = None


def _load_maps():
    """fill_auction_from_api.py에서 매핑 로드 (lazy, 1회)."""
    global _ITEM_SEARCH_MAP, _ITEM_OVERRIDES
    if _ITEM_SEARCH_MAP is not None:
        return
    try:
        from execution.fill_auction_from_api import ITEM_SEARCH_MAP, ITEM_OVERRIDES
        _ITEM_SEARCH_MAP = ITEM_SEARCH_MAP
        _ITEM_OVERRIDES = ITEM_OVERRIDES
    except Exception:
        _ITEM_SEARCH_MAP = {}
        _ITEM_OVERRIDES = {}


ITEM_CATEGORIES = {
    "엽채류": [
        "청경채", "쑥갓", "알배기배추", "근대", "깻잎(1KG)", "깻잎(4KG,큰잎)",
        "치커리", "케일", "시금치", "아욱", "배추", "양상추(일반)", "로메인(일반)", "통로메인",
    ],
    "근채류": ["양파", "깐양파", "감자", "당근 수입", "무", "깐마늘 대서"],
    "과채류": ["청양고추", "홍고추", "꽈리고추", "취청오이", "애호박", "토마토 완숙"],
    "양념채류": ["흙대파", "깐대파", "깐쪽파", "부추(일반)", "돌미나리", "미나리"],
    "배추류": ["깐양배추45", "양배추45", "빨간양배추 국산", "브로콜리 국산"],
    "버섯류": ["팽이", "새송이", "맛느타리", "꽃느타리", "양송이", "생표고 수입"],
    "과일류": ["귤"],
}

# 전체 품목 (경매 데이터 있는 것만)
ALL_ITEMS = []
for _items in ITEM_CATEGORIES.values():
    ALL_ITEMS.extend(_items)


# ---------------------------------------------------------------------------
# Altair Slate Indigo 테마
# ---------------------------------------------------------------------------
def _register_dok_theme():
    """Altair에 DoK Slate Indigo 테마 등록."""
    def _dok_theme():
        return {
            "config": {
                "background": "transparent",
                "axis": {
                    "labelColor": "#94A3B8",
                    "titleColor": "#94A3B8",
                    "gridColor": "#334155",
                    "domainColor": "#334155",
                    "tickColor": "#334155",
                },
                "legend": {
                    "labelColor": "#E2E8F0",
                    "titleColor": "#94A3B8",
                },
                "title": {"color": "#E2E8F0"},
                "view": {"stroke": "transparent"},
                "range": {
                    "category": [
                        "#818CF8", "#4ADE80", "#FBBF24", "#F87171",
                        "#6366F1", "#64748B",
                    ],
                },
            }
        }
    alt.themes.register("dok_slate", _dok_theme)
    alt.themes.enable("dok_slate")


# ---------------------------------------------------------------------------
# API 검색어 매핑
# ---------------------------------------------------------------------------
def _get_search_term(item_name):
    """품목명 → API 검색어 변환. None이면 경매 데이터 없음."""
    _load_maps()
    if item_name in _ITEM_OVERRIDES:
        return _ITEM_OVERRIDES[item_name].get("search")
    if item_name in _ITEM_SEARCH_MAP:
        return _ITEM_SEARCH_MAP[item_name]
    return item_name


# ---------------------------------------------------------------------------
# API 호출
# ---------------------------------------------------------------------------
def _build_api_url(item_name, s_date, s_date_p, s_date_p7):
    """가락시장 API URL 직접 구성 (params= 사용 금지 — !! 인코딩 문제)."""
    encoded_item = urllib.parse.quote(item_name, encoding="utf-8")
    return (
        f"{API_BASE}?id={API_ID}&passwd={API_PASSWD}&dataid={API_DATAID}"
        f"&pagesize=100&pageidx=1&portal.templet=false"
        f"&s_date={s_date}&s_date_p={s_date_p}&s_date_p7={s_date_p7}"
        f"&p_pos_gubun=1&s_pum_nm=2&s_pummok={encoded_item}"
    )


def _parse_xml_response(text):
    """XML(CDATA) 응답 파싱 → list[dict]."""
    results = []
    rows = re.findall(r"<list>(.*?)</list>", text, re.DOTALL)
    for row in rows:
        item = {
            "품목": _extract_tag(row, "PUM_NM"),
            "등급": _extract_tag(row, "G_NAME"),
            "거래단위": _extract_tag(row, "U_NAME"),
            "최고가": _parse_int(_extract_tag(row, "MA_P")),
            "평균가": _parse_int(_extract_tag(row, "AV_P")),
            "최저가": _parse_int(_extract_tag(row, "MI_P")),
            "전일대비": _extract_tag(row, "FLUC"),
        }
        results.append(item)
    return results


def _extract_tag(text, tag):
    m = re.search(rf"<{tag}><!\[CDATA\[(.*?)\]\]></{tag}>", text)
    if m:
        return m.group(1).strip()
    m = re.search(rf"<{tag}>(.*?)</{tag}>", text)
    if m:
        return m.group(1).strip()
    return ""


def _parse_int(s):
    try:
        return int(str(s).replace(",", ""))
    except (ValueError, AttributeError):
        return 0


@st.cache_data(ttl=300)
def _check_api_health():
    """가락시장 API 서버 연결 가능 여부 확인."""
    try:
        today = datetime.date.today()
        url = _build_api_url(
            "배추",
            today.strftime("%Y%m%d"),
            (today - datetime.timedelta(days=1)).strftime("%Y%m%d"),
            (today - datetime.timedelta(days=7)).strftime("%Y%m%d"),
        )
        resp = requests.get(url, timeout=10)
        resp.encoding = "utf-8"
        return resp.status_code == 200 and "<list>" in resp.text
    except Exception:
        return False


@st.cache_data(ttl=1800)
def _fetch_single_day(search_term, date_str):
    """캐시된 단일일 API 호출. date_str=YYYYMMDD."""
    d = datetime.datetime.strptime(date_str, "%Y%m%d").date()
    prev = (d - datetime.timedelta(days=1)).strftime("%Y%m%d")
    prev7 = (d - datetime.timedelta(days=7)).strftime("%Y%m%d")
    url = _build_api_url(search_term, date_str, prev, prev7)
    try:
        resp = requests.get(url, timeout=15)
        resp.encoding = "utf-8"
        if resp.status_code != 200 or "<html" in resp.text[:200].lower():
            return []
        return _parse_xml_response(resp.text)
    except Exception:
        return []


def _fetch_multi_day(search_term, dates):
    """ThreadPoolExecutor(max_workers=6) 병렬 호출 → {date_str: [rows]}."""
    date_strs = [d.strftime("%Y%m%d") for d in dates]
    results = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
        futures = {
            executor.submit(_fetch_single_day, search_term, ds): ds
            for ds in date_strs
        }
        for future in concurrent.futures.as_completed(futures):
            ds = futures[future]
            try:
                results[ds] = future.result()
            except Exception:
                results[ds] = []
    return results


# ---------------------------------------------------------------------------
# 차트 함수
# ---------------------------------------------------------------------------
def _chart_grades(data, item_name):
    """등급별 가로 막대 차트 (평균가 + 최고가)."""
    valid = [r for r in data if r["평균가"] > 0]
    if not valid:
        return None

    rows = []
    for r in valid:
        label = f"{r['등급']} / {r['거래단위']}"
        rows.append({"등급": label, "가격": r["평균가"], "유형": "평균가"})
        rows.append({"등급": label, "가격": r["최고가"], "유형": "최고가"})

    df = pd.DataFrame(rows)
    chart = (
        alt.Chart(df)
        .mark_bar(cornerRadiusEnd=4)
        .encode(
            y=alt.Y("등급:N", sort=None, title=None),
            x=alt.X("가격:Q", title="원"),
            color=alt.Color(
                "유형:N",
                scale=alt.Scale(
                    domain=["평균가", "최고가"],
                    range=["#818CF8", "#4ADE80"],
                ),
                legend=alt.Legend(orient="top"),
            ),
            yOffset="유형:N",
            tooltip=["등급:N", "유형:N", alt.Tooltip("가격:Q", format=",")],
        )
        .properties(height=max(len(valid) * 50, 120), title=f"{item_name} 등급별 경매가")
    )
    return chart


def _chart_trend(trend_data, item_name, period_label):
    """추이 라인 + 밴드(최고~최저) 차트."""
    df = pd.DataFrame(trend_data)
    if df.empty:
        return None

    # 밴드 (최고가~최저가 범위)
    band = (
        alt.Chart(df)
        .mark_area(opacity=0.15, color="#818CF8")
        .encode(
            x=alt.X("날짜:T", title=None),
            y=alt.Y("최고가:Q", title="원"),
            y2="최저가:Q",
        )
    )

    # 라인 (평균가)
    line = (
        alt.Chart(df)
        .mark_line(strokeWidth=2.5, color="#818CF8", point=True)
        .encode(
            x=alt.X("날짜:T", title=None),
            y=alt.Y("평균가:Q", title="원"),
            tooltip=[
                alt.Tooltip("날짜:T", format="%m/%d"),
                alt.Tooltip("평균가:Q", format=",", title="평균가"),
                alt.Tooltip("최고가:Q", format=",", title="최고가"),
                alt.Tooltip("최저가:Q", format=",", title="최저가"),
            ],
        )
    )

    chart = (
        (band + line)
        .properties(height=320, title=f"{item_name} {period_label} 추이")
        .interactive()
    )
    return chart


def _chart_yoy(this_year_data, last_year_data, item_name):
    """전년 비교 이중 라인 차트."""
    rows = []
    for i, r in enumerate(this_year_data):
        rows.append({"주차": f"W{i+1}", "평균가": r["평균가"], "구분": "금년"})
    for i, r in enumerate(last_year_data):
        rows.append({"주차": f"W{i+1}", "평균가": r["평균가"], "구분": "전년"})

    df = pd.DataFrame(rows)
    if df.empty:
        return None

    chart = (
        alt.Chart(df)
        .mark_line(point=True, strokeWidth=2.5)
        .encode(
            x=alt.X("주차:O", title=None),
            y=alt.Y("평균가:Q", title="원"),
            color=alt.Color(
                "구분:N",
                scale=alt.Scale(
                    domain=["금년", "전년"],
                    range=["#818CF8", "#64748B"],
                ),
            ),
            strokeDash=alt.StrokeDash(
                "구분:N",
                scale=alt.Scale(
                    domain=["금년", "전년"],
                    range=[[0], [6, 4]],
                ),
                legend=None,
            ),
            tooltip=["주차:O", "구분:N", alt.Tooltip("평균가:Q", format=",")],
        )
        .properties(height=320, title=f"{item_name} 전년 동기 비교")
        .interactive()
    )
    return chart


def _chart_compare(compare_data):
    """다중 품목 비교 grouped bar."""
    rows = []
    for r in compare_data:
        rows.append({"품목": r["품목"], "가격": r["평균가"], "유형": "평균가"})
        rows.append({"품목": r["품목"], "가격": r["최고가"], "유형": "최고가"})

    df = pd.DataFrame(rows)
    if df.empty:
        return None

    chart = (
        alt.Chart(df)
        .mark_bar(cornerRadiusEnd=4)
        .encode(
            x=alt.X("품목:N", sort=None, title=None),
            y=alt.Y("가격:Q", title="원"),
            color=alt.Color(
                "유형:N",
                scale=alt.Scale(
                    domain=["평균가", "최고가"],
                    range=["#818CF8", "#4ADE80"],
                ),
                legend=alt.Legend(orient="top"),
            ),
            xOffset="유형:N",
            tooltip=["품목:N", "유형:N", alt.Tooltip("가격:Q", format=",")],
        )
        .properties(height=350, title="품목별 경매가 비교")
    )
    return chart


# ---------------------------------------------------------------------------
# 데이터 집계 헬퍼
# ---------------------------------------------------------------------------
def _aggregate_day(rows):
    """하루치 API rows → {평균가, 최고가, 최저가, 건수}. rows=[] 이면 None."""
    valid = [r for r in rows if r["평균가"] > 0]
    if not valid:
        return None
    return {
        "평균가": round(sum(r["평균가"] for r in valid) / len(valid)),
        "최고가": max(r["최고가"] for r in valid),
        "최저가": min((r["최저가"] for r in valid if r["최저가"] > 0), default=0),
        "건수": len(valid),
    }


# ---------------------------------------------------------------------------
# Tab 렌더러
# ---------------------------------------------------------------------------
def _render_item_tab(get_date_str, get_date_compact):
    """Tab 1: 품목 조회."""
    col_cat, col_item, col_date = st.columns([1.2, 1.5, 1])
    with col_cat:
        category = st.selectbox(
            "카테고리", list(ITEM_CATEGORIES.keys()), key="auc_cat"
        )
    with col_item:
        items_in_cat = ITEM_CATEGORIES[category]
        selected = st.selectbox("품목", items_in_cat, key="auc_item")
    with col_date:
        target_date = st.date_input(
            "경매일", value=datetime.date.today(), key="auc_date"
        )

    col_q, col_spacer = st.columns([1, 3])
    with col_q:
        query_clicked = st.button("조회", type="primary", key="btn_auc_query")

    # 다중 비교
    with st.expander("다중 품목 비교", expanded=False):
        multi_items = st.multiselect(
            "품목 선택 (최대 5개)", ALL_ITEMS, default=[], key="auc_multi", max_selections=5
        )
        compare_clicked = st.button("비교 조회", key="btn_auc_compare")

    # 단일 조회
    if query_clicked:
        search = _get_search_term(selected)
        if search is None:
            st.warning(f"'{selected}'은(는) 경매 데이터가 없는 품목입니다 (수입/가공품).")
            return
        date_str = target_date.strftime("%Y%m%d")
        with st.spinner(f"'{selected}' 경매가 조회 중..."):
            data = _fetch_single_day(search, date_str)

        if not data:
            st.warning(
                f"'{selected}' {target_date.strftime('%m/%d')} 경매 데이터가 없습니다.\n\n"
                "가능한 원인: 경매 미진행 (휴일/주말), 검색어 불일치, API 서버 점검"
            )
            return

        # Metrics
        agg = _aggregate_day(data)
        if agg:
            c1, c2, c3 = st.columns(3)
            c1.metric(f"{selected} 최고가", f"{agg['최고가']:,}원")
            c2.metric(f"{selected} 평균가", f"{agg['평균가']:,}원")
            c3.metric("거래 건수", f"{agg['건수']}건")

        # 등급별 차트
        chart = _chart_grades(data, selected)
        if chart:
            st.altair_chart(chart, use_container_width=True)

        with st.expander("원시 데이터"):
            st.dataframe(pd.DataFrame(data), use_container_width=True)

    # 다중 비교
    if compare_clicked and multi_items:
        with st.spinner(f"{len(multi_items)}개 품목 비교 조회 중..."):
            compare_data = []
            for item in multi_items:
                search = _get_search_term(item)
                if search is None:
                    compare_data.append({"품목": item, "최고가": 0, "평균가": 0, "최저가": 0, "거래건수": 0})
                    continue
                date_str = target_date.strftime("%Y%m%d")
                rows = _fetch_single_day(search, date_str)
                agg = _aggregate_day(rows)
                if agg:
                    compare_data.append({
                        "품목": item,
                        "최고가": agg["최고가"],
                        "평균가": agg["평균가"],
                        "최저가": agg["최저가"],
                        "거래건수": agg["건수"],
                    })
                else:
                    compare_data.append({"품목": item, "최고가": 0, "평균가": 0, "최저가": 0, "거래건수": 0})

        df = pd.DataFrame(compare_data)
        st.dataframe(df, use_container_width=True)
        chart = _chart_compare([r for r in compare_data if r["평균가"] > 0])
        if chart:
            st.altair_chart(chart, use_container_width=True)

    # CLI 안내
    st.divider()
    with st.expander("단가시트 경매가 직접 입력 (CLI)"):
        cli_date = target_date.strftime("%Y-%m-%d")
        st.code(
            f'python execution/fill_auction_from_api.py '
            f'--master "G:/내 드라이브/1. 도크_주문 명세서/0. 매입단가_자료/도크발주관리데이터.xlsx" '
            f'--date {cli_date}',
            language="bash",
        )


def _render_trend_tab(get_date_str, get_date_compact):
    """Tab 2: 추이 분석."""
    col_item, col_period = st.columns([2, 1])
    with col_item:
        selected = st.selectbox("품목", ALL_ITEMS, key="auc_trend_item")
    with col_period:
        period = st.radio("기간", ["7일", "30일", "전년비교"], horizontal=True, key="auc_period")

    if not st.button("추이 조회", type="primary", key="btn_auc_trend"):
        return

    search = _get_search_term(selected)
    if search is None:
        st.warning(f"'{selected}'은(는) 경매 데이터가 없는 품목입니다.")
        return

    today = datetime.date.today()

    if period == "전년비교":
        _render_yoy(search, selected, today)
    else:
        days = 7 if period == "7일" else 30
        dates = [today - datetime.timedelta(days=i) for i in range(days)]

        progress = st.progress(0, text="데이터 수집 중...")
        raw = _fetch_multi_day(search, dates)
        progress.progress(100, text="완료")

        trend_data = []
        for d in sorted(dates):
            ds = d.strftime("%Y%m%d")
            rows = raw.get(ds, [])
            agg = _aggregate_day(rows)
            if agg:
                trend_data.append({
                    "날짜": d,
                    "평균가": agg["평균가"],
                    "최고가": agg["최고가"],
                    "최저가": agg["최저가"],
                })

        if not trend_data:
            st.warning(f"{period}간 데이터가 없습니다.")
            return

        # 차트
        chart = _chart_trend(trend_data, selected, period)
        if chart:
            st.altair_chart(chart, use_container_width=True)

        # 통계 Metrics
        avgs = [r["평균가"] for r in trend_data]
        maxs = [r["최고가"] for r in trend_data]
        mins = [r["최저가"] for r in trend_data if r["최저가"] > 0]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("기간 최고", f"{max(maxs):,}원")
        c2.metric("기간 최저", f"{min(mins):,}원" if mins else "-")
        c3.metric("평균", f"{round(sum(avgs)/len(avgs)):,}원")
        if len(avgs) >= 2 and avgs[0] > 0:
            change_pct = ((avgs[-1] - avgs[0]) / avgs[0]) * 100
            c4.metric("변동폭", f"{change_pct:+.1f}%")
        else:
            c4.metric("변동폭", "-")

        # 급등/급락 알림
        if len(avgs) >= 2:
            latest = avgs[-1]
            prev = avgs[-2]
            if prev > 0:
                daily_change = ((latest - prev) / prev) * 100
                if daily_change > 20:
                    st.error(f"가격 급등: {selected} 전일 대비 +{daily_change:.1f}%")
                elif daily_change < -20:
                    st.warning(f"가격 급락: {selected} 전일 대비 {daily_change:.1f}%")

        with st.expander("일별 데이터"):
            df = pd.DataFrame(trend_data)
            df["날짜"] = pd.to_datetime(df["날짜"]).dt.strftime("%m/%d")
            st.dataframe(df, use_container_width=True)


def _render_yoy(search, item_name, today):
    """전년비교: 금년 4주(28일 일별) + 전년 동기 (매주 수요일 4회)."""
    # 금년: 최근 28일
    this_dates = [today - datetime.timedelta(days=i) for i in range(28)]

    # 전년: 동기 4주, 수요일만 (7일 간격 × 4)
    one_year_ago = today.replace(year=today.year - 1)
    last_dates = []
    for w in range(4):
        d = one_year_ago - datetime.timedelta(days=w * 7)
        # 가장 가까운 수요일로 이동
        days_to_wed = (d.weekday() - 2) % 7
        d = d - datetime.timedelta(days=days_to_wed)
        last_dates.append(d)
    last_dates.sort()

    all_dates = this_dates + last_dates
    progress = st.progress(0, text="전년비교 데이터 수집 중...")
    raw = _fetch_multi_day(search, all_dates)
    progress.progress(100, text="완료")

    # 금년 주차별 집계
    this_weekly = []
    for w in range(4):
        week_start = 7 * (3 - w)
        week_dates = this_dates[week_start:week_start + 7]
        week_avgs = []
        for d in week_dates:
            rows = raw.get(d.strftime("%Y%m%d"), [])
            agg = _aggregate_day(rows)
            if agg:
                week_avgs.append(agg["평균가"])
        if week_avgs:
            this_weekly.append({"평균가": round(sum(week_avgs) / len(week_avgs))})
        else:
            this_weekly.append({"평균가": 0})

    # 전년 주차별
    last_weekly = []
    for d in last_dates:
        rows = raw.get(d.strftime("%Y%m%d"), [])
        agg = _aggregate_day(rows)
        if agg:
            last_weekly.append({"평균가": agg["평균가"]})
        else:
            last_weekly.append({"평균가": 0})

    # 차트
    chart = _chart_yoy(this_weekly, last_weekly, item_name)
    if chart:
        st.altair_chart(chart, use_container_width=True)

    # 금년/전년 비교 metrics
    this_avg = [r["평균가"] for r in this_weekly if r["평균가"] > 0]
    last_avg = [r["평균가"] for r in last_weekly if r["평균가"] > 0]

    c1, c2, c3 = st.columns(3)
    c1.metric("금년 평균", f"{round(sum(this_avg)/len(this_avg)):,}원" if this_avg else "-")
    c2.metric("전년 평균", f"{round(sum(last_avg)/len(last_avg)):,}원" if last_avg else "-")
    if this_avg and last_avg:
        yoy = ((sum(this_avg)/len(this_avg)) / (sum(last_avg)/len(last_avg)) - 1) * 100
        c3.metric("전년대비", f"{yoy:+.1f}%")
    else:
        c3.metric("전년대비", "-")


def _render_search_tab(get_date_str, get_date_compact):
    """Tab 3: 자유 검색."""
    col_term, col_date = st.columns([2, 1])
    with col_term:
        search_term = st.text_input("검색어 입력 (예: 딸기, 사과)", key="auc_free_search")
    with col_date:
        search_date = st.date_input("경매일", value=datetime.date.today(), key="auc_free_date")

    if not st.button("검색", type="primary", key="btn_auc_free"):
        return

    if not search_term.strip():
        st.warning("검색어를 입력해주세요.")
        return

    date_str = search_date.strftime("%Y%m%d")
    with st.spinner(f"'{search_term}' 검색 중..."):
        data = _fetch_single_day(search_term.strip(), date_str)

    if not data:
        st.warning(
            f"'{search_term}' {search_date.strftime('%m/%d')} 검색 결과가 없습니다.\n\n"
            "팁: 품목명을 짧게 입력해보세요 (예: '딸기', '사과', '배')"
        )
        return

    # Metrics
    agg = _aggregate_day(data)
    if agg:
        c1, c2, c3 = st.columns(3)
        c1.metric("최고가", f"{agg['최고가']:,}원")
        c2.metric("평균가", f"{agg['평균가']:,}원")
        c3.metric("거래 건수", f"{agg['건수']}건")

    # 등급별 차트
    chart = _chart_grades(data, search_term)
    if chart:
        st.altair_chart(chart, use_container_width=True)

    # 데이터 테이블
    st.dataframe(pd.DataFrame(data), use_container_width=True)


# ---------------------------------------------------------------------------
# Main render
# ---------------------------------------------------------------------------
def render(get_date_str=None, get_date_compact=None):
    """경매가 모니터링 페이지 메인."""
    _register_dok_theme()

    st.header("9. 경매가 모니터링")

    # API 상태
    api_ok = _check_api_health()
    if api_ok:
        st.caption("가락시장 경매가 | API 상태: :green[● 정상]")
    else:
        st.caption("가락시장 경매가 | API 상태: :red[● 연결 불가]")
        st.warning(
            "가락시장 API 서버에 연결할 수 없습니다. "
            "서버 점검 중이거나 인증 정보가 만료되었을 수 있습니다."
        )

    # 탭
    tab1, tab2, tab3 = st.tabs(["품목 조회", "추이 분석", "자유 검색"])

    with tab1:
        _render_item_tab(get_date_str, get_date_compact)
    with tab2:
        _render_trend_tab(get_date_str, get_date_compact)
    with tab3:
        _render_search_tab(get_date_str, get_date_compact)
