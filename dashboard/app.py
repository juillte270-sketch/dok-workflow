"""
DoK Workflow Dashboard - Main Entry Point
Run: streamlit run dashboard/app.py
"""
import streamlit as st
import sys
import datetime
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from dashboard.utils import (
    today_str, count_completed, load_progress, STAGES,
    load_settings, render_stage_badge,
)
from dashboard.drive_service import is_cloud

st.set_page_config(
    page_title="DoK Workflow",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- Design System: Slate Indigo Palette ---
# Primary: #818CF8 (Indigo-400)  |  Surface: #1E293B (Slate-800)
# BG: #0F172A (Slate-900)        |  Text: #F1F5F9 (Slate-100)
# Success: #4ADE80  Error: #F87171  Running: #FBBF24  Muted: #64748B
# See: .claude/skills/design-guide/SKILL.md

st.markdown("""
<style>
    /* === Base Layout — 토스 스타일: 넉넉한 여백 === */
    .block-container {
        padding-top: 2.5rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 100% !important;
    }
    hr { margin: 1rem 0 !important; border-color: #334155; }

    /* 전체 폰트 기본 크기 확보 */
    html, body, [class*="css"] {
        font-size: 16px !important;
    }

    /* === Sidebar === */
    section[data-testid="stSidebar"] {
        min-width: 280px;
    }
    section[data-testid="stSidebar"] > div { padding-top: 0.75rem; }
    section[data-testid="stSidebar"] .stRadio > label {
        font-size: 0.9rem;
        letter-spacing: 0.01em;
    }
    section[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label {
        padding: 0.4rem 0.6rem;
        border-radius: 6px;
        transition: background 0.15s ease;
        font-size: 0.9rem;
    }
    section[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label:hover {
        background: rgba(129, 140, 248, 0.08);
    }
    section[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label[data-checked="true"] {
        background: rgba(129, 140, 248, 0.15);
    }

    /* === 토스 스타일 카드 컨테이너 === */
    .tok-card {
        background: #1E293B;
        border: 1px solid #334155;
        border-radius: 14px;
        padding: 20px 24px;
        margin-bottom: 12px;
    }
    .tok-card-title {
        font-size: 0.78rem;
        font-weight: 600;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 12px;
    }

    /* === 상태 배지 (컬러 도트 + 텍스트) === */
    .badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        font-size: 0.82rem;
        font-weight: 500;
    }
    .badge .dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        display: inline-block;
    }
    .dot-ok   { background: #4ADE80; }
    .dot-fail { background: #F87171; }
    .dot-run  { background: #FBBF24; animation: pulse 1.5s ease-in-out infinite; }
    .dot-wait { background: #475569; }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.4; }
    }

    /* === 스텝 번호 원형 배지 === */
    .step-num {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: #334155;
        color: #E2E8F0;
        font-size: 0.78rem;
        font-weight: 600;
    }
    .step-num-sm {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 20px;
        height: 20px;
        border-radius: 50%;
        background: #334155;
        color: #CBD5E1;
        font-size: 0.65rem;
        font-weight: 600;
    }

    /* === Metrics Cards === */
    [data-testid="stMetric"] {
        background: #1E293B;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 16px 20px;
    }
    [data-testid="stMetricLabel"] { font-size: 0.85rem !important; color: #94A3B8 !important; }
    [data-testid="stMetricValue"] { font-size: 1.8rem !important; font-weight: 700 !important; }

    /* === Buttons === */
    .stButton > button {
        width: 100%;
        margin-bottom: 4px;
        border-radius: 8px;
        font-weight: 500;
        font-size: 0.9rem;
        padding: 0.5rem 1rem;
        letter-spacing: 0.01em;
        transition: all 0.15s ease;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #818CF8, #6366F1);
        border: none;
        font-weight: 600;
        box-shadow: 0 2px 8px rgba(99, 102, 241, 0.25);
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #6366F1, #4F46E5);
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.35);
    }
    .stButton > button[kind="secondary"] {
        border: 1px solid #334155;
        background: transparent;
    }
    .stButton > button[kind="secondary"]:hover {
        border-color: #818CF8;
        background: rgba(129, 140, 248, 0.06);
    }

    /* === Tabs === */
    .stTabs [data-baseweb="tab-list"] { gap: 2px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0 0;
        padding: 10px 20px;
        font-size: 0.9rem;
    }

    /* === Expanders === */
    .streamlit-expanderHeader {
        font-size: 0.95rem;
        font-weight: 500;
        border-radius: 8px;
    }

    /* === Data Tables === */
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

    /* === Code / Log area === */
    .log-area {
        font-family: 'Consolas', 'SF Mono', 'Fira Code', monospace;
        font-size: 0.8rem;
        line-height: 1.5;
    }
    code {
        font-size: 0.85rem !important;
    }

    /* === Progress bar === */
    .stProgress > div > div {
        background: linear-gradient(90deg, #818CF8, #6366F1) !important;
        border-radius: 4px;
    }

    /* === Toast / Alerts === */
    .stAlert { border-radius: 8px; font-size: 0.9rem; }

    /* === Headers — 토스 스타일: 깔끔한 위계 === */
    h1 { font-size: 1.5rem !important; font-weight: 700 !important; letter-spacing: -0.02em; margin-bottom: 0.2rem !important; }
    h2 { font-size: 1.2rem !important; font-weight: 600 !important; color: #E2E8F0 !important; }
    h3 { font-size: 1.1rem !important; font-weight: 600 !important; color: #CBD5E1 !important; }

    /* === Sidebar Progress & Stage List === */
    section[data-testid="stSidebar"] .stProgress > div > div {
        height: 6px !important;
    }
    section[data-testid="stSidebar"] p {
        font-size: 0.88rem;
        line-height: 1.4;
    }
    .sidebar-stage {
        padding: 6px 10px;
        border-radius: 8px;
        margin-bottom: 2px;
        font-size: 0.84rem;
        line-height: 1.5;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .sidebar-stage-name {
        flex: 1;
        color: #E2E8F0;
    }

    /* === Text area / Input 크기 확보 === */
    .stTextArea textarea { font-size: 0.9rem !important; }
    .stSelectbox, .stDateInput { font-size: 0.9rem; }

    /* === caption 크기 키움 === */
    .stCaption, [data-testid="stCaptionContainer"] {
        font-size: 0.82rem !important;
    }

    /* === Mobile Responsive — 768px 이하 === */
    @media (max-width: 768px) {
        /* 컨테이너 패딩 축소 */
        .block-container {
            padding-top: 1rem !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-bottom: 1rem !important;
        }

        /* 사이드바 최소 너비 해제 */
        section[data-testid="stSidebar"] {
            min-width: unset !important;
        }

        /* 헤더 크기 축소 */
        h1 { font-size: 1.25rem !important; }
        h2 { font-size: 1.05rem !important; }
        h3 { font-size: 0.95rem !important; }

        /* 메트릭 카드 — 콤팩트 */
        [data-testid="stMetric"] {
            padding: 10px 12px !important;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.3rem !important;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.78rem !important;
        }

        /* 버튼 — 터치 친화적 (최소 44px) */
        .stButton > button {
            padding: 0.6rem 0.75rem !important;
            font-size: 0.85rem !important;
            min-height: 44px !important;
        }

        /* 탭 — 콤팩트 */
        .stTabs [data-baseweb="tab"] {
            padding: 8px 12px !important;
            font-size: 0.82rem !important;
        }

        /* 카드 — 패딩 축소 */
        .tok-card {
            padding: 14px 16px !important;
            border-radius: 10px !important;
        }

        /* 스텝 번호 — 약간 축소 */
        .step-num {
            width: 24px !important;
            height: 24px !important;
            font-size: 0.7rem !important;
        }

        /* 코드/로그 영역 — 가로 스크롤 허용 */
        pre, code, .log-area {
            font-size: 0.72rem !important;
            word-break: break-all !important;
        }

        /* 데이터프레임 가로 스크롤 */
        [data-testid="stDataFrame"] {
            overflow-x: auto !important;
        }

        /* 사이드바 라디오 — 터치 영역 확대 */
        section[data-testid="stSidebar"] .stRadio > div[role="radiogroup"] > label {
            padding: 0.55rem 0.8rem !important;
            font-size: 0.88rem !important;
        }

        /* 캡션 — 가독성 확보 */
        .stCaption, [data-testid="stCaptionContainer"] {
            font-size: 0.78rem !important;
        }

        /* 배지 — 약간 축소 */
        .badge { font-size: 0.75rem !important; }

        /* 텍스트 영역 / 입력 — 풀 너비 */
        .stTextArea textarea,
        .stSelectbox,
        .stDateInput {
            font-size: 0.85rem !important;
        }
    }

    /* === 초소형 모바일 — 480px 이하 === */
    @media (max-width: 480px) {
        .block-container {
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
        }

        h1 { font-size: 1.1rem !important; }

        [data-testid="stMetricValue"] {
            font-size: 1.1rem !important;
        }

        /* 사이드바 내 타이틀 축소 */
        section[data-testid="stSidebar"] h3 {
            font-size: 0.9rem !important;
        }
    }
</style>
""", unsafe_allow_html=True)


# --- Shared date in session state (synced across pages) ---
if "target_date" not in st.session_state:
    st.session_state.target_date = datetime.date.today()


# --- Sidebar ---
with st.sidebar:
    st.markdown("### DoK Workflow")
    # Mode badge
    if is_cloud():
        st.markdown(
            '<span style="background:#334155; color:#FBBF24; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600;">CLOUD</span>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<span style="background:#334155; color:#4ADE80; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600;">LOCAL</span>',
            unsafe_allow_html=True,
        )
    st.caption(f"오늘: {today_str('%Y년 %m월 %d일')}")

    # Global date picker (synced across all pages)
    new_date = st.date_input(
        "작업 날짜",
        value=st.session_state.target_date,
        key="global_date_picker",
    )
    if new_date != st.session_state.target_date:
        st.session_state.target_date = new_date

    date_str = st.session_state.target_date.strftime("%Y-%m-%d")

    # Progress summary
    completed = count_completed(date_str)
    total = len(STAGES)
    pct = completed / total if total else 0
    st.markdown(
        f'<div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:4px;">'
        f'<span style="font-size:0.82rem; color:#94A3B8;">진행률</span>'
        f'<span style="font-size:1.1rem; font-weight:700; color:#E2E8F0;">{completed}<span style="font-size:0.82rem; color:#64748B; font-weight:400">/{total}</span></span>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.progress(pct)

    st.divider()

    # Page → stages mapping (for status display)
    progress = load_progress(date_str)
    _page_stage_map = {}
    for s in STAGES:
        pg = s.get("page", "")
        _page_stage_map.setdefault(pg, []).append(s["id"])

    _page_options = [
        "메인 대시보드",
        "1. 발주 정규화",
        "2. 발주 처리",
        "3. 발주시트 & 명세서",
        "4. 단가표 관리",
        "5. 매입가 & 판매가",
        "6. 최종 명세서",
        "7. 동원발주 스크래핑",
        "8. 매입가 입력 (영수증)",
        "9. 경매가 모니터링",
        "10. 실행 이력",
        "11. 설정",
        "12. 배송 관리",
        "13. 일일보고서",
    ]

    def _fmt_page(name):
        stage_ids = _page_stage_map.get(name, [])
        if not stage_ids:
            return name
        done = sum(1 for sid in stage_ids if progress.get(sid, {}).get("status") == "completed")
        failed = sum(1 for sid in stage_ids if progress.get(sid, {}).get("status") == "failed")
        total = len(stage_ids)
        if done == total:
            return f"{name}  ✓"
        elif failed > 0:
            return f"{name}  ✗"
        elif done > 0:
            return f"{name}  ({done}/{total})"
        return name

    page = st.radio(
        "메뉴",
        _page_options,
        format_func=_fmt_page,
        label_visibility="collapsed",
    )


# --- Page Router ---
def get_date_str():
    return st.session_state.target_date.strftime("%Y-%m-%d")

def get_date_compact():
    return st.session_state.target_date.strftime("%Y%m%d")


if page == "메인 대시보드":
    from dashboard.views import p0_main
    p0_main.render(get_date_str, get_date_compact)
elif page == "1. 발주 정규화":
    from dashboard.views import p0b_normalize
    p0b_normalize.render(get_date_str, get_date_compact)
elif page == "2. 발주 처리":
    from dashboard.views import p1_orders
    p1_orders.render(get_date_str, get_date_compact)
elif page == "3. 발주시트 & 명세서":
    from dashboard.views import p2_sheet_invoice
    p2_sheet_invoice.render(get_date_str, get_date_compact)
elif page == "4. 단가표 관리":
    from dashboard.views import p3_price_sheet
    p3_price_sheet.render(get_date_str, get_date_compact)
elif page == "5. 매입가 & 판매가":
    from dashboard.views import p4_purchase_price
    p4_purchase_price.render(get_date_str, get_date_compact)
elif page == "6. 최종 명세서":
    from dashboard.views import p5_final_invoice
    p5_final_invoice.render(get_date_str, get_date_compact)
elif page == "7. 동원발주 스크래핑":
    from dashboard.views import p4a_helo_second
    p4a_helo_second.render(get_date_str, get_date_compact)
elif page == "8. 매입가 입력 (영수증)":
    from dashboard.views import p4b_receipt
    p4b_receipt.render(get_date_str, get_date_compact)
elif page == "9. 경매가 모니터링":
    from dashboard.views import p6_auction
    p6_auction.render(get_date_str, get_date_compact)
elif page == "10. 실행 이력":
    from dashboard.views import p8_history
    p8_history.render(get_date_str)
elif page == "11. 설정":
    from dashboard.views import p7_settings
    p7_settings.render()
elif page == "12. 배송 관리":
    from dashboard.views import p9_delivery
    p9_delivery.render(get_date_str, get_date_compact)
elif page == "13. 일일보고서":
    from dashboard.views import p10_report
    p10_report.render(get_date_str, get_date_compact)
