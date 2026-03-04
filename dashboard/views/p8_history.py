"""Page 8: Execution History - Timeline of all script executions."""
import streamlit as st
import json
from dashboard.utils import (
    load_history, read_log, STAGE_MAP, HISTORY_DIR, LOGS_DIR,
    format_elapsed, today_str,
)


def render(get_date_str):
    date_str = get_date_str()

    st.header("실행 이력")

    tab1, tab2 = st.tabs(["타임라인", "전체 로그"])

    with tab1:
        _render_timeline(date_str)

    with tab2:
        _render_full_log(date_str)


def _render_timeline(date_str):
    """Show execution history as a timeline."""
    history = load_history(date_str)

    if not history:
        st.info(f"{date_str} 실행 이력이 없습니다.")
        return

    st.caption(f"{len(history)}건의 실행 기록")

    for entry in reversed(history):
        stage_id = entry.get("stage_id", "unknown")
        status = entry.get("status", "")
        ts = entry.get("timestamp", "")
        elapsed = entry.get("elapsed", 0)
        error = entry.get("error", "")

        stage_info = STAGE_MAP.get(stage_id, {})
        stage_name = stage_info.get("name", stage_id)

        # Time display
        short_ts = ts[11:19] if len(ts) > 19 else ts

        # Status icon
        if status == "completed":
            icon = "✅"
            color = "#4ADE80"
        elif status == "failed":
            icon = "❌"
            color = "#F87171"
        elif status == "running":
            icon = "⏳"
            color = "#FBBF24"
        else:
            icon = "⬜"
            color = "#64748B"

        # Card
        with st.container():
            c1, c2, c3 = st.columns([1, 3, 2])
            c1.markdown(f"`{short_ts}`")
            c2.markdown(f"{icon} **{stage_name}** — {status}")
            if elapsed:
                c3.caption(format_elapsed(elapsed))

            if error:
                with st.expander("오류 상세"):
                    st.code(error[:1000], language="text")

    # History file list
    st.divider()
    st.caption("이력 파일")
    if HISTORY_DIR.exists():
        files = sorted(HISTORY_DIR.glob("history_*.jsonl"), reverse=True)
        for f in files[:7]:
            line_count = sum(1 for _ in open(f, "r", encoding="utf-8"))
            st.markdown(f"- `{f.name}` ({line_count}건)")


def _render_full_log(date_str):
    """Show full execution log."""
    log_text = read_log(date_str, tail=500)

    if not log_text:
        st.info("로그가 없습니다.")
        return

    # Filter controls
    col1, col2 = st.columns([1, 1])
    with col1:
        level_filter = st.multiselect(
            "로그 레벨",
            ["INFO", "ERROR", "START", "END", "TIMEOUT"],
            default=["INFO", "ERROR", "START", "END"],
            key="log_level_filter",
        )
    with col2:
        search = st.text_input("검색", key="log_search", placeholder="키워드...")

    # Apply filters
    lines = log_text.split("\n")
    filtered = []
    for line in lines:
        if level_filter:
            if not any(f"[{lvl}]" in line for lvl in level_filter):
                # Also match START/END patterns
                if not any(lvl in line for lvl in level_filter):
                    continue
        if search and search.lower() not in line.lower():
            continue
        filtered.append(line)

    st.code("\n".join(filtered[-200:]), language="text")
    st.caption(f"표시: {len(filtered)}줄 / 전체: {len(lines)}줄")

    # Download log
    if LOGS_DIR.exists():
        date_compact = date_str.replace("-", "")
        log_path = LOGS_DIR / f"log_{date_compact}.txt"
        if log_path.exists():
            with open(log_path, "r", encoding="utf-8") as f:
                st.download_button(
                    "로그 다운로드",
                    data=f.read(),
                    file_name=log_path.name,
                    mime="text/plain",
                )
