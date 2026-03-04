"""Page 3: Price Sheet Management (Stage 3, 3-1, 5)."""
import streamlit as st
from pathlib import Path
from dashboard.utils import (
    load_settings, run_script, mark_stage, get_stage_status,
    get_processed_orders_path, check_master_before_run,
    format_elapsed, get_timeout,
)


def render(get_date_str, get_date_compact):
    date_str = get_date_str()

    st.header("4. 단가표 관리")
    settings = load_settings()

    master_ok = check_master_before_run(settings)

    st.divider()

    # --- Three columns ---
    col1, col2, col3 = st.columns(3)

    # === 3: Price sheet creation ===
    with col1:
        st.subheader("단가표 생성")
        status_price = get_stage_status("stage3_price", date_str)
        if status_price == "completed":
            st.success("완료됨")

        processed = get_processed_orders_path(date_str)
        has_processed = processed.exists()
        if not has_processed:
            st.caption("⚠️ 처리된 발주 파일 필요")

        can_run_price = has_processed and master_ok
        if st.button("단가표 행 생성", type="primary", key="btn_create_price", disabled=not can_run_price):
            with st.spinner("단가표 생성 중..."):
                success, stdout, stderr, elapsed = run_script(
                    "create_price_sheet.py",
                    [
                        "--input", str(processed),
                        "--master", settings["master_file"],
                        "--date", date_str,
                    ],
                )
            if success:
                mark_stage("stage3_price", "completed", {"elapsed": elapsed}, date_str=date_str)
                st.toast(f"✅ 단가표 생성 완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                mark_stage("stage3_price", "failed", {"error": stderr[-500:]}, date_str=date_str)
                st.error("실패")
                st.code(stderr[-2000:], language="text")

    # === 3-1: Sikbom prices ===
    with col2:
        st.subheader("식봄 가격 조회")
        status_sikbom = get_stage_status("stage3_sikbom", date_str)
        if status_sikbom == "completed":
            st.success("완료됨")

        st.caption("Selenium 브라우저 자동 조회 (2~5분)")

        if st.button("식봄 가격 조회", type="primary", key="btn_sikbom", disabled=not master_ok):
            with st.spinner("식봄 가격 조회 중... (브라우저 자동화)"):
                success, stdout, stderr, elapsed = run_script(
                    "fill_sikbom_targeted.py",
                    ["--master", settings["master_file"], "--date", date_str],
                    timeout=get_timeout("sikbom"),
                )
            if success:
                mark_stage("stage3_sikbom", "completed", {"elapsed": elapsed}, date_str=date_str)
                st.toast(f"✅ 식봄 조회 완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                mark_stage("stage3_sikbom", "failed", {"error": stderr[-500:]}, date_str=date_str)
                st.error("실패")
                st.code(stderr[-2000:], language="text")
                if st.button("재시도", key="btn_sikbom_retry"):
                    st.rerun()

    # === 5: Auction prices ===
    with col3:
        st.subheader("경매가 조회")
        status_auction = get_stage_status("stage5_auction", date_str)
        if status_auction == "completed":
            st.success("완료됨")

        st.caption("가락시장 공공데이터 API")

        if st.button("경매가 조회", type="primary", key="btn_auction", disabled=not master_ok):
            with st.spinner("경매가 조회 중..."):
                success, stdout, stderr, elapsed = run_script(
                    "fill_auction_from_api.py",
                    ["--master", settings["master_file"], "--date", date_str],
                    timeout=get_timeout("auction"),
                )
            if success:
                mark_stage("stage5_auction", "completed", {"elapsed": elapsed}, date_str=date_str)
                st.toast(f"✅ 경매가 조회 완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                mark_stage("stage5_auction", "failed", {"error": stderr[-500:]}, date_str=date_str)
                st.error("실패")
                st.code(stderr[-2000:], language="text")
                if st.button("재시도", key="btn_auction_retry"):
                    st.rerun()

    st.divider()

    # --- Price sheet preview ---
    st.subheader("단가표 현황")
    _show_price_preview(settings["master_file"], date_str)


@st.cache_data(ttl=30, show_spinner="단가표 로딩 중...")
def _load_price_data(master_path, date_str):
    """Load price sheet data with caching (30s TTL)."""
    try:
        import openpyxl
        from datetime import datetime as dt

        p = Path(master_path)
        if not p.exists():
            return None, "파일 없음"

        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        if "단가" not in wb.sheetnames:
            wb.close()
            return None, "'단가' 시트 없음"

        ws = wb["단가"]
        target = dt.strptime(date_str, "%Y-%m-%d")

        rows_data = []
        for row in ws.iter_rows(min_row=2, max_col=12, values_only=True):
            cell_date = row[1]
            if cell_date is None:
                continue

            match = False
            if isinstance(cell_date, dt):
                match = cell_date.month == target.month and cell_date.day == target.day
            elif isinstance(cell_date, (int, float)):
                pass
            elif isinstance(cell_date, str):
                try:
                    if f"{target.month}월" in cell_date and f"{target.day}일" in cell_date:
                        match = True
                except Exception:
                    pass

            if match:
                rows_data.append({
                    "구분": str(row[0] or ""),
                    "품목": str(row[2] or ""),
                    "등급": str(row[3] or ""),
                    "단위": str(row[4] or ""),
                    "경매최고": row[5] if row[5] else None,
                    "경매평균": row[6] if row[6] else None,
                    "식봄가": row[7] if row[7] else None,
                    "매입가": row[8] if row[8] else None,
                })
        wb.close()
        return rows_data, None
    except PermissionError:
        return None, "마스터 파일이 열려 있습니다. 엑셀을 닫고 새로고침하세요."
    except Exception as e:
        return None, f"로드 실패: {e}"


def _show_price_preview(master_path, date_str):
    rows_data, error = _load_price_data(master_path, date_str)

    if error:
        st.warning(error)
        return

    if not rows_data:
        st.info(f"{date_str} 날짜의 단가 데이터가 없습니다.")
        return

    import pandas as pd
    df = pd.DataFrame(rows_data)

    # Stats
    price_cols = ["경매최고", "경매평균", "식봄가", "매입가"]
    empty_count = df[price_cols].isna().sum().sum()
    filled = df[price_cols].notna().sum().sum()
    total_cells = len(df) * len(price_cols)

    # Per-column completion
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("품목 수", len(df))

    for col_widget, col_name in zip([c2, c3, c4, c5], price_cols):
        col_filled = df[col_name].notna().sum()
        pct = (col_filled / len(df) * 100) if len(df) > 0 else 0
        col_widget.metric(col_name, f"{col_filled}/{len(df)}")

    # Overall completion bar
    pct_total = (filled / total_cells * 100) if total_cells > 0 else 0
    st.progress(pct_total / 100, text=f"가격 완성도: {pct_total:.0f}% ({filled}/{total_cells})")

    # Highlight empty cells with background color
    def highlight_empty(val):
        if val is None or (isinstance(val, (int, float)) and val == 0):
            return "background-color: #3B1C1C"
        return ""

    styled = df.style.map(highlight_empty, subset=price_cols)
    st.dataframe(styled, use_container_width=True, height=min(500, 35 * len(df) + 40))

    # Export
    col_r, col_e = st.columns([1, 3])
    with col_r:
        if st.button("새로고침", key="btn_refresh_price"):
            _load_price_data.clear()
            st.rerun()
    with col_e:
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("CSV 내보내기", data=csv, file_name=f"단가표_{date_str}.csv", mime="text/csv")
