"""Page 4: Purchase Price Entry & Selling Price Calculation (Stage 7)."""
import streamlit as st
from pathlib import Path
from dashboard.utils import (
    load_settings, run_script, mark_stage, get_stage_status,
    check_master_before_run, append_log,
    format_elapsed, STATE_DIR, PROJECT_ROOT,
    resolve_master_path, sync_master_back,
)
from dashboard.drive_service import is_cloud
import json
import os


def render(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    st.header("5. 매입가 입력 & 판매가 계산")
    settings = load_settings()

    master_ok = check_master_before_run(settings)
    cloud_mode = is_cloud()

    # Resolve master path once
    master_path, from_drive = resolve_master_path(settings)

    st.divider()

    # --- Receipt auto-input ---
    if cloud_mode:
        st.subheader("영수증 자동 입력")
        st.info("Cloud 환경에서는 영수증 자동 입력이 지원되지 않습니다. (로컬 파일 시스템 필요)")
    else:
        _render_receipt_section(settings, date_str, master_ok)

    st.divider()

    # --- Editable price table ---
    st.subheader("매입가 미입력 품목")
    if master_ok:
        _show_and_edit_prices(master_path, date_str, from_drive=from_drive)
    else:
        st.caption("마스터 파일 잠금 해제 후 매입가 편집 가능합니다.")

    st.divider()

    # --- Fill prices ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("판매단가 자동 계산")
        status_fill = get_stage_status("stage7_fill", date_str)
        if status_fill == "completed":
            st.success("완료됨")

        if st.button("판매단가 계산 실행", type="primary", key="btn_fill_prices", disabled=not master_ok):
            with st.spinner("매입가/판매가 계산 중..."):
                success, stdout, stderr, elapsed = run_script(
                    "fill_prices.py",
                    ["--date", date_str, "--master-file", master_path],
                )
            if success:
                if from_drive:
                    sync_master_back(master_path)
                mark_stage("stage7_fill", "completed", {"elapsed": elapsed}, date_str=date_str)
                st.toast(f"판매단가 계산 완료 ({format_elapsed(elapsed)})")
                if stdout:
                    with st.expander("실행 결과"):
                        st.code(stdout[-3000:], language="text")
            else:
                mark_stage("stage7_fill", "failed", {"error": stderr[-500:]}, date_str=date_str)
                st.error("실패")
                st.code(stderr[-2000:], language="text")

    with col2:
        st.subheader("가격 미리보기")
        if st.button("발주시트 가격 현황 조회", key="btn_preview_prices"):
            _show_order_prices(master_path, date_str)


# ==================== 영수증 자동 입력 섹션 ====================

def _render_receipt_section(settings, date_str, master_ok):
    """영수증 빠른 입력 + 자동 감시 데몬 상태."""
    st.subheader("영수증 자동 입력")

    tab_quick, tab_watcher = st.tabs(["빠른 입력", "자동 감시 데몬"])

    with tab_quick:
        _render_quick_receipt(settings, date_str, master_ok)

    with tab_watcher:
        _render_watcher_control(settings, date_str)


def _render_quick_receipt(settings, date_str, master_ok):
    """URL/이미지/텍스트로 빠른 영수증 입력."""
    input_type = st.radio(
        "입력 유형", ["URL", "이미지 업로드", "텍스트"],
        horizontal=True, key="receipt_input_type",
    )

    dry_run = st.checkbox("미리보기만 (dry-run)", value=True, key="receipt_dry_run")

    if input_type == "URL":
        url = st.text_input("영수증 URL (itanet / marketbom)", key="receipt_url",
                            placeholder="https://www.itanet.co.kr/...")
        supplier = st.text_input("공급처 (자동감지, 비워도 됨)", key="receipt_supplier_url")
        if st.button("URL 영수증 입력", key="btn_receipt_url", disabled=not master_ok or not url):
            args = ["--master", settings["master_file"], "--date", date_str, "--url", url]
            if supplier:
                args += ["--supplier", supplier]
            if dry_run:
                args.append("--dry-run")
            with st.spinner("URL 파싱 + 입력 중..."):
                success, stdout, stderr, elapsed = run_script("fill_receipt_prices.py", args)
            _show_receipt_result(success, stdout, stderr, elapsed)

    elif input_type == "이미지 업로드":
        uploaded = st.file_uploader("영수증 이미지", type=["jpg", "jpeg", "png"], key="receipt_img")
        supplier = st.text_input("공급처 (자동감지, 비워도 됨)", key="receipt_supplier_img")
        if st.button("이미지 영수증 입력", key="btn_receipt_img", disabled=not master_ok or not uploaded):
            # 임시 파일 저장
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), f"receipt_upload_{uploaded.name}")
            with open(tmp, "wb") as f:
                f.write(uploaded.getvalue())
            args = ["--master", settings["master_file"], "--date", date_str, "--image", tmp]
            if supplier:
                args += ["--supplier", supplier]
            if dry_run:
                args.append("--dry-run")
            with st.spinner("이미지 OCR + 입력 중... (10~15초)"):
                success, stdout, stderr, elapsed = run_script("fill_receipt_prices.py", args)
            _show_receipt_result(success, stdout, stderr, elapsed)

    elif input_type == "텍스트":
        text = st.text_area("영수증 텍스트 (품목 가격 형식)", key="receipt_text",
                            placeholder="감자 62000\n당근 35000\n...")
        supplier = st.text_input("공급처 (필수)", key="receipt_supplier_text")
        if st.button("텍스트 영수증 입력", key="btn_receipt_text",
                      disabled=not master_ok or not text or not supplier):
            args = ["--master", settings["master_file"], "--date", date_str,
                    "--text", text, "--supplier", supplier]
            if dry_run:
                args.append("--dry-run")
            with st.spinner("텍스트 파싱 + 입력 중..."):
                success, stdout, stderr, elapsed = run_script("fill_receipt_prices.py", args)
            _show_receipt_result(success, stdout, stderr, elapsed)


def _show_receipt_result(success, stdout, stderr, elapsed):
    """영수증 처리 결과 표시."""
    if success:
        st.success(f"처리 완료 ({format_elapsed(elapsed)})")
        if stdout:
            # 결과 요약 추출
            lines = stdout.strip().split('\n')
            summary_lines = [l for l in lines if any(k in l for k in ['[OK]', '[DRY]', '[SKIP]', '[MISS]', '[WARN]', '입력:', '결과'])]
            if summary_lines:
                st.code('\n'.join(summary_lines[-20:]), language="text")
            with st.expander("전체 로그"):
                st.code(stdout[-3000:], language="text")
    else:
        st.error("처리 실패")
        st.code(stderr[-2000:] if stderr else "Unknown error", language="text")


def _render_watcher_control(settings, date_str):
    """영수증 자동 감시 데몬 상태 + 실행 안내."""
    st.markdown("""
**영수증 자동 감시 데몬** (`receipt_watcher.py`)

카카오톡 다운로드 폴더와 클립보드를 실시간 감시하여,
영수증을 자동으로 파싱 → 단가시트에 입력합니다.

| 트리거 | 사용자 동작 | 대상 |
|--------|-----------|------|
| 폴더 감시 | 카톡 이미지 클릭 | POS/손글씨 ~15개 공급처 |
| 클립보드 | URL 복사 | 건영농산, 오복상회 |
| 클립보드 | 텍스트 복사 | 텍스트 영수증 |
""")

    col_cmd, col_status = st.columns([2, 1])

    with col_cmd:
        st.markdown("**실행 명령어:**")
        mode = "--dry-run " if st.checkbox("Dry-run", value=False, key="watcher_dry") else ""
        cmd = f"python execution/receipt_watcher.py --date {date_str} {mode}--console"
        st.code(cmd, language="bash")
        st.caption("별도 터미널에서 실행하세요. (대시보드와 독립 실행)")

    with col_status:
        st.markdown("**감시 폴더:**")
        kakao_dir = os.environ.get('KAKAO_DOWNLOAD_DIR', '')
        if kakao_dir and os.path.isdir(kakao_dir):
            st.success(f"감지됨")
            st.caption(kakao_dir)
        else:
            # 자동 탐지 시도
            from execution.receipt_watcher import find_kakao_download_dir
            detected = find_kakao_download_dir()
            if detected:
                st.success("자동 감지됨")
                st.caption(detected)
            else:
                st.warning("미감지")
                st.caption(".env KAKAO_DOWNLOAD_DIR 설정 필요")


def _show_and_edit_prices(master_path, date_str, from_drive=False):
    """Load price sheet, show empty-price items in data_editor, allow save."""
    try:
        import openpyxl
        import pandas as pd
        from datetime import datetime as dt

        p = Path(master_path)
        if not p.exists():
            st.info("마스터 파일 없음")
            return

        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        if "단가" not in wb.sheetnames:
            wb.close()
            return

        ws = wb["단가"]
        target = dt.strptime(date_str, "%Y-%m-%d")

        all_rows = []
        empty_rows = []
        row_map = {}  # df_index -> excel_row_number

        for idx, row in enumerate(ws.iter_rows(min_row=2, max_col=12, values_only=True), start=2):
            cell_date = row[1]
            if cell_date is None:
                continue

            match = False
            if isinstance(cell_date, dt):
                match = cell_date.month == target.month and cell_date.day == target.day
            elif isinstance(cell_date, str):
                if f"{target.month}월" in str(cell_date) and f"{target.day}일" in str(cell_date):
                    match = True

            if match:
                entry = {
                    "품목": str(row[2] or ""),
                    "등급": str(row[3] or ""),
                    "단위": str(row[4] or ""),
                    "경매최고": row[5] if row[5] else 0,
                    "경매평균": row[6] if row[6] else 0,
                    "식봄가": row[7] if row[7] else 0,
                    "매입가": row[8] if row[8] else 0,
                }
                all_rows.append(entry)
                purchase = row[8]
                if not purchase or purchase == 0:
                    empty_rows.append(entry)
                    row_map[len(empty_rows) - 1] = idx

        wb.close()

        if not all_rows:
            st.info(f"{date_str} 단가 데이터 없음")
            return

        filled = len(all_rows) - len(empty_rows)
        st.progress(filled / len(all_rows) if all_rows else 0, text=f"매입가 입력: {filled}/{len(all_rows)}")

        if not empty_rows:
            st.success("모든 품목의 매입가가 입력되어 있습니다!")
            return

        df = pd.DataFrame(empty_rows)

        edited = st.data_editor(
            df,
            column_config={
                "품목": st.column_config.TextColumn("품목", disabled=True, width="medium"),
                "등급": st.column_config.TextColumn("등급", disabled=True, width="small"),
                "단위": st.column_config.TextColumn("단위", disabled=True, width="small"),
                "경매최고": st.column_config.NumberColumn("경매최고", disabled=True, format="%d"),
                "경매평균": st.column_config.NumberColumn("경매평균", disabled=True, format="%d"),
                "식봄가": st.column_config.NumberColumn("식봄가", disabled=True, format="%d"),
                "매입가": st.column_config.NumberColumn(
                    "매입가 (입력)",
                    format="%d",
                    min_value=0,
                    help="매입가를 직접 입력하세요",
                ),
            },
            use_container_width=True,
            num_rows="fixed",
            key="price_editor",
        )

        # Save to persistent file for cross-page persistence
        _save_edit_mapping(date_str, row_map)

        # Count changes
        changes = sum(1 for i in range(len(edited)) if edited.iloc[i]["매입가"] != df.iloc[i]["매입가"] and edited.iloc[i]["매입가"] > 0)

        col_save, col_info = st.columns([1, 2])
        with col_save:
            if st.button(f"매입가 저장 ({changes}건)", type="primary", key="btn_save_prices", disabled=changes == 0):
                _do_save(master_path, date_str, edited, row_map, from_drive=from_drive)
        with col_info:
            if changes > 0:
                st.caption(f"변경된 {changes}건이 저장됩니다.")

    except PermissionError:
        st.error("마스터 파일이 열려 있습니다. 엑셀을 닫아주세요.")
    except Exception as e:
        st.warning(f"로드 실패: {e}")


def _save_edit_mapping(date_str, row_map):
    """Persist row index mapping for reliability."""
    mapping_file = STATE_DIR / f"price_edit_map_{date_str.replace('-', '')}.json"
    with open(mapping_file, "w") as f:
        json.dump({str(k): v for k, v in row_map.items()}, f)


def _do_save(master_path, date_str, edited_df, row_map, from_drive=False):
    """Save edited purchase prices to master file."""
    try:
        import openpyxl

        wb = openpyxl.load_workbook(str(master_path))
        ws = wb["단가"]

        saved = 0
        for df_idx in range(len(edited_df)):
            price = edited_df.iloc[df_idx]["매입가"]
            if price and price > 0 and df_idx in row_map:
                row_num = row_map[df_idx]
                ws.cell(row=row_num, column=9, value=int(price))  # Col I = 9 (1-based)
                saved += 1

        if saved > 0:
            wb.save(str(master_path))
            append_log(f"매입가 {saved}건 저장됨 (단가시트)")
            # Cloud mode: upload modified file back to Drive
            if from_drive:
                sync_master_back(master_path)
            st.success(f"매입가 {saved}건 저장 완료!")
        else:
            st.info("저장할 변경사항이 없습니다.")
        wb.close()

    except PermissionError:
        st.error("마스터 파일이 열려 있습니다. 엑셀을 닫아주세요.")
    except Exception as e:
        st.error(f"저장 실패: {e}")


def _show_order_prices(master_path, date_str):
    """Show order sheet with filled prices."""
    try:
        import openpyxl
        import pandas as pd
        from datetime import datetime as dt

        p = Path(master_path)
        if not p.exists():
            return

        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        if "발주" not in wb.sheetnames:
            wb.close()
            return

        ws = wb["발주"]
        target = dt.strptime(date_str, "%Y-%m-%d")

        rows_data = []
        # 발주시트 컬럼: A(0)=매장, D(3)=날짜, E(4)=품목, F(5)=공급업체,
        #               G(6)=단위, H(7)=수량, I(8)=매입단가, K(10)=판매단가
        for row in ws.iter_rows(min_row=2, max_col=14, values_only=True):
            cell_date = row[3]  # D열: 날짜
            if cell_date is None:
                continue

            match = False
            if isinstance(cell_date, dt):
                match = cell_date.month == target.month and cell_date.day == target.day and cell_date.year == target.year
            elif hasattr(cell_date, 'strftime'):
                try:
                    match = cell_date.strftime('%Y-%m-%d') == target.strftime('%Y-%m-%d')
                except Exception:
                    pass
            elif isinstance(cell_date, str):
                if f"{target.month}월" in str(cell_date) and f"{target.day}일" in str(cell_date):
                    match = True

            if match:
                purchase = row[8] if row[8] else 0   # I열: 매입단가
                selling = row[10] if row[10] else 0   # K열: 판매단가
                margin = 0
                if selling and purchase and selling > 0:
                    margin = ((selling - purchase) / selling) * 100

                rows_data.append({
                    "매장": str(row[0] or ""),        # A열
                    "품목": str(row[4] or ""),         # E열
                    "공급업체": str(row[5] or ""),     # F열
                    "단위": str(row[6] or ""),         # G열
                    "수량": row[7] or "",              # H열
                    "매입가": int(purchase) if purchase else 0,
                    "판매가": int(selling) if selling else 0,
                    "마진(%)": round(margin, 1),
                })

        wb.close()

        if rows_data:
            df = pd.DataFrame(rows_data)

            # Highlight low margins
            def color_margin(val):
                if isinstance(val, (int, float)):
                    if val < 0:
                        return "color: #F87171; font-weight: bold"
                    elif val < 10:
                        return "color: #FBBF24"
                    elif val > 50:
                        return "color: #4ADE80; font-weight: bold"
                return ""

            # 표시 컬럼 순서 정리
            display_cols = ["매장", "품목", "공급업체", "단위", "수량", "매입가", "판매가", "마진(%)"]
            display_df = df[[c for c in display_cols if c in df.columns]]

            styled = display_df.style.map(color_margin, subset=["마진(%)"])
            st.dataframe(styled, use_container_width=True, height=400)

            # Summary metrics
            c_m1, c_m2, c_m3, c_m4 = st.columns(4)
            total_purchase = df[df["매입가"] > 0]["매입가"].sum()
            total_selling = df[df["판매가"] > 0]["판매가"].sum()
            avg_margin = df[df["마진(%)"] > 0]["마진(%)"].mean()
            low_margin = df[df["마진(%)"] < 15]

            c_m1.metric("총 매입액", f"{total_purchase:,.0f}원")
            c_m2.metric("총 판매액", f"{total_selling:,.0f}원")
            c_m3.metric("평균 마진", f"{avg_margin:.1f}%" if avg_margin > 0 else "N/A")
            c_m4.metric("저마진 품목", f"{len(low_margin)}건")

            if len(low_margin) > 0:
                st.warning(f"저마진 품목: {', '.join(low_margin['품목'].tolist()[:10])}")

            # Margin distribution chart
            valid_margins = df[df["마진(%)"] > 0]
            if len(valid_margins) > 3:
                st.subheader("마진율 분포")
                chart_df = valid_margins[["품목", "마진(%)"]].set_index("품목").sort_values("마진(%)")
                st.bar_chart(chart_df, use_container_width=True)

            # Store-level summary
            if "매장" in df.columns:
                store_summary = df.groupby("매장").agg(
                    품목수=("품목", "count"),
                    평균마진=("마진(%)", "mean"),
                    총매입=("매입가", "sum"),
                ).round(1)
                if len(store_summary) > 1:
                    st.subheader("매장별 요약")
                    st.dataframe(store_summary, use_container_width=True)
        else:
            st.info(f"{date_str} 발주 데이터 없음")

    except PermissionError:
        st.error("마스터 파일이 열려 있습니다.")
    except Exception as e:
        st.warning(f"미리보기 실패: {e}")
