"""Page 2: Order Sheet & Invoice Generation (Stage 2a, 2b)."""
import streamlit as st
import zipfile
import io
from pathlib import Path
from dashboard.utils import (
    load_settings, run_script, mark_stage, get_stage_status,
    get_processed_orders_path, OUTPUTS_DIR, check_master_before_run,
    format_elapsed,
)


def render(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    st.header("3. 발주시트 & 명세서")
    settings = load_settings()

    processed = get_processed_orders_path(date_str)
    if not processed.exists():
        st.warning(f"처리된 발주 파일이 없습니다: {processed.name}")

    st.divider()

    # --- Two-column layout ---
    col1, col2 = st.columns(2)

    # === 2a: Order Sheet ===
    with col1:
        st.subheader("발주시트 입력")
        status_order = get_stage_status("stage2_order", date_str)
        if status_order == "completed":
            st.success("완료됨")

        master_ok = check_master_before_run(settings)
        if st.button("발주시트 업데이트", type="primary", key="btn_update_order", disabled=not master_ok):
            if master_ok:
                with st.spinner("발주시트 업데이트 중..."):
                    success, stdout, stderr, elapsed = run_script(
                        "update_order_sheet.py",
                        ["--date", date_str, "--master-file", settings["master_file"]],
                    )
                if success:
                    mark_stage("stage2_order", "completed", {"elapsed": elapsed}, date_str=date_str)
                    st.toast(f"✅ 발주시트 완료 ({format_elapsed(elapsed)})")
                    if stdout:
                        with st.expander("실행 결과"):
                            st.code(stdout[-2000:], language="text")
                else:
                    mark_stage("stage2_order", "failed", {"error": stderr[-500:]}, date_str=date_str)
                    st.error("실패")
                    st.code(stderr[-2000:], language="text")

    # === 2b: Invoice Generation ===
    with col2:
        st.subheader("가명세서 생성")
        status_inv = get_stage_status("stage2_invoice", date_str)
        if status_inv == "completed":
            st.success("완료됨")

        upload_to_drive = st.checkbox("Drive에 업로드", value=False, key="s2_drive_upload")

        if st.button("가명세서 생성", type="primary", key="btn_gen_invoice"):
            args = ["--date", date_str]
            if not upload_to_drive:
                args.append("--local-only")
            with st.spinner("가명세서 생성 중..."):
                success, stdout, stderr, elapsed = run_script("generate_invoices.py", args)
            if success:
                mark_stage("stage2_invoice", "completed", {"elapsed": elapsed}, date_str=date_str)
                st.toast(f"✅ 가명세서 완료 ({format_elapsed(elapsed)})")
                if stdout:
                    with st.expander("실행 결과"):
                        st.code(stdout[-2000:], language="text")
                st.rerun()
            else:
                mark_stage("stage2_invoice", "failed", {"error": stderr[-500:]}, date_str=date_str)
                st.error("실패")
                st.code(stderr[-2000:], language="text")

    st.divider()

    # --- Generated invoices list ---
    st.subheader("생성된 명세서")
    invoice_dir = OUTPUTS_DIR / f"invoices_{date_compact}"
    if invoice_dir.exists() and invoice_dir.is_dir():
        files = sorted(invoice_dir.glob("*.xlsx"))
        if files:
            st.info(f"{len(files)}개 매장 명세서 생성됨")

            # Table view
            for f in files:
                size_kb = f.stat().st_size / 1024
                c_name, c_size, c_dl = st.columns([4, 1, 1])
                c_name.markdown(f"`{f.name}`")
                c_size.caption(f"{size_kb:.0f}KB")
                with open(f, "rb") as fh:
                    c_dl.download_button("↓", data=fh.read(), file_name=f.name, key=f"dl2_{f.stem}")

            # ZIP download
            st.divider()
            zip_path = OUTPUTS_DIR / f"거래명세서_{date_compact}.zip"
            if zip_path.exists():
                with open(zip_path, "rb") as zf:
                    st.download_button("전체 ZIP 다운로드", data=zf.read(), file_name=zip_path.name, mime="application/zip")
            else:
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for f in files:
                        zf.write(f, f.name)
                st.download_button(
                    "전체 ZIP 다운로드",
                    data=buf.getvalue(),
                    file_name=f"거래명세서_{date_compact}.zip",
                    mime="application/zip",
                )
        else:
            st.info("생성된 명세서가 없습니다.")
    else:
        st.info("명세서 폴더가 없습니다.")
