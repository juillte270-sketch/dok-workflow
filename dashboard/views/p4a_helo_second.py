"""Page 4a: HELO Scraping (Stage 4) & Second Order Message (Stage 4-1)."""
import streamlit as st
from datetime import datetime, timedelta
from dashboard.utils import (
    load_settings, run_script, mark_stage, get_stage_status,
    check_master_before_run, format_elapsed, get_timeout, append_log,
    resolve_master_path, sync_master_back,
)
from dashboard.drive_service import is_cloud


def render(get_date_str, get_date_compact):
    date_str = get_date_str()

    st.header("7. 동원발주 스크래핑")

    settings = load_settings()
    master_ok = check_master_before_run(settings)
    master_path, from_drive = resolve_master_path(settings)

    st.divider()

    col1, col2 = st.columns(2)

    # === Stage 4: HELO scraping + 재고현황 생성 ===
    with col1:
        st.subheader("HELO 2차출고 스크래핑")
        st.caption("동원 HELO 스크래핑 → 2차출고 입력 → 다음날짜 재고현황 자동 생성")

        status_helo = get_stage_status("stage4_helo", date_str)
        if status_helo == "completed":
            st.success("완료됨")
        elif status_helo == "failed":
            st.error("실패")

        st.info("오후 10시 (동원 마감 후) 실행 권장")

        if st.button(
            "HELO 스크래핑 실행",
            type="primary",
            key="btn_helo",
            disabled=not master_ok,
        ):
            # Step 1: HELO scraping
            with st.spinner("HELO 스크래핑 중... (브라우저 자동화)"):
                success, stdout, stderr, elapsed = run_script(
                    "scrape_helo.py",
                    ["--master", master_path, "--date", date_str],
                    timeout=get_timeout("default"),
                )
            if success:
                mark_stage("stage4_helo", "completed", {"elapsed": elapsed}, date_str=date_str)
                st.toast(f"✅ HELO 스크래핑 완료 ({format_elapsed(elapsed)})")

                # Step 2: 다음날짜 재고현황 자동 생성
                next_date = (datetime.strptime(date_str, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
                with st.spinner("다음날짜 재고현황 생성 중..."):
                    inv_ok, inv_out, inv_err, inv_elapsed = run_script(
                        "manage_inventory_date.py",
                        ["--prev-date", date_str, "--new-date", next_date, "--master", master_path],
                        timeout=300,
                    )
                if inv_ok:
                    st.toast(f"✅ 재고현황 생성 완료 ({format_elapsed(inv_elapsed)})")
                    combined = (stdout or "") + "\n\n=== 재고현황 생성 ===\n" + (inv_out or "")
                else:
                    st.warning("⚠️ 재고현황 생성 실패 (HELO 입력은 완료)")
                    combined = (stdout or "") + "\n\n=== 재고현황 생성 실패 ===\n" + (inv_err or "")

                if combined.strip():
                    with st.expander("실행 결과", expanded=True):
                        st.code(combined[-3000:], language="text")
                st.rerun()
            else:
                mark_stage("stage4_helo", "failed", {"error": stderr[-500:]}, date_str=date_str)
                st.error("스크래핑 실패")
                st.code(stderr[-2000:], language="text")
                if st.button("재시도", key="btn_helo_retry"):
                    st.rerun()

    # === Stage 4-1: Second order message ===
    with col2:
        st.subheader("2차 발주 메시지 전송")
        st.caption("HELO 스크래핑 결과를 기반으로 2차 발주 메시지를 카카오톡으로 전송합니다.")

        status_second = get_stage_status("stage4_second", date_str)
        if status_second == "completed":
            st.success("완료됨")
        elif status_second == "failed":
            st.error("실패")

        # HELO should be done first
        helo_done = get_stage_status("stage4_helo", date_str) == "completed"
        if not helo_done:
            st.caption("⚠️ HELO 스크래핑을 먼저 실행하세요")

        # Preview option
        preview_mode = st.checkbox("미리보기만 (전송 안 함)", value=True, key="second_preview")

        if st.button(
            "2차 발주 메시지",
            type="primary",
            key="btn_second_order",
            disabled=not helo_done,
        ):
            args = ["--date", date_str, "--master", master_path]
            if preview_mode:
                args.append("--preview")

            with st.spinner("2차 발주 메시지 생성 중..."):
                success, stdout, stderr, elapsed = run_script(
                    "second_order.py",
                    args,
                    timeout=get_timeout("default"),
                )
            if success:
                if not preview_mode:
                    mark_stage("stage4_second", "completed", {"elapsed": elapsed}, date_str=date_str)
                    st.toast(f"✅ 2차 발주 전송 완료 ({format_elapsed(elapsed)})")
                else:
                    st.info("미리보기 모드 - 실제 전송되지 않았습니다.")
                if stdout:
                    with st.expander("메시지 내용", expanded=True):
                        st.code(stdout[-3000:], language="text")
            else:
                if not preview_mode:
                    mark_stage("stage4_second", "failed", {"error": stderr[-500:]}, date_str=date_str)
                st.error("실패")
                st.code(stderr[-2000:], language="text")
