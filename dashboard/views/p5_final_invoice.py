"""Page 5: Final Invoice Generation (Stage F) + Ledger update."""
import streamlit as st
import zipfile
import io
from pathlib import Path
from dashboard.utils import (
    load_settings, run_script, mark_stage, get_stage_status,
    OUTPUTS_DIR, check_master_before_run, check_prerequisites,
    format_elapsed, get_timeout, append_log,
)


def render(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    st.header("6. 최종 명세서")
    st.caption("가격이 채워진 송부용 거래명세서 + 거래원장")

    settings = load_settings()

    # Prerequisites check (info only — not blocking)
    st.divider()
    _check_and_show_prerequisites(date_str)

    st.divider()

    # --- Two columns ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("최종 명세서 생성")

        # 날짜 선택: 단일 또는 범위
        import datetime as _dt
        date_mode = st.radio(
            "날짜 선택", ["사이드바 날짜", "날짜 직접 지정", "날짜 범위"],
            horizontal=True, key="final_date_mode",
        )

        if date_mode == "사이드바 날짜":
            target_dates = [date_str]
        elif date_mode == "날짜 직접 지정":
            picked = st.date_input(
                "생성할 날짜",
                value=st.session_state.get("target_date", _dt.date.today()),
                key="final_single_date",
            )
            target_dates = [picked.strftime("%Y-%m-%d")]
        else:  # 날짜 범위
            c_s, c_e = st.columns(2)
            with c_s:
                d_start = st.date_input(
                    "시작일",
                    value=st.session_state.get("target_date", _dt.date.today()) - _dt.timedelta(days=1),
                    key="final_range_start",
                )
            with c_e:
                d_end = st.date_input(
                    "종료일",
                    value=st.session_state.get("target_date", _dt.date.today()),
                    key="final_range_end",
                )
            target_dates = []
            d = d_start
            while d <= d_end:
                target_dates.append(d.strftime("%Y-%m-%d"))
                d += _dt.timedelta(days=1)

        # 상태 표시
        for td in target_dates:
            s = get_stage_status("stage8_final", td)
            if s == "completed":
                st.success(f"{td} 완료됨")
            elif s == "failed":
                st.error(f"{td} 실패")

        btn_label = f"최종 명세서 생성 ({len(target_dates)}일)" if len(target_dates) > 1 else "최종 명세서 생성"
        if st.button(btn_label, type="primary", key="btn_final_inv"):
            if not check_master_before_run(settings):
                pass
            else:
                all_stdout = []
                any_fail = False
                for i, td in enumerate(target_dates):
                    with st.spinner(f"최종 명세서 생성 중... ({td}, {i+1}/{len(target_dates)})"):
                        success, stdout, stderr, elapsed = run_script(
                            "generate_final_invoices.py",
                            ["--date", td],
                        )
                    if success:
                        mark_stage("stage8_final", "completed", {"elapsed": elapsed}, date_str=td)
                        st.toast(f"✅ {td} 완료 ({format_elapsed(elapsed)})")
                        if stdout:
                            all_stdout.append(f"=== {td} ===\n{stdout}")
                    else:
                        mark_stage("stage8_final", "failed", {"error": stderr[-500:]}, date_str=td)
                        st.error(f"❌ {td} 실패")
                        if stderr:
                            st.code(stderr[-1500:], language="text")
                        any_fail = True

                if all_stdout:
                    with st.expander("실행 결과", expanded=True):
                        st.code("\n\n".join(all_stdout)[-5000:], language="text")
                st.rerun()

    with col2:
        st.subheader("거래원장 업데이트")
        st.caption("최종 명세서와 반드시 함께 생성해야 합니다.")

        import datetime
        target_date = st.session_state.get("target_date", datetime.date.today())
        first_of_month = target_date.replace(day=1)

        start_date = st.date_input("시작일", value=first_of_month, key="ledger_start")
        end_date = st.date_input("종료일", value=target_date, key="ledger_end")

        if st.button("거래원장 업데이트", key="btn_ledger"):
            if not check_master_before_run(settings):
                pass
            else:
                with st.spinner("거래원장 업데이트 중..."):
                    success, stdout, stderr, elapsed = run_script(
                        "batch_generate_ledgers_v3.py",
                        [
                            "--master", settings["master_file"],
                            "--start", start_date.strftime("%Y-%m-%d"),
                            "--end", end_date.strftime("%Y-%m-%d"),
                        ],
                        timeout=get_timeout("ledger"),
                    )
                if success:
                    append_log(f"거래원장 업데이트 완료 ({format_elapsed(elapsed)})")
                    st.toast(f"✅ 거래원장 완료 ({format_elapsed(elapsed)})")
                    if stdout:
                        with st.expander("실행 결과"):
                            st.code(stdout[-3000:], language="text")
                else:
                    st.error("실패")
                    st.code(stderr[-2000:], language="text")

    st.divider()

    # --- Generated files ---
    st.subheader("생성된 파일")
    # Show files for all target dates (if range selected), else sidebar date
    shown_compacts = set()
    for td in target_dates:
        dc = td.replace("-", "")
        if dc not in shown_compacts:
            shown_compacts.add(dc)
            _show_final_files(dc)


def _check_and_show_prerequisites(date_str):
    """Show prerequisite status (informational only, does NOT block execution).

    대시보드 진행기록 기준이므로 CLI 실행 등은 감지 못함.
    실제 실행 가능 여부는 마스터 파일 데이터 유무로 결정됨.
    """
    prereqs = [
        ("stage1", "발주 처리"),
        ("stage2_order", "발주시트"),
        ("stage2_invoice", "가명세서"),
        ("stage7_fill", "매입/판매가"),
    ]

    done = sum(1 for sid, _ in prereqs if get_stage_status(sid, date_str) == "completed")
    total = len(prereqs)

    if done == total:
        st.caption(f"선행 단계: {done}/{total} 완료")
    else:
        missing = [name for sid, name in prereqs if get_stage_status(sid, date_str) != "completed"]
        st.caption(f"선행 단계: {done}/{total} (대시보드 기록 기준 — CLI 실행 시 미반영)")

    return done == total


def _show_final_files(date_compact):
    """Display generated final invoice files with download buttons."""
    final_dir = OUTPUTS_DIR / f"final_invoices_{date_compact}"

    if not final_dir.exists() or not final_dir.is_dir():
        st.caption(f"{date_compact[:4]}-{date_compact[4:6]}-{date_compact[6:]}: 폴더 없음")
        return

    files = sorted(final_dir.glob("*.xlsx"))
    if not files:
        st.caption(f"{date_compact}: 생성된 파일 없음")
        return

    st.info(f"**{date_compact[:4]}-{date_compact[4:6]}-{date_compact[6:]}** — {len(files)}개 최종 명세서")

    for f in files:
        size_kb = f.stat().st_size / 1024
        c_name, c_size, c_dl = st.columns([4, 1, 1])
        c_name.markdown(f"`{f.name}`")
        c_size.caption(f"{size_kb:.0f}KB")
        with open(f, "rb") as fh:
            c_dl.download_button(
                "↓",
                data=fh.read(),
                file_name=f.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl5_{date_compact}_{f.stem}",
            )

    # ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.name)
    st.download_button(
        f"ZIP 다운로드 ({date_compact})",
        data=buf.getvalue(),
        file_name=f"최종명세서_{date_compact}.zip",
        mime="application/zip",
        key=f"zip5_{date_compact}",
    )
