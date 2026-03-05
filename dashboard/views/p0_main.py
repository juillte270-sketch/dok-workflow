"""Page 0: Main Dashboard - Today's workflow overview with batch execution."""
import os
import streamlit as st
from dashboard.utils import (
    load_progress, save_progress, mark_stage, STAGES, STAGE_MAP,
    read_log, load_settings, check_master_before_run, run_script,
    run_script_safe, get_processed_orders_path, count_completed, append_log,
    check_prerequisites, render_stage_badge, format_elapsed,
    get_lock_info, backup_master_file, _state_path,
    resolve_master_path, sync_master_back,
)
from dashboard.drive_service import is_cloud

# All stages now use Playwright (headless) — Cloud-compatible
CLOUD_DISABLED_STAGES = set()


def _render_master_selector(settings):
    """Compact master file selector at top of main dashboard."""
    from pathlib import Path
    cloud_mode = is_cloud()

    if cloud_mode:
        from dashboard.drive_service import check_drive_connection
        ok, msg = check_drive_connection()
        icon = "🟢" if ok else "🔴"
        short_msg = msg.split("(")[0].strip() if ok else msg
        st.caption(f"{icon} Drive: {short_msg}")
        return

    master_dir = Path(r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료")
    xlsx_files = []
    if master_dir.exists():
        xlsx_files = sorted(
            [f for f in master_dir.glob("도크발주관리데이터*.xlsx")
             if not f.name.startswith("~$")],
            key=lambda f: f.stat().st_mtime, reverse=True,
        )

    current_master = settings.get("master_file", "")

    if xlsx_files:
        options = [str(f) for f in xlsx_files]
        labels = []
        for f in xlsx_files:
            size_mb = f.stat().st_size / (1024 * 1024)
            is_prod = f.name == "도크발주관리데이터.xlsx"
            tag = "운영" if is_prod else "테스트"
            labels.append(f"[{tag}] {f.name} ({size_mb:.1f}MB)")

        try:
            current_idx = options.index(current_master)
        except ValueError:
            current_idx = 0

        selected_idx = st.selectbox(
            "마스터 파일",
            range(len(options)),
            index=current_idx,
            format_func=lambda i: labels[i],
            key="main_master_select",
            label_visibility="collapsed",
        )

        if options[selected_idx] != current_master:
            settings["master_file"] = options[selected_idx]
            from dashboard.utils import save_settings
            save_settings(settings)
            st.rerun()
    else:
        p = Path(current_master)
        if p.exists():
            st.caption(f"📁 {p.name}")
        else:
            st.caption("⚠️ 마스터 파일 없음")


def _get_receipt_dir():
    """카카오톡 다운로드 폴더 경로 (p0_main용)."""
    user_profile = os.environ.get('USERPROFILE', '')
    for path in [
        os.path.join(user_profile, 'Documents', '카카오톡 받은 파일'),
        os.path.join(user_profile, '문서', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'OneDrive', '문서', '카카오톡 받은 파일'),
    ]:
        if os.path.isdir(path):
            return path
    return ""


def render(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    settings = load_settings()
    progress = load_progress(date_str)

    # --- Master file quick selector ---
    _render_master_selector(settings)

    # Top: title + metrics (pure HTML — no st.columns, always horizontal)
    completed = count_completed(date_str)
    total = len(STAGES)
    failed = sum(1 for v in progress.values() if isinstance(v, dict) and v.get("status") == "failed")
    waiting = total - completed - failed

    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">'
        f'<b style="font-size:1.1rem;flex:1">업무 현황</b>'
        f'<div style="display:flex;gap:6px">'
        f'<span style="background:#1E293B;border:1px solid #334155;border-radius:6px;padding:3px 10px;font-size:0.8rem">'
        f'<span style="color:#94A3B8">완료</span> <b style="color:#4ADE80">{completed}/{total}</b></span>'
        f'<span style="background:#1E293B;border:1px solid #334155;border-radius:6px;padding:3px 10px;font-size:0.8rem">'
        f'<span style="color:#94A3B8">실패</span> <b style="color:#F87171">{failed}</b></span>'
        f'<span style="background:#1E293B;border:1px solid #334155;border-radius:6px;padding:3px 10px;font-size:0.8rem">'
        f'<span style="color:#94A3B8">대기</span> <b style="color:#E2E8F0">{waiting}</b></span>'
        f'</div></div>',
        unsafe_allow_html=True,
    )

    # Batch buttons
    col_batch, col_kakao, col_reset = st.columns(3)
    with col_batch:
        if st.button("일괄 (1→5)", type="primary", key="btn_batch_morning"):
            _run_batch_morning(date_str, settings)
    with col_kakao:
        if st.button("카톡 전송", key="btn_kakao_summary"):
            _send_kakao_summary(date_str, progress)
    with col_reset:
        if st.button("초기화", key="btn_reset_progress"):
            _reset_progress(date_str)

    # Stage-by-stage control (compact 2-column layout)
    for s in STAGES:
        sid = s["id"]
        info = progress.get(sid, {})
        status = info.get("status", "pending")
        badge = render_stage_badge(status, info.get("timestamp", ""))
        sub = s.get("sub", "")
        sub_html = f' <span style="color:#64748B;font-size:0.75rem">{sub}</span>' if sub else ""

        c_info, c_btn = st.columns([3.5, 1])
        c_info.markdown(
            f'<div style="display:flex;align-items:center;gap:6px;min-height:32px">'
            f'<span class="step-num-sm">{s["icon"]}</span>'
            f'<span style="flex:1"><b style="font-size:0.88rem">{s["name"]}</b>{sub_html}</span>'
            f'{badge}'
            f'</div>',
            unsafe_allow_html=True,
        )

        btn_key = f"main_run_{sid}"
        cloud_mode = is_cloud()
        if s.get("dev"):
            c_btn.button("개발중", key=btn_key, disabled=True)
        elif cloud_mode and sid in CLOUD_DISABLED_STAGES:
            c_btn.button("미지원", key=btn_key, disabled=True)
        elif status == "running":
            c_btn.button("...", key=btn_key, disabled=True)
        else:
            if c_btn.button("실행", key=btn_key):
                _run_single_stage(sid, date_str, settings)

        # Show error inline if failed
        if status == "failed" and info.get("error"):
            with st.expander(f"오류: {s['name']}", expanded=False):
                st.code(info["error"][-1500:], language="text")
                if st.button("재시도", key=f"retry_{sid}"):
                    _run_single_stage(sid, date_str, settings)

    # Execution log
    with st.expander("실행 로그", expanded=False):
        log_text = read_log(date_str)
        if log_text:
            st.code(log_text[-5000:], language="text")
        else:
            st.info("실행 로그가 없습니다.")


def _get_stage_args(stage_id, date_str, settings, master_override=None):
    """Build correct command-line arguments for each stage."""
    master = master_override or settings["master_file"]
    processed = str(get_processed_orders_path(date_str))
    orders = str(get_processed_orders_path(date_str)).replace("processed_orders_", "orders_")

    # 다음날짜 재고현황: prev-date=당일, new-date=익일
    from datetime import datetime, timedelta
    next_date_str = (datetime.strptime(date_str, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")

    args_map = {
        "stage0_normalize": ["--input", orders],
        "stage1": ["--input", orders, "--date", date_str, "--master-file", master],
        "stage2_order": ["--date", date_str, "--master-file", master],
        "stage2_invoice": ["--date", date_str],
        "stage3_price": ["--input", processed, "--master", master, "--date", date_str],
        "stage3_sikbom": ["--master", master, "--date", date_str],
        "stage5_auction": ["--master", master, "--date", date_str],
        "stage4_helo": ["--master", master, "--date", date_str],
        "stage4_inventory": ["--prev-date", date_str, "--new-date", next_date_str, "--master", master],
        "stage4_second": ["--date", date_str, "--master", master, "--preview"],
        "stage_receipt": ["--master", master, "--date", date_str, "--image-dir", _get_receipt_dir(), "--since-minutes", "1440"],
        "stage7_fill": ["--date", date_str, "--master-file", master],
        "stage8_final": ["--date", date_str],
        "stage_report": ["--date", date_str],
    }
    return args_map.get(stage_id, [])


def _run_single_stage(stage_id, date_str, settings):
    """Execute one stage with lock, backup, and error handling."""
    stage = STAGE_MAP[stage_id]
    script = stage["script"]

    # Cloud mode: block unsupported stages
    if is_cloud() and stage_id in CLOUD_DISABLED_STAGES:
        st.warning(f"'{stage['name']}'은(는) Cloud 환경에서 지원되지 않습니다.")
        return

    # Check concurrent execution
    lock_info = get_lock_info()
    if lock_info:
        st.warning(f"다른 스크립트 실행 중: {lock_info.get('script', '?')}")
        return

    # Check master file for stages that need it
    needs_master = stage_id not in ("stage1",)
    if needs_master and not check_master_before_run(settings):
        return

    # Resolve master path (cloud: download from Drive)
    master_path, from_drive = resolve_master_path(settings)

    # Stages that modify master file get a backup
    modifying_stages = {"stage2_order", "stage3_price", "stage3_sikbom", "stage5_auction", "stage4_helo", "stage_receipt", "stage7_fill"}
    needs_backup = stage_id in modifying_stages

    mark_stage(stage_id, "running", date_str=date_str)
    args = _get_stage_args(stage_id, date_str, settings, master_override=master_path)

    timeout_key = {
        "stage3_sikbom": "sikbom",
        "stage5_auction": "auction",
        "stage4_helo": "helo",
        "stage8_final": "default",
    }.get(stage_id, "default")
    timeout = settings.get("timeouts", {}).get(timeout_key, 300)

    with st.spinner(f"{stage['name']} 실행 중..."):
        success, stdout, stderr, elapsed = run_script_safe(
            script, args, timeout=timeout, backup=needs_backup,
        )

    if success:
        # Cloud mode: upload modified master back to Drive
        if from_drive and stage_id in modifying_stages:
            try:
                sync_master_back(master_path)
            except Exception as e:
                st.warning(f"Drive 업로드 실패: {e}")

        mark_stage(stage_id, "completed", {"elapsed": elapsed}, date_str=date_str)
        st.toast(f"✅ {stage['name']} 완료 ({format_elapsed(elapsed)})")

        # stage4_helo 완료 후 → 다음날짜 재고현황 자동 생성
        if stage_id == "stage4_helo":
            inv_args = _get_stage_args("stage4_inventory", date_str, settings)
            with st.spinner("다음날짜 재고현황 생성 중..."):
                inv_ok, inv_out, inv_err, inv_elapsed = run_script_safe(
                    "manage_inventory_date.py", inv_args, timeout=300, backup=True,
                )
            if inv_ok:
                st.toast(f"✅ 재고현황 생성 완료 ({format_elapsed(inv_elapsed)})")
                if inv_out:
                    stdout = (stdout or "") + "\n\n=== 재고현황 생성 ===\n" + inv_out
            else:
                st.warning("⚠️ 재고현황 생성 실패 (HELO 입력은 완료)")
                if inv_err:
                    st.code(inv_err[-1000:], language="text")
    else:
        mark_stage(stage_id, "failed", {"error": stderr[-1000:], "elapsed": elapsed}, date_str=date_str)
        st.error(f"❌ {stage['name']} 실패")
        st.code(stderr[-2000:], language="text")

    if stdout:
        with st.expander("실행 결과", expanded=success):
            st.code(stdout[-3000:], language="text")

    st.rerun()


def _run_batch_morning(date_str, settings):
    """Run batch: stages 1 → 2a → 2b → 3 → 3-1(sikbom)."""
    batch_stages = ["stage1", "stage2_order", "stage2_invoice", "stage3_price", "stage3_sikbom"]

    # Cloud mode: skip unsupported stages
    cloud_mode = is_cloud()
    if cloud_mode:
        batch_stages = [s for s in batch_stages if s not in CLOUD_DISABLED_STAGES]

    append_log("=== BATCH START: 일괄 실행 ===")

    # Resolve master path once for the entire batch
    master_path, from_drive = resolve_master_path(settings)
    modifying_stages = {"stage2_order", "stage3_price", "stage3_sikbom", "stage5_auction", "stage4_helo", "stage_receipt", "stage7_fill"}

    progress_bar = st.progress(0, text="일괄 실행 준비 중...")
    total = len(batch_stages)

    for i, sid in enumerate(batch_stages):
        stage = STAGE_MAP[sid]
        progress_bar.progress((i) / total, text=f"{stage['name']} 실행 중...")

        # Skip if already completed
        current_status = load_progress(date_str).get(sid, {}).get("status")
        if current_status == "completed":
            st.toast(f"⏭️ {stage['name']} 이미 완료됨, 건너뜀")
            continue

        # Master check (skip for stage1)
        if sid != "stage1":
            if not check_master_before_run(settings):
                break

        mark_stage(sid, "running", date_str=date_str)
        args = _get_stage_args(sid, date_str, settings, master_override=master_path)
        timeout_key = {
            "stage3_sikbom": "sikbom",
            "stage5_auction": "auction",
            "stage4_helo": "helo",
        }.get(sid, "default")
        timeout = settings.get("timeouts", {}).get(timeout_key, 300)
        success, stdout, stderr, elapsed = run_script(stage["script"], args, timeout=timeout)

        if success:
            # Cloud: sync after each modifying stage
            if from_drive and sid in modifying_stages:
                try:
                    sync_master_back(master_path)
                    # Re-download for next stage to get fresh copy
                    master_path, from_drive = resolve_master_path(settings)
                except Exception as e:
                    st.warning(f"Drive 동기화 실패: {e}")
            mark_stage(sid, "completed", {"elapsed": elapsed}, date_str=date_str)
            st.toast(f"✅ {stage['name']} 완료")
        else:
            mark_stage(sid, "failed", {"error": stderr[-1000:], "elapsed": elapsed}, date_str=date_str)
            st.error(f"❌ {stage['name']} 실패 - 일괄 실행 중단")
            st.code(stderr[-2000:], language="text")
            break

    progress_bar.progress(1.0, text="일괄 실행 완료")
    append_log("=== BATCH END ===")
    st.rerun()


def _reset_progress(date_str):
    """Reset progress for the given date."""
    p = _state_path(date_str)
    if p.exists():
        p.unlink()
    append_log(f"진행상황 초기화: {date_str}")
    st.toast("진행상황이 초기화되었습니다.")
    st.rerun()


def _send_kakao_summary(date_str, progress):
    """Send daily summary via KakaoTalk."""
    results = {}
    for s in STAGES:
        info = progress.get(s["id"], {})
        results[s["id"]] = info.get("status") == "completed"

    import json
    results_json = json.dumps(results)

    with st.spinner("카톡 요약 전송 중..."):
        success, stdout, stderr, elapsed = run_script(
            "send_daily_summary.py",
            ["--date", date_str, "--results", results_json],
        )
    if success:
        st.toast("✅ 카톡 전송 완료!")
        append_log(f"카톡 요약 전송 완료 ({format_elapsed(elapsed)})")
    else:
        st.error("카톡 전송 실패")
        st.code(stderr[-1000:], language="text")
