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


DRIVE_RECEIPT_DIR = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\카톡 영수증"


def _get_receipt_dir():
    """영수증 이미지 폴더 경로. Drive 카톡 영수증 폴더 우선, 없으면 카카오톡 폴더."""
    if os.path.isdir(DRIVE_RECEIPT_DIR):
        return DRIVE_RECEIPT_DIR
    user_profile = os.environ.get('USERPROFILE', '')
    for path in [
        os.path.join(user_profile, 'OneDrive', '문서', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'Documents', '카카오톡 받은 파일'),
        os.path.join(user_profile, '문서', '카카오톡 받은 파일'),
    ]:
        if os.path.isdir(path):
            return path
    return ""


def _sync_receipts_to_drive():
    """카카오톡 받은 파일 → Drive 카톡 영수증 폴더 동기화 (1회)."""
    try:
        from dashboard.utils import run_script
        success, stdout, stderr, elapsed = run_script(
            "sync_receipts.py", ["--once", "--since", "1440"], timeout=30,
        )
        if success and stdout:
            # 복사 건수 추출
            for line in stdout.strip().split("\n"):
                if "동기화 완료" in line:
                    return line.strip()
        return None
    except Exception:
        return None


def render(get_date_str, get_date_compact):
    try:
        _render_inner(get_date_str, get_date_compact)
    except Exception as e:
        import traceback
        st.error(f"대시보드 렌더링 오류: {e}")
        st.code(traceback.format_exc(), language="text")


def _render_inner(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    settings = load_settings()
    progress = load_progress(date_str)

    # Auto-recover stale "running" stages (>15 min)
    from dashboard.utils import _recover_stale_running
    _recover_stale_running(progress, date_str)

    # === Handle stage execution triggered by HTML link click ===
    params = st.query_params
    run_target = params.get("run")
    if run_target and run_target in STAGE_MAP:
        st.query_params.clear()
        _run_single_stage(run_target, date_str, settings)

    # === Header: metrics + Drive status (ONE HTML block) ===
    completed = count_completed(date_str)
    total = len(STAGES)
    failed = sum(1 for v in progress.values() if isinstance(v, dict) and v.get("status") == "failed")
    waiting = total - completed - failed

    # Drive/Master status (inline)
    drive_html = ""
    try:
        cloud_mode = is_cloud()
        if cloud_mode:
            from dashboard.drive_service import check_drive_connection
            ok, msg = check_drive_connection()
            short = msg.split("(")[0].strip() if ok else msg
            color = "#4ADE80" if ok else "#F87171"
            drive_html = f'<div style="font-size:0.72rem;color:#64748B;margin-bottom:2px">{short}</div>'
        else:
            from pathlib import Path
            master = settings.get("master_file", "")
            mname = Path(master).name if master else ""
            if mname:
                drive_html = f'<div style="font-size:0.72rem;color:#64748B;margin-bottom:2px">{mname}</div>'
    except Exception:
        pass

    # === 배송 현황 요약 ===
    _render_delivery_summary(date_compact)

    # === Stage list — ONE HTML block ===
    rows_html = ""
    for s in STAGES:
        sid = s["id"]
        info = progress.get(sid, {})
        status = info.get("status", "pending")
        sub = s.get("sub", "")
        sub_html = f' <span style="color:#64748B;font-size:0.7rem">{sub}</span>' if sub else ""
        ts = info.get("timestamp", "")
        ts_short = ts[11:16] if len(ts) > 16 else ""

        # Status dot + right-side action
        if status == "completed":
            dot = "background:#4ADE80"
            right_html = (
                f'<span style="color:#4ADE80;font-size:0.8rem">완료 {ts_short}</span>'
                f'&nbsp;<a href="?run={sid}" target="_self" '
                f'style="color:#64748B;font-size:0.7rem;text-decoration:none;'
                f'padding:1px 5px;border:1px solid #475569;border-radius:3px">재</a>'
            )
            bg = "background:rgba(74,222,128,0.04);"
        elif status == "failed":
            dot = "background:#F87171"
            right_html = (
                f'<a href="?run={sid}" target="_self" '
                f'style="color:#F87171;font-size:0.8rem;text-decoration:none;'
                f'padding:2px 8px;border:1px solid #F87171;border-radius:4px">재시도</a>'
            )
            bg = "background:rgba(248,113,113,0.04);"
        elif status == "running":
            dot = "background:#FBBF24;animation:pulse 1.5s ease-in-out infinite"
            right_html = '<span style="color:#FBBF24;font-size:0.8rem">실행중</span>'
            bg = "background:rgba(251,191,36,0.04);"
        elif s.get("dev"):
            dot = "background:#475569"
            right_html = '<span style="color:#64748B;font-size:0.75rem">개발중</span>'
            bg = ""
        else:
            dot = "background:#475569"
            right_html = (
                f'<a href="?run={sid}" target="_self" '
                f'style="color:#818CF8;font-size:0.8rem;text-decoration:none;'
                f'padding:2px 10px;border:1px solid #818CF8;border-radius:4px">실행</a>'
            )
            bg = ""

        rows_html += (
            f'<div style="display:flex;align-items:center;gap:8px;padding:8px 10px;'
            f'border-bottom:1px solid #1E293B;{bg}">'
            f'<span style="display:inline-flex;align-items:center;justify-content:center;'
            f'width:22px;height:22px;border-radius:50%;background:#334155;'
            f'color:#CBD5E1;font-size:0.65rem;font-weight:600;flex-shrink:0">{s["icon"]}</span>'
            f'<span style="flex:1;font-size:0.88rem;line-height:1.3"><b>{s["name"]}</b>{sub_html}</span>'
            f'<span style="flex-shrink:0">{right_html}</span>'
            f'</div>'
        )

    st.markdown(
        f'{drive_html}'
        f'<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;flex-wrap:wrap">'
        f'<b style="font-size:1rem;flex:1">업무 현황</b>'
        f'<span style="font-size:0.75rem">'
        f'<b style="color:#4ADE80">{completed}</b>/<span style="color:#64748B">{total}</span>'
        f'&nbsp; 실패 <b style="color:#F87171">{failed}</b>'
        f'&nbsp; 대기 <b style="color:#94A3B8">{waiting}</b>'
        f'</span></div>'
        f'<div style="border:1px solid #334155;border-radius:10px;overflow:hidden;margin:4px 0">'
        f'{rows_html}</div>',
        unsafe_allow_html=True,
    )

    # === Batch + utility (expander) ===
    with st.expander("일괄 실행 & 기타", expanded=False):
        st.caption(
            "**일괄 실행**: 오전 5단계를 순차 실행합니다.\n"
            "1️⃣ 발주처리 → 2️⃣ 발주시트 → 3️⃣ 가명세서 → 4️⃣ 단가표 → 5️⃣ 식봄가\n\n"
            "HELO·경매가·영수증 등은 별도 시간대에 개별 실행하세요."
        )
        if st.button("일괄 실행 (1-5)", type="primary", key="btn_batch_morning"):
            _run_batch_morning(date_str, settings)
        if st.button("카톡 요약 전송", key="btn_kakao_summary"):
            _send_kakao_summary(date_str, progress)
        if st.button("진행상황 초기화", key="btn_reset_progress"):
            _reset_progress(date_str)

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
        "stage8_final": ["--date", date_str, "--master", master],
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

    # 영수증 스테이지: 실행 전 카카오톡 → Drive 동기화
    if stage_id == "stage_receipt" and not is_cloud():
        with st.spinner("카카오톡 영수증 동기화 중..."):
            sync_msg = _sync_receipts_to_drive()
        if sync_msg:
            st.toast(f"📋 {sync_msg}")

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

    try:
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

    except Exception as e:
        # Ensure "running" never gets stuck — always mark failed on crash
        mark_stage(stage_id, "failed", {"error": str(e)[:500]}, date_str=date_str)
        st.error(f"❌ {stage['name']} 예외 발생: {e}")

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


def _render_delivery_summary(date_compact):
    """메인 대시보드에 배송 현황 요약 카드 표시."""
    try:
        from dashboard import firebase_service as fb
        deliveries = fb.get_deliveries(date_compact)
    except Exception:
        deliveries = None

    if not deliveries:
        st.markdown(
            '<div style="border:1px solid #334155;border-radius:10px;padding:12px 14px;margin:8px 0;cursor:pointer">'
            '<div style="display:flex;align-items:center;justify-content:space-between">'
            '<b style="font-size:0.92rem">🚛 배송현황</b>'
            '<span style="color:#64748B;font-size:0.8rem">데이터 없음</span>'
            '</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("배송 관리 →", key="go_delivery_empty", type="tertiary"):
            st.session_state["nav_page"] = "12. 배송 관리"
            st.rerun()
        return

    stats = fb.compute_delivery_stats(deliveries)
    total = stats["total"]
    done = stats["delivered"]
    transit = stats["in_transit"]
    arrived = stats["arrived"]
    pending = stats["pending"]
    issue = stats["issue"]
    pct = int(done / total * 100) if total else 0

    bar_color = "#4ADE80" if pct == 100 else "#818CF8"
    issue_html = f'<span style="color:#F87171;margin-left:6px">문제 {issue}</span>' if issue else ""

    st.markdown(
        f'<div style="border:1px solid #334155;border-radius:10px;padding:12px 14px;margin:8px 0">'
        f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">'
        f'<b style="font-size:0.92rem">🚛 배송현황</b>'
        f'<span style="font-size:0.78rem;color:#94A3B8">'
        f'완료 <b style="color:#4ADE80">{done}</b> · '
        f'이동 <b style="color:#60A5FA">{transit}</b> · '
        f'입고 <b style="color:#FBBF24">{arrived}</b> · '
        f'대기 <b style="color:#94A3B8">{pending}</b>'
        f'{issue_html}'
        f'</span></div>'
        f'<div style="background:#1E293B;border-radius:4px;height:6px;overflow:hidden">'
        f'<div style="background:{bar_color};width:{pct}%;height:100%;border-radius:4px"></div>'
        f'</div>'
        f'<div style="text-align:right;font-size:0.7rem;color:#64748B;margin-top:2px">'
        f'{done}/{total} ({pct}%)</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    if st.button("배송 관리 →", key="go_delivery", type="tertiary"):
        st.session_state["nav_page"] = "12. 배송 관리"
        st.rerun()
