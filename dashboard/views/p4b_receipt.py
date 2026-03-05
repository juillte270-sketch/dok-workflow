"""Page 4b: 매입가 입력 - 영수증 일괄 처리 (Stage receipt).

중단 기능: Popen으로 실행 → PID를 state 파일에 저장 → 중단 버튼 클릭 시 kill.
Cloud 모드: st.file_uploader → Gemini API 파싱 → 단가시트 입력 → Drive 업로드.
"""
import streamlit as st
import os
import json
import signal
import sys
import tempfile
import time
import subprocess
import threading
from datetime import datetime, timedelta
from pathlib import Path
from dashboard.utils import (
    load_settings, run_script, run_script_streaming, mark_stage, get_stage_status,
    check_master_before_run, format_elapsed, get_timeout, append_log,
    PROJECT_ROOT, resolve_master_path, sync_master_back,
)
from dashboard.drive_service import is_cloud

# Drive 영수증 폴더 ID 캐시
_receipt_folder_id = None

# PID 파일 경로 (실행 중인 프로세스 추적)
_PID_FILE = os.path.join(PROJECT_ROOT, 'dashboard', 'state', 'receipt_pid.json')
_LOG_FILE = os.path.join(PROJECT_ROOT, 'dashboard', 'state', 'receipt_live_log.txt')


def _save_pid(pid):
    """실행 중인 프로세스 PID 저장."""
    with open(_PID_FILE, 'w') as f:
        json.dump({"pid": pid, "started": datetime.now().isoformat()}, f)


def _clear_pid():
    """PID 파일 삭제."""
    try:
        os.remove(_PID_FILE)
    except OSError:
        pass


def _load_pid():
    """저장된 PID 로드. 없으면 None."""
    try:
        with open(_PID_FILE, 'r') as f:
            data = json.load(f)
        return data.get("pid")
    except (OSError, json.JSONDecodeError):
        return None


def _is_process_alive(pid):
    """PID가 살아있는지 확인."""
    if pid is None:
        return False
    try:
        os.kill(pid, 0)
        return True
    except (OSError, ProcessLookupError):
        return False


def _kill_process(pid):
    """프로세스 종료."""
    try:
        os.kill(pid, signal.SIGTERM)
        return True
    except (OSError, ProcessLookupError):
        return False


def _find_receipt_dir():
    """카카오톡 다운로드 폴더 탐지."""
    user_profile = os.environ.get('USERPROFILE', '')
    candidates = [
        os.path.join(user_profile, 'Documents', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'Documents', 'KakaoTalk Downloads'),
        os.path.join(user_profile, '문서', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'OneDrive', '문서', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'OneDrive', 'Documents', '카카오톡 받은 파일'),
    ]
    for path in candidates:
        if os.path.isdir(path):
            return path
    return None


def _list_recent_images(folder, since_minutes=1440):
    """폴더 내 최근 이미지 파일 목록."""
    if not folder or not os.path.isdir(folder):
        return []

    cutoff = datetime.now() - timedelta(minutes=since_minutes)
    extensions = {'.jpg', '.jpeg', '.png', '.bmp'}
    files = []

    for fname in os.listdir(folder):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in extensions:
            continue
        fpath = os.path.join(folder, fname)
        mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
        if mtime >= cutoff:
            fsize = os.path.getsize(fpath)
            files.append({
                "name": fname,
                "path": fpath,
                "mtime": mtime,
                "size_kb": round(fsize / 1024, 1),
            })

    files.sort(key=lambda x: x["mtime"], reverse=True)
    return files


def _run_background(script_name, args, date_str, dry_run, num_images, timeout):
    """별도 스레드에서 프로세스를 실행, 결과를 로그 파일에 기록."""
    proc = run_script_streaming(script_name, args)
    _save_pid(proc.pid)

    # 실시간 로그를 파일에 쓰기
    with open(_LOG_FILE, 'w', encoding='utf-8') as log_f:
        log_f.write(f"[START] PID={proc.pid}\n")
        log_f.flush()

        start_t = time.time()
        while True:
            line = proc.stdout.readline()
            if line:
                log_f.write(line)
                log_f.flush()

            if proc.poll() is not None:
                remaining = proc.stdout.read()
                if remaining:
                    log_f.write(remaining)
                break

            if time.time() - start_t > timeout:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except Exception:
                    proc.kill()
                log_f.write(f"\n[TIMEOUT] {timeout}초 초과\n")
                break

        elapsed = time.time() - start_t
        returncode = proc.returncode if proc.returncode is not None else -1
        stderr = proc.stderr.read() if proc.stderr else ""

        # 결과 기록
        if returncode == 0 and not dry_run:
            mark_stage("stage_receipt", "completed", {"elapsed": elapsed}, date_str=date_str)
            append_log(f"영수증 일괄 처리 완료 ({num_images}개, {format_elapsed(elapsed)})")
        elif returncode != 0:
            # 중단(SIGTERM) 또는 에러
            error_msg = "사용자 중단" if returncode < 0 else (stderr or "")[-500:]
            mark_stage("stage_receipt", "failed", {"error": error_msg}, date_str=date_str)
            append_log(f"영수증 처리 실패/중단 (rc={returncode}, {format_elapsed(elapsed)})")

        log_f.write(f"\n[END] rc={returncode} elapsed={elapsed:.1f}s\n")

    _clear_pid()


def render(get_date_str, get_date_compact):
    date_str = get_date_str()

    st.header("8. 매입가 입력 (영수증)")

    settings = load_settings()

    # Cloud mode: file uploader 방식
    if is_cloud():
        st.caption("영수증 이미지를 업로드하여 Gemini API로 파싱 → 단가시트 매입가(Col I) 입력")
        _render_cloud_mode(date_str, settings)
        return

    st.caption("카카오톡 다운로드 폴더의 영수증 이미지를 일괄 파싱하여 단가시트 매입가(Col I)에 입력합니다.")
    master_ok = check_master_before_run(settings)

    # 스테이지 상태
    status = get_stage_status("stage_receipt", date_str)

    # 실행 중 프로세스 확인
    running_pid = _load_pid()
    is_running = running_pid is not None and _is_process_alive(running_pid)

    # PID 파일 있지만 프로세스 죽어있으면 정리
    if running_pid is not None and not is_running:
        _clear_pid()
        if status == "running":
            mark_stage("stage_receipt", "failed", {"error": "프로세스 비정상 종료"}, date_str=date_str)
        status = get_stage_status("stage_receipt", date_str)  # 재조회

    if is_running:
        st.warning("실행 중...")
        # 실시간 로그 표시
        _show_live_log()
        # 중단 버튼
        if st.button("중단", type="secondary", key="btn_cancel_receipt"):
            if _kill_process(running_pid):
                _clear_pid()
                mark_stage("stage_receipt", "failed", {"error": "사용자 중단"}, date_str=date_str)
                append_log("영수증 처리 사용자 중단")
                st.toast("처리가 중단되었습니다.")
                time.sleep(1)
                st.rerun()
            else:
                st.error("프로세스 종료 실패")
        # 자동 새로고침 (3초마다)
        time.sleep(3)
        st.rerun()
        return
    elif status == "completed":
        st.success("완료됨")
        # 최종 결과 로그 표시
        _show_final_log()
    elif status == "failed":
        st.error("이전 실행 실패")
        _show_final_log()

    st.divider()

    # === 폴더 설정 ===
    receipt_dir = _find_receipt_dir()
    custom_dir = st.text_input(
        "영수증 폴더 경로",
        value=receipt_dir or "",
        key="receipt_folder_path",
        help="카카오톡 다운로드 폴더 또는 영수증 이미지가 있는 폴더",
    )

    if custom_dir:
        receipt_dir = custom_dir

    if not receipt_dir or not os.path.isdir(receipt_dir):
        st.warning("유효한 영수증 폴더를 지정해주세요.")
        return

    st.divider()

    # === 최근 이미지 목록 ===
    col_list, col_action = st.columns([3, 1])

    with col_action:
        since_hours = st.selectbox(
            "기간 필터",
            [6, 12, 24, 48],
            index=2,
            format_func=lambda h: f"최근 {h}시간",
            key="receipt_since_hours",
        )

    since_minutes = since_hours * 60
    images = _list_recent_images(receipt_dir, since_minutes)

    with col_list:
        st.subheader(f"영수증 이미지 ({len(images)}개)")

    if not images:
        st.info(f"최근 {since_hours}시간 이내 이미지가 없습니다.")
        st.caption(f"폴더: {receipt_dir}")
        return

    # === 공급처 식별 ===
    supplier_map = st.session_state.get("receipt_supplier_map", {})

    col_table, col_identify = st.columns([3, 1])
    with col_identify:
        if st.button("공급처 확인", key="btn_identify_suppliers", help="Gemini로 각 이미지의 공급처를 식별합니다"):
            with st.spinner(f"공급처 식별 중... ({len(images)}개, 이미지당 3~5초)"):
                success, stdout, stderr, elapsed = run_script(
                    "parse_receipt.py",
                    ["--identify-dir", receipt_dir, "--since-minutes", str(since_minutes)],
                    timeout=max(300, len(images) * 10 + 60),
                )
            if success and stdout:
                try:
                    lines = stdout.strip().split('\n')
                    json_line = lines[-1]
                    id_results = json.loads(json_line)
                    supplier_map = {r["file"]: r["supplier"] for r in id_results}
                    st.session_state["receipt_supplier_map"] = supplier_map
                    st.toast(f"공급처 식별 완료 ({format_elapsed(elapsed)})")
                    st.rerun()
                except (json.JSONDecodeError, KeyError):
                    st.warning("식별 결과 파싱 실패")
                    if stdout:
                        st.code(stdout[-1000:], language="text")
            else:
                st.error("공급처 식별 실패")
                if stderr:
                    st.code(stderr[-500:], language="text")

    # 이미지 목록 테이블
    import pandas as pd
    table_data = []
    for img in images:
        row = {
            "파일명": img["name"],
            "공급처": supplier_map.get(img["name"], "-"),
            "시간": img["mtime"].strftime("%H:%M"),
            "크기": f"{img['size_kb']:.0f}KB",
        }
        table_data.append(row)
    df = pd.DataFrame(table_data)
    with col_table:
        st.dataframe(df, use_container_width=True, height=min(len(images) * 35 + 38, 300))

    # 이미지 미리보기 (최근 6개)
    preview_images = images[:6]
    if preview_images:
        with st.expander(f"이미지 미리보기 (최근 {len(preview_images)}개)", expanded=False):
            cols = st.columns(3)
            for i, img in enumerate(preview_images):
                with cols[i % 3]:
                    supplier_label = supplier_map.get(img["name"], "")
                    caption = f"{img['name']}"
                    if supplier_label and supplier_label != "-":
                        caption += f" [{supplier_label}]"
                    try:
                        st.image(img["path"], caption=caption, width=200)
                    except Exception:
                        st.caption(f"{img['name']} (미리보기 불가)")

    st.divider()

    # === 일괄 처리 실행 ===
    col_btn, col_opt = st.columns([2, 1])

    with col_opt:
        dry_run = st.checkbox("미리보기만 (dry-run)", value=False, key="receipt_batch_dry")

    with col_btn:
        btn_label = f"영수증 일괄 처리 ({len(images)}개)" if not dry_run else f"미리보기 실행 ({len(images)}개)"
        if st.button(
            btn_label,
            type="primary",
            key="btn_batch_receipt",
            disabled=not master_ok or not images,
        ):
            state_dir = os.path.join(PROJECT_ROOT, 'dashboard', 'state')
            args = [
                "--master", settings["master_file"],
                "--date", date_str,
                "--image-dir", receipt_dir,
                "--since-minutes", str(since_minutes),
                "--state-dir", state_dir,
            ]
            if dry_run:
                args.append("--dry-run")

            timeout = max(get_timeout("default"), len(images) * 20 + 60)
            mark_stage("stage_receipt", "running", date_str=date_str)
            append_log(f"START: fill_receipt_prices.py ({len(images)}개)")

            # 로그 파일 초기화
            with open(_LOG_FILE, 'w', encoding='utf-8') as f:
                f.write("")

            # 백그라운드 스레드에서 실행
            t = threading.Thread(
                target=_run_background,
                args=("fill_receipt_prices.py", args, date_str, dry_run, len(images), timeout),
                daemon=True,
            )
            t.start()
            time.sleep(1)  # 프로세스 시작 대기
            st.rerun()

    # === 개별 입력 (보조) ===
    st.divider()
    with st.expander("개별 영수증 입력 (URL/텍스트)"):
        _render_manual_input(settings, date_str, master_ok)


def _get_receipt_folder_id():
    """Drive 카톡 영수증 폴더 ID (마스터 파일 부모 → '카톡 영수증' 서브폴더)."""
    global _receipt_folder_id
    if _receipt_folder_id:
        return _receipt_folder_id
    from dashboard.drive_service import (
        get_drive_service, get_drive_file_ids, find_or_create_subfolder,
    )
    service = get_drive_service()
    file_ids = get_drive_file_ids()
    master_id = file_ids.get("master_file_id")
    meta = service.files().get(fileId=master_id, fields="parents").execute()
    parent_id = meta.get("parents", [None])[0]
    _receipt_folder_id = find_or_create_subfolder(parent_id, "카톡 영수증")
    return _receipt_folder_id


def _list_drive_receipts(since_hours=24):
    """Drive 카톡 영수증 폴더에서 최근 이미지 목록."""
    from dashboard.drive_service import list_folder
    folder_id = _get_receipt_folder_id()
    all_files = list_folder(folder_id)

    image_types = {"image/jpeg", "image/png", "image/bmp"}
    cutoff = datetime.now() - timedelta(hours=since_hours)

    result = []
    for f in all_files:
        if f.get("mimeType") not in image_types:
            continue
        mod_time = datetime.fromisoformat(f["modifiedTime"].replace("Z", "+00:00"))
        mod_local = mod_time.astimezone().replace(tzinfo=None)
        if mod_local < cutoff:
            continue
        result.append({
            "id": f["id"],
            "name": f["name"],
            "mtime": mod_local,
            "size_kb": round(int(f.get("size", 0)) / 1024, 1),
        })
    return result


def _process_drive_receipts(images, master_path, date_str, dry_run, from_drive):
    """Drive 이미지 다운로드 → parse_receipt → fill_receipt_prices → sync."""
    from dashboard.drive_service import download_file_by_id

    exec_dir = os.path.join(PROJECT_ROOT, 'execution')
    if exec_dir not in sys.path:
        sys.path.insert(0, exec_dir)

    try:
        import streamlit as _st
        gemini_key = _st.secrets.get("GEMINI_API_KEY", "")
        if gemini_key and not os.environ.get("GEMINI_API_KEY"):
            os.environ["GEMINI_API_KEY"] = gemini_key
    except Exception:
        pass

    from parse_receipt import parse_receipt
    from fill_receipt_prices import fill_receipt_prices

    mark_stage("stage_receipt", "running", date_str=date_str)
    all_receipts = []
    output_lines = []
    start_t = time.time()

    with st.spinner(f"Drive 영수증 파싱 중... ({len(images)}개)"):
        for img in images:
            tmp_path = download_file_by_id(img["id"])
            try:
                result = parse_receipt(image_path=tmp_path)
                if "error" in result:
                    output_lines.append(f"[WARN] {img['name']}: {result['error']}")
                else:
                    result["source_file"] = img["name"]
                    all_receipts.append(result)
                    n_items = len(result.get("items", []))
                    supplier = result.get("supplier", "?")
                    output_lines.append(f"[PARSE] {img['name']}: {supplier} ({n_items}개 품목)")
            except Exception as e:
                output_lines.append(f"[WARN] {img['name']}: 파싱 실패 - {e}")
            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

    if not all_receipts:
        elapsed = time.time() - start_t
        mark_stage("stage_receipt", "failed", {"error": "파싱된 영수증 없음"}, date_str=date_str)
        st.warning("파싱된 영수증이 없습니다.")
        if output_lines:
            st.code('\n'.join(output_lines), language="text")
        return

    with st.spinner("단가시트 매입가 입력 중..."):
        try:
            fill_result = fill_receipt_prices(
                master_file=master_path,
                target_date=date_str,
                receipts=all_receipts,
                dry_run=dry_run,
            )
            for detail in fill_result.get("details", []):
                output_lines.append(detail)
        except Exception as e:
            elapsed = time.time() - start_t
            mark_stage("stage_receipt", "failed", {"error": str(e)[-500:]}, date_str=date_str)
            st.error(f"매입가 입력 실패: {e}")
            if output_lines:
                st.code('\n'.join(output_lines), language="text")
            return

    elapsed = time.time() - start_t

    if from_drive and not dry_run:
        try:
            sync_master_back(master_path)
            output_lines.append("[SYNC] Drive 업로드 완료")
        except Exception as e:
            output_lines.append(f"[WARN] Drive 업로드 실패: {e}")

    if not dry_run:
        mark_stage("stage_receipt", "completed", {"elapsed": elapsed}, date_str=date_str)
        append_log(f"영수증 Drive 처리 완료 ({len(images)}개, {format_elapsed(elapsed)})")
        st.toast(f"영수증 처리 완료 ({format_elapsed(elapsed)})")
    else:
        mark_stage("stage_receipt", "pending", date_str=date_str)
        st.info("미리보기 모드 - 실제 입력되지 않았습니다.")

    stdout_text = '\n'.join(output_lines)
    _show_batch_result(stdout_text, "")


def _render_cloud_mode(date_str, settings):
    """Cloud 환경: Drive 영수증 + file uploader → Gemini 파싱 → 단가시트 입력."""
    master_path, from_drive = resolve_master_path(settings)

    # 스테이지 상태
    status = get_stage_status("stage_receipt", date_str)
    if status == "completed":
        st.success("완료됨")
    elif status == "failed":
        st.error("이전 실행 실패")

    st.divider()

    tab_drive, tab_upload = st.tabs(["Drive 영수증", "직접 업로드"])

    with tab_drive:
        _render_drive_tab(master_path, date_str, from_drive)

    with tab_upload:
        _render_upload_tab(master_path, date_str, from_drive)

    # 개별 입력 (URL/텍스트) — Cloud에서도 작동
    st.divider()
    with st.expander("개별 영수증 입력 (URL/텍스트)"):
        _render_manual_input_cloud(master_path, date_str, from_drive)


def _render_drive_tab(master_path, date_str, from_drive):
    """Drive 카톡 영수증 폴더에서 이미지 목록 → 일괄 처리."""
    import pandas as pd

    col_list, col_action = st.columns([3, 1])
    with col_action:
        since_hours = st.selectbox(
            "기간 필터",
            [6, 12, 24, 48],
            index=2,
            format_func=lambda h: f"최근 {h}시간",
            key="drive_receipt_since_hours",
        )

    try:
        images = _list_drive_receipts(since_hours)
    except Exception as e:
        st.error(f"Drive 폴더 접근 실패: {e}")
        return

    with col_list:
        st.subheader(f"Drive 영수증 ({len(images)}개)")

    if not images:
        st.info(f"최근 {since_hours}시간 이내 이미지가 없습니다.")
        return

    # 공급처 식별
    supplier_map = st.session_state.get("drive_receipt_supplier_map", {})

    col_table, col_identify = st.columns([3, 1])
    with col_identify:
        if st.button("공급처 확인", key="btn_drive_identify_suppliers",
                      help="Gemini로 각 이미지의 공급처를 식별합니다"):
            _identify_drive_suppliers(images)

    # 이미지 목록 테이블
    table_data = []
    for img in images:
        table_data.append({
            "파일명": img["name"],
            "공급처": supplier_map.get(img["name"], "-"),
            "시간": img["mtime"].strftime("%H:%M"),
            "크기": f"{img['size_kb']:.0f}KB",
        })
    df = pd.DataFrame(table_data)
    with col_table:
        st.dataframe(df, use_container_width=True, height=min(len(images) * 35 + 38, 300))

    # 옵션 + 실행 버튼
    col_btn, col_opt = st.columns([2, 1])
    with col_opt:
        dry_run = st.checkbox("미리보기만 (dry-run)", value=False, key="drive_receipt_dry")

    with col_btn:
        btn_label = f"영수증 일괄 처리 ({len(images)}개)" if not dry_run else f"미리보기 실행 ({len(images)}개)"
        if st.button(
            btn_label,
            type="primary",
            key="btn_drive_receipt",
            disabled=not images,
        ):
            _process_drive_receipts(images, master_path, date_str, dry_run, from_drive)


def _identify_drive_suppliers(images):
    """Drive 영수증 이미지를 다운로드하여 Gemini로 공급처 식별."""
    from dashboard.drive_service import download_file_by_id

    exec_dir = os.path.join(PROJECT_ROOT, 'execution')
    if exec_dir not in sys.path:
        sys.path.insert(0, exec_dir)

    try:
        import streamlit as _st
        gemini_key = _st.secrets.get("GEMINI_API_KEY", "")
        if gemini_key and not os.environ.get("GEMINI_API_KEY"):
            os.environ["GEMINI_API_KEY"] = gemini_key
    except Exception:
        pass

    from parse_receipt import identify_supplier_from_image

    supplier_map = {}
    progress = st.progress(0, text="공급처 식별 중...")

    for i, img in enumerate(images):
        progress.progress((i + 1) / len(images), text=f"공급처 식별 중... ({i+1}/{len(images)})")
        tmp_path = download_file_by_id(img["id"])
        try:
            supplier = identify_supplier_from_image(tmp_path)
            supplier_map[img["name"]] = supplier
        except Exception as e:
            supplier_map[img["name"]] = f"오류({e})"
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    progress.empty()
    st.session_state["drive_receipt_supplier_map"] = supplier_map
    st.toast(f"공급처 식별 완료 ({len(images)}개)")
    st.rerun()


def _render_upload_tab(master_path, date_str, from_drive):
    """기존 파일 업로드 방식."""
    uploaded_files = st.file_uploader(
        "영수증 이미지 업로드",
        type=["jpg", "jpeg", "png", "bmp"],
        accept_multiple_files=True,
        key="cloud_receipt_upload",
    )

    if uploaded_files:
        st.caption(f"{len(uploaded_files)}개 이미지 선택됨")

        with st.expander(f"이미지 미리보기 ({len(uploaded_files)}개)", expanded=False):
            cols = st.columns(3)
            for i, uf in enumerate(uploaded_files):
                with cols[i % 3]:
                    st.image(uf, caption=uf.name, width=200)

    st.divider()

    col_btn, col_opt = st.columns([2, 1])
    with col_opt:
        dry_run = st.checkbox("미리보기만 (dry-run)", value=False, key="cloud_receipt_dry")

    with col_btn:
        num = len(uploaded_files) if uploaded_files else 0
        btn_label = f"영수증 일괄 처리 ({num}개)" if not dry_run else f"미리보기 실행 ({num}개)"
        if st.button(
            btn_label,
            type="primary",
            key="btn_cloud_receipt",
            disabled=not uploaded_files,
        ):
            _process_cloud_receipts(uploaded_files, master_path, date_str, dry_run, from_drive)


def _process_cloud_receipts(uploaded_files, master_path, date_str, dry_run, from_drive):
    """Cloud: 업로드된 이미지 → temp 저장 → parse_receipt → fill_receipt_prices → sync."""
    # Ensure execution/ is in path
    exec_dir = os.path.join(PROJECT_ROOT, 'execution')
    if exec_dir not in sys.path:
        sys.path.insert(0, exec_dir)

    # Set GEMINI_API_KEY from Streamlit secrets if not in env
    try:
        import streamlit as _st
        gemini_key = _st.secrets.get("GEMINI_API_KEY", "")
        if gemini_key and not os.environ.get("GEMINI_API_KEY"):
            os.environ["GEMINI_API_KEY"] = gemini_key
    except Exception:
        pass

    from parse_receipt import parse_receipt, map_to_danga_items
    from fill_receipt_prices import fill_receipt_prices

    mark_stage("stage_receipt", "running", date_str=date_str)
    all_receipts = []
    output_lines = []
    start_t = time.time()

    with st.spinner(f"영수증 파싱 중... ({len(uploaded_files)}개)"):
        for uf in uploaded_files:
            # temp 파일 저장
            suffix = os.path.splitext(uf.name)[1] or ".jpg"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uf.getbuffer())
                tmp_path = tmp.name

            try:
                result = parse_receipt(image_path=tmp_path)
                if "error" in result:
                    output_lines.append(f"[WARN] {uf.name}: {result['error']}")
                else:
                    result["source_file"] = uf.name
                    all_receipts.append(result)
                    n_items = len(result.get("items", []))
                    supplier = result.get("supplier", "?")
                    output_lines.append(f"[PARSE] {uf.name}: {supplier} ({n_items}개 품목)")
            except Exception as e:
                output_lines.append(f"[WARN] {uf.name}: 파싱 실패 - {e}")
            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

    if not all_receipts:
        elapsed = time.time() - start_t
        mark_stage("stage_receipt", "failed", {"error": "파싱된 영수증 없음"}, date_str=date_str)
        st.warning("파싱된 영수증이 없습니다.")
        if output_lines:
            st.code('\n'.join(output_lines), language="text")
        return

    # 단가시트 입력
    with st.spinner("단가시트 매입가 입력 중..."):
        try:
            fill_result = fill_receipt_prices(
                master_file=master_path,
                target_date=date_str,
                receipts=all_receipts,
                dry_run=dry_run,
            )
            # Collect output from fill result
            for detail in fill_result.get("details", []):
                output_lines.append(detail)

        except Exception as e:
            elapsed = time.time() - start_t
            mark_stage("stage_receipt", "failed", {"error": str(e)[-500:]}, date_str=date_str)
            st.error(f"매입가 입력 실패: {e}")
            if output_lines:
                st.code('\n'.join(output_lines), language="text")
            return

    elapsed = time.time() - start_t

    # Drive 동기화
    if from_drive and not dry_run:
        try:
            sync_master_back(master_path)
            output_lines.append("[SYNC] Drive 업로드 완료")
        except Exception as e:
            output_lines.append(f"[WARN] Drive 업로드 실패: {e}")

    if not dry_run:
        mark_stage("stage_receipt", "completed", {"elapsed": elapsed}, date_str=date_str)
        append_log(f"영수증 Cloud 처리 완료 ({len(uploaded_files)}개, {format_elapsed(elapsed)})")
        st.toast(f"영수증 처리 완료 ({format_elapsed(elapsed)})")
    else:
        mark_stage("stage_receipt", "pending", date_str=date_str)
        st.info("미리보기 모드 - 실제 입력되지 않았습니다.")

    # 결과 표시
    stdout_text = '\n'.join(output_lines)
    _show_batch_result(stdout_text, "")


def _render_manual_input_cloud(master_path, date_str, from_drive):
    """Cloud용 개별 입력 (URL/텍스트) — subprocess에 master_path 전달."""
    input_type = st.radio(
        "입력 유형", ["URL", "텍스트"],
        horizontal=True, key="receipt_manual_type",
    )

    if input_type == "URL":
        url = st.text_input("영수증 URL", key="receipt_manual_url",
                            placeholder="https://www.itanet.co.kr/...")
        if st.button("URL 입력", key="btn_manual_url", disabled=not url):
            args = ["--master", master_path, "--date", date_str, "--url", url]
            with st.spinner("URL 파싱 + 입력 중..."):
                success, stdout, stderr, elapsed = run_script("fill_receipt_prices.py", args)
            if success:
                if from_drive:
                    try:
                        sync_master_back(master_path)
                    except Exception:
                        pass
                st.success(f"완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                st.error("실패")
                st.code(stderr[-1000:] if stderr else "Error", language="text")

    elif input_type == "텍스트":
        supplier = st.text_input("공급처명 (필수)", key="receipt_manual_supplier")
        text = st.text_area("영수증 텍스트", key="receipt_manual_text",
                            placeholder="감자 62000\n당근 35000")
        if st.button("텍스트 입력", key="btn_manual_text",
                      disabled=not text or not supplier):
            args = ["--master", master_path, "--date", date_str,
                    "--text", text, "--supplier", supplier]
            with st.spinner("텍스트 파싱 + 입력 중..."):
                success, stdout, stderr, elapsed = run_script("fill_receipt_prices.py", args)
            if success:
                if from_drive:
                    try:
                        sync_master_back(master_path)
                    except Exception:
                        pass
                st.success(f"완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                st.error("실패")
                st.code(stderr[-1000:] if stderr else "Error", language="text")


def _show_live_log():
    """실행 중 실시간 로그 표시."""
    try:
        with open(_LOG_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        if content.strip():
            lines = content.strip().split('\n')
            tail = lines[-25:]  # 최근 25줄
            st.code('\n'.join(tail), language="text")
            st.caption(f"총 {len(lines)} lines")
    except OSError:
        pass


def _show_final_log():
    """완료/실패 후 최종 로그 표시."""
    try:
        with open(_LOG_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        if content.strip():
            _show_batch_result(content, "")
    except OSError:
        pass


def _show_batch_result(stdout, stderr):
    """일괄 처리 결과를 파싱하여 보기 좋게 표시."""
    if not stdout:
        return

    lines = stdout.strip().split('\n')
    ok_lines = [l for l in lines if '[OK]' in l or '[DRY]' in l]
    skip_lines = [l for l in lines if '[SKIP]' in l]
    past_skip_lines = [l for l in skip_lines if '이전 날짜에 처리됨' in l]
    miss_lines = [l for l in lines if '[MISS]' in l or 'NOT FOUND' in l]
    warn_lines = [l for l in lines if '[WARN]' in l or '[!]' in l]
    crossref_lines = [l for l in lines if '[CROSSREF]' in l]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("입력", f"{len(ok_lines)}건")
    skip_label = f"{len(skip_lines)}건"
    if past_skip_lines:
        skip_label += f" (전거래일 {len(past_skip_lines)})"
    c2.metric("스킵", skip_label)
    c3.metric("미매칭", f"{len(miss_lines)}건")
    c4.metric("경고", f"{len(warn_lines)}건")

    if ok_lines:
        with st.expander(f"입력 완료 ({len(ok_lines)}건)", expanded=True):
            st.code('\n'.join(ok_lines), language="text")

    if crossref_lines:
        with st.expander(f"교차매칭 (발주리스트 대조)", expanded=False):
            st.code('\n'.join(crossref_lines), language="text")

    if miss_lines or warn_lines:
        with st.expander(f"미매칭/경고 ({len(miss_lines) + len(warn_lines)}건)", expanded=True):
            st.code('\n'.join(miss_lines + warn_lines), language="text")

    with st.expander("전체 로그"):
        st.code(stdout[-5000:], language="text")


def _render_manual_input(settings, date_str, master_ok):
    """URL/텍스트 개별 입력 (보조 기능)."""
    input_type = st.radio(
        "입력 유형", ["URL", "텍스트"],
        horizontal=True, key="receipt_manual_type",
    )

    if input_type == "URL":
        url = st.text_input("영수증 URL", key="receipt_manual_url",
                            placeholder="https://www.itanet.co.kr/...")
        if st.button("URL 입력", key="btn_manual_url", disabled=not master_ok or not url):
            args = ["--master", settings["master_file"], "--date", date_str, "--url", url]
            with st.spinner("URL 파싱 + 입력 중..."):
                success, stdout, stderr, elapsed = run_script("fill_receipt_prices.py", args)
            if success:
                st.success(f"완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                st.error("실패")
                st.code(stderr[-1000:] if stderr else "Error", language="text")

    elif input_type == "텍스트":
        supplier = st.text_input("공급처명 (필수)", key="receipt_manual_supplier")
        text = st.text_area("영수증 텍스트", key="receipt_manual_text",
                            placeholder="감자 62000\n당근 35000")
        if st.button("텍스트 입력", key="btn_manual_text",
                      disabled=not master_ok or not text or not supplier):
            args = ["--master", settings["master_file"], "--date", date_str,
                    "--text", text, "--supplier", supplier]
            with st.spinner("텍스트 파싱 + 입력 중..."):
                success, stdout, stderr, elapsed = run_script("fill_receipt_prices.py", args)
            if success:
                st.success(f"완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                st.error("실패")
                st.code(stderr[-1000:] if stderr else "Error", language="text")
