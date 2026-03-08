"""Page 10: 일일보고서 생성 (Stage report).

로컬: subprocess로 generate_daily_report.py 실행
Cloud: Drive에서 이전 보고서 다운로드 → 날짜 교체 → Drive 업로드
"""
import os
import sys
import tempfile
import streamlit as st
from pathlib import Path
from datetime import datetime

from dashboard.utils import (
    load_settings, run_script, mark_stage, get_stage_status,
    format_elapsed, append_log, PROJECT_ROOT,
)
from dashboard.drive_service import is_cloud


def render(get_date_str, get_date_compact):
    date_str = get_date_str()

    st.header("13. 일일보고서")

    status = get_stage_status("stage_report", date_str)
    if status == "completed":
        st.success("완료됨")
    elif status == "failed":
        st.error("이전 실행 실패")

    st.divider()

    if is_cloud():
        _render_cloud(date_str)
    else:
        _render_local(date_str)


def _render_local(date_str):
    """로컬 모드: subprocess로 실행."""
    st.caption("전날 HWP 보고서를 복사 → 날짜만 교체하여 오늘 보고서 생성")

    col1, col2 = st.columns([2, 1])
    with col2:
        dry_run = st.checkbox("미리보기만", value=False, key="report_dry_run")

    with col1:
        btn_label = "보고서 생성" if not dry_run else "미리보기"
        if st.button(btn_label, type="primary", key="btn_report_local"):
            args = ["--date", date_str]
            if dry_run:
                args.append("--dry-run")

            with st.spinner("보고서 생성 중..."):
                success, stdout, stderr, elapsed = run_script(
                    "generate_daily_report.py", args, timeout=60,
                )

            if success:
                if not dry_run:
                    mark_stage("stage_report", "completed", {"elapsed": elapsed}, date_str=date_str)
                    st.toast(f"보고서 생성 완료 ({format_elapsed(elapsed)})")
                if stdout:
                    st.code(stdout[-2000:], language="text")
            else:
                mark_stage("stage_report", "failed", {"error": (stderr or "")[-500:]}, date_str=date_str)
                st.error("보고서 생성 실패")
                st.code(stderr[-2000:] if stderr else "Error", language="text")


def _render_cloud(date_str):
    """Cloud 모드: Drive에서 이전 보고서 → 날짜 교체 → Drive 업로드."""
    st.caption("Drive에서 이전 보고서를 가져와 날짜를 교체한 후 다시 Drive에 업로드합니다.")

    from dashboard.drive_service import (
        get_drive_file_ids, find_or_create_subfolder,
        search_file_in_folder, download_file_by_id,
        upload_file_to_folder,
    )

    # Check secrets
    file_ids = get_drive_file_ids()
    report_folder_id = file_ids.get("report_folder_id")
    if not report_folder_id:
        st.error("secrets.toml에 `[drive] report_folder_id`가 설정되지 않았습니다.")
        return

    target_date = datetime.strptime(date_str, "%Y-%m-%d")

    # Ensure execution module is importable
    exec_dir = os.path.join(PROJECT_ROOT, 'execution')
    if exec_dir not in sys.path:
        sys.path.insert(0, exec_dir)
    from generate_daily_report import (
        month_folder_name, report_filename, get_day_kr,
    )

    target_fname = report_filename(target_date)
    month_name = month_folder_name(target_date)

    st.info(f"생성 대상: **{target_fname}** (폴더: {month_name})")

    col1, col2 = st.columns([2, 1])
    with col2:
        dry_run = st.checkbox("미리보기만", value=False, key="report_cloud_dry")

    with col1:
        btn_label = "보고서 생성" if not dry_run else "미리보기"
        if st.button(btn_label, type="primary", key="btn_report_cloud"):
            _run_cloud_report(
                target_date, date_str, report_folder_id,
                month_name, target_fname, dry_run,
            )


def _run_cloud_report(target_date, date_str, report_folder_id,
                      month_name, target_fname, dry_run):
    """Cloud 보고서 생성 실행.

    Service Account는 자체 저장 용량이 없으므로 files().create() 불가.
    대신 files().copy()로 Drive 내에서 복사 → 다운로드 → 날짜 교체 → update로 덮어쓰기.
    copy()는 원본 소유자의 용량을 사용하므로 quota 문제 없음.
    """
    import time
    from dashboard.drive_service import (
        get_drive_service,
        find_or_create_subfolder, search_file_in_folder,
        download_file_by_id,
    )

    exec_dir = os.path.join(PROJECT_ROOT, 'execution')
    if exec_dir not in sys.path:
        sys.path.insert(0, exec_dir)
    from generate_daily_report import (
        generate_report, month_folder_name, report_filename, get_day_kr,
    )
    from datetime import timedelta

    mark_stage("stage_report", "running", date_str=date_str)
    start_t = time.time()

    try:
        with st.spinner("Drive에서 이전 보고서 검색 중..."):
            # 1. Find or create month subfolder
            month_folder_id = find_or_create_subfolder(report_folder_id, month_name)

            # 2. Check if target already exists
            existing = search_file_in_folder(month_folder_id, target_fname.replace(".hwp", ""))
            if existing:
                for f in existing:
                    if f["name"] == target_fname:
                        elapsed = time.time() - start_t
                        mark_stage("stage_report", "completed", {"elapsed": elapsed}, date_str=date_str)
                        st.warning(f"이미 존재: {target_fname}")
                        return

            # 3. Search for previous report (try multiple lookback days)
            prev_file = None
            prev_date = None
            for i in range(1, 15):
                check_date = target_date - timedelta(days=i)
                check_month = month_folder_name(check_date)

                # Search in same or different month folder
                if check_month == month_name:
                    search_folder_id = month_folder_id
                else:
                    # Try to find prev month folder
                    prev_month_files = search_file_in_folder(report_folder_id, check_month)
                    if not prev_month_files:
                        continue
                    search_folder_id = prev_month_files[0]["id"]

                # Search for file by date pattern
                date_pattern = f"{check_date.month}.{check_date.day}"
                candidates = search_file_in_folder(search_folder_id, date_pattern)
                hwp_candidates = [f for f in candidates if f["name"].endswith(".hwp")]
                if hwp_candidates:
                    prev_file = hwp_candidates[0]
                    prev_date = check_date
                    break

            if prev_file is None:
                elapsed = time.time() - start_t
                mark_stage("stage_report", "failed", {"error": "이전 보고서 없음"}, date_str=date_str)
                st.error("Drive에서 이전 보고서를 찾을 수 없습니다 (최근 14일).")
                return

            st.caption(f"이전 보고서: {prev_file['name']}")

        if dry_run:
            elapsed = time.time() - start_t
            mark_stage("stage_report", "pending", date_str=date_str)
            st.info(f"[DRY-RUN] 생성 예정: {target_fname}")
            st.caption(f"이전 보고서: {prev_file['name']}")
            st.caption(f"날짜 변경: {prev_date.month}/{prev_date.day}({get_day_kr(prev_date)}) → "
                       f"{target_date.month}/{target_date.day}({get_day_kr(target_date)})")
            return

        with st.spinner("Drive에서 보고서 복사 중..."):
            # 4. Copy previous report in Drive (avoids SA storage quota issue)
            service = get_drive_service()
            copy_metadata = {
                "name": target_fname,
                "parents": [month_folder_id],
            }
            copied = service.files().copy(
                fileId=prev_file["id"],
                body=copy_metadata,
                fields="id",
            ).execute()
            new_file_id = copied.get("id")

        with st.spinner("보고서 날짜 교체 중..."):
            # 5. Download the copied file
            tmp_new = os.path.join(tempfile.gettempdir(), target_fname)
            download_file_by_id(new_file_id, tmp_new)

            # 6. Replace dates locally
            generate_report(
                target_date,
                dry_run=False,
                prev_path=tmp_new,
                output_path=tmp_new,
            )

            # 7. Upload modified content back (update, not create)
            from googleapiclient.http import MediaFileUpload
            media = MediaFileUpload(tmp_new, resumable=True)
            service.files().update(
                fileId=new_file_id,
                media_body=media,
            ).execute()

        elapsed = time.time() - start_t
        mark_stage("stage_report", "completed", {"elapsed": elapsed}, date_str=date_str)
        append_log(f"일일보고서 Cloud 생성 완료: {target_fname} ({format_elapsed(elapsed)})")
        st.toast(f"보고서 생성 완료 ({format_elapsed(elapsed)})")
        st.success(f"생성됨: {target_fname}")

        # Download button
        with open(tmp_new, "rb") as f:
            st.download_button(
                "보고서 다운로드",
                data=f.read(),
                file_name=target_fname,
                mime="application/octet-stream",
                key="btn_download_report",
            )

        # Cleanup
        try:
            os.unlink(tmp_new)
        except OSError:
            pass

    except Exception as e:
        elapsed = time.time() - start_t
        mark_stage("stage_report", "failed", {"error": str(e)[-500:]}, date_str=date_str)
        st.error(f"보고서 생성 실패: {e}")
