"""Page 7: Settings - Paths, tokens, mappings, system info."""
import streamlit as st
import json
from pathlib import Path
from dashboard.utils import (
    load_settings, save_settings, MAPPINGS_PATH, check_master_file,
    PROJECT_ROOT, STATE_DIR, LOGS_DIR, HISTORY_DIR, BACKUP_DIR,
    save_progress, DEFAULT_SETTINGS, backup_master_file, list_backups,
)
from dashboard.drive_service import is_cloud, check_drive_connection


def render():
    st.header("9. 설정")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["파일 경로", "토큰 상태", "타임아웃", "매핑 편집", "백업", "시스템"])

    with tab1:
        _render_paths()
    with tab2:
        _render_tokens()
    with tab3:
        _render_timeouts()
    with tab4:
        _render_mappings()
    with tab5:
        _render_backups()
    with tab6:
        _render_system()


def _render_paths():
    st.subheader("파일 경로 설정")
    settings = load_settings()
    cloud_mode = is_cloud()

    # --- Cloud mode: Drive connection status ---
    if cloud_mode:
        st.markdown("##### Google Drive 연결 (Cloud 모드)")
        ok, msg = check_drive_connection()
        if ok:
            st.success(f"Drive: {msg}")
        else:
            st.error(f"Drive: {msg}")
            st.caption("Streamlit Secrets에 gcp_service_account 및 drive 섹션을 설정해주세요.")

        st.divider()
        st.markdown("##### Drive 파일 ID 설정")
        st.caption("Streamlit Cloud > Settings > Secrets에서 관리됩니다.")

        try:
            from dashboard.drive_service import get_drive_file_ids
            file_ids = get_drive_file_ids()
            for key, val in file_ids.items():
                st.text_input(key, value=val or "", disabled=True, key=f"drive_id_{key}")
        except Exception:
            st.info("Drive 설정이 아직 구성되지 않았습니다.")

        return

    # --- Local mode: file browser ---
    st.markdown("##### 마스터 파일")

    # 매입단가_자료 폴더에서 xlsx 파일 스캔 (현재 폴더만)
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
        # 셀렉트박스용 옵션 구성
        options = [str(f) for f in xlsx_files]
        labels = []
        for f in xlsx_files:
            size_mb = f.stat().st_size / (1024 * 1024)
            name = f.name.replace("도크발주관리데이터", "").replace(".xlsx", "").strip(" -")
            label = f.name if not name else f.name
            labels.append(f"{label}  ({size_mb:.1f}MB)")

        # 현재 선택된 파일의 인덱스 찾기
        try:
            current_idx = options.index(current_master)
        except ValueError:
            current_idx = 0

        selected_idx = st.selectbox(
            "파일 선택",
            range(len(options)),
            index=current_idx,
            format_func=lambda i: labels[i],
            key="set_master_select",
        )
        master = options[selected_idx]
    else:
        master = st.text_input("마스터 파일 경로", value=current_master, key="set_master")

    # 현재 활성 파일 표시
    if master:
        p = Path(master)
        if p.exists():
            is_production = p.name == "도크발주관리데이터.xlsx"
            badge = "운영" if is_production else "테스트/백업"
            color = "#4ADE80" if is_production else "#FBBF24"
            st.markdown(
                f"현재: <span style='color:{color}; font-weight:600'>[{badge}]</span> `{p.name}`",
                unsafe_allow_html=True,
            )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("저장", type="primary", key="btn_save_paths"):
            settings["master_file"] = master
            save_settings(settings)
            st.success("저장 완료!")
            st.rerun()

    with col2:
        if st.button("운영 파일로 초기화", key="btn_reset_paths"):
            settings["master_file"] = DEFAULT_SETTINGS["master_file"]
            save_settings(settings)
            st.success("운영 파일로 초기화됨")
            st.rerun()

    # Verify
    st.divider()
    ok, msg = check_master_file({"master_file": master})
    if ok:
        size_mb = Path(master).stat().st_size / (1024 * 1024) if Path(master).exists() else 0
        st.success(f"마스터 파일: OK ({size_mb:.1f}MB)")
    else:
        st.error(f"마스터 파일: {msg}")

    # --- 기타 경로 ---
    st.divider()
    st.markdown("##### 기타 경로")
    drive_folder = st.text_input("Drive 배송리스트 폴더 ID", value=settings.get("drive_folder_id", ""), key="set_drive")
    delivery_path = st.text_input("배송리스트 Drive 경로", value=settings.get("delivery_list_drive_path", ""), key="set_dlpath")

    if st.button("기타 경로 저장", key="btn_save_other_paths"):
        settings["drive_folder_id"] = drive_folder
        settings["delivery_list_drive_path"] = delivery_path
        save_settings(settings)
        st.success("저장 완료!")


def _render_tokens():
    st.subheader("토큰 상태")

    # Kakao
    kakao_path = PROJECT_ROOT / "kakao_token.json"
    st.markdown("### 카카오톡")
    if kakao_path.exists():
        try:
            with open(kakao_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            has_access = bool(data.get("access_token"))
            has_refresh = bool(data.get("refresh_token"))

            c1, c2 = st.columns(2)
            if has_access:
                c1.success("Access Token: 존재")
            else:
                c1.warning("Access Token: 없음")
            if has_refresh:
                c2.success("Refresh Token: 존재")
            else:
                c2.error("Refresh Token: 없음")
        except Exception as e:
            st.error(f"토큰 파일 오류: {e}")
    else:
        st.warning("kakao_token.json 없음")

    st.divider()

    # Google
    st.markdown("### Google")
    cred_path = PROJECT_ROOT / "credentials.json"
    token_path = PROJECT_ROOT / "token.json"

    c1, c2 = st.columns(2)
    c1.success("credentials.json: 존재") if cred_path.exists() else c1.warning("credentials.json: 없음")
    c2.success("token.json: 존재") if token_path.exists() else c2.warning("token.json: 없음")

    st.divider()

    # .env
    st.markdown("### 환경변수 (.env)")
    env_path = PROJECT_ROOT / ".env"
    if env_path.exists():
        with open(env_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        keys = [l.split("=")[0].strip() for l in lines if "=" in l and not l.strip().startswith("#")]

        required_keys = ["KAKAO_REST_API_KEY", "GEMINI_API_KEY"]
        for k in required_keys:
            if k in keys:
                st.success(f"{k}: 설정됨")
            else:
                st.warning(f"{k}: 없음")

        other_keys = [k for k in keys if k not in required_keys]
        if other_keys:
            st.caption(f"기타 키: {', '.join(other_keys)}")
    else:
        st.error(".env 파일 없음")


def _render_timeouts():
    st.subheader("스크립트 타임아웃 설정")
    st.caption("각 스크립트의 최대 실행 시간 (초)")

    settings = load_settings()
    timeouts = settings.get("timeouts", DEFAULT_SETTINGS["timeouts"])

    t_default = st.number_input("기본 타임아웃", value=timeouts.get("default", 300), min_value=30, max_value=3600, step=30, key="t_default")
    t_sikbom = st.number_input("식봄 조회 (Selenium)", value=timeouts.get("sikbom", 600), min_value=60, max_value=3600, step=60, key="t_sikbom")
    t_auction = st.number_input("경매가 API", value=timeouts.get("auction", 120), min_value=30, max_value=600, step=30, key="t_auction")
    t_ledger = st.number_input("거래원장 생성", value=timeouts.get("ledger", 600), min_value=60, max_value=3600, step=60, key="t_ledger")

    if st.button("타임아웃 저장", key="btn_save_timeouts"):
        settings["timeouts"] = {
            "default": t_default,
            "sikbom": t_sikbom,
            "auction": t_auction,
            "ledger": t_ledger,
        }
        save_settings(settings)
        st.success("저장 완료!")


def _render_mappings():
    st.subheader("매핑 설정")

    if not MAPPINGS_PATH.exists():
        st.error(f"매핑 파일 없음: {MAPPINGS_PATH}")
        return

    with open(MAPPINGS_PATH, "r", encoding="utf-8") as f:
        mappings = json.load(f)

    section = st.selectbox("섹션 선택", list(mappings.keys()), key="mapping_section")

    if section:
        data = mappings[section]

        # Show summary
        if isinstance(data, dict):
            st.caption(f"키: {len(data)}개")
        elif isinstance(data, list):
            st.caption(f"항목: {len(data)}개")

        edited = st.text_area(
            f"JSON 편집: {section}",
            value=json.dumps(data, ensure_ascii=False, indent=2),
            height=400,
            key=f"mapping_edit_{section}",
        )

        col1, col2 = st.columns(2)
        with col1:
            if st.button("매핑 저장", key="btn_save_mapping"):
                try:
                    parsed = json.loads(edited)
                    mappings[section] = parsed
                    with open(MAPPINGS_PATH, "w", encoding="utf-8") as f:
                        json.dump(mappings, f, ensure_ascii=False, indent=4)
                    st.success(f"'{section}' 저장 완료!")
                except json.JSONDecodeError as e:
                    st.error(f"JSON 형식 오류: {e}")

        with col2:
            # Validate JSON
            try:
                json.loads(edited)
                st.success("JSON 유효")
            except json.JSONDecodeError:
                st.error("JSON 형식 오류")


def _render_backups():
    st.subheader("마스터 파일 백업")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("지금 백업 생성", type="primary", key="btn_create_backup"):
            path = backup_master_file()
            if path:
                st.success(f"백업 생성됨: {path.name}")
            else:
                st.error("백업 생성 실패")

    with col2:
        st.caption("마스터 파일 수정 전 자동 백업됩니다.")

    st.divider()
    st.markdown("### 백업 목록")
    backups = list_backups()

    if not backups:
        st.info("백업 파일이 없습니다.")
    else:
        for b in backups[:10]:
            size_mb = b.stat().st_size / (1024 * 1024)
            mtime = b.stat().st_mtime
            import datetime
            ts = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            c1, c2, c3 = st.columns([3, 1, 1])
            c1.markdown(f"`{b.name}`")
            c2.caption(f"{size_mb:.1f}MB | {ts}")
            with open(b, "rb") as f:
                c3.download_button("↓", data=f.read(), file_name=b.name, key=f"dl_bk_{b.stem}")

        # Total size
        total = sum(b.stat().st_size for b in backups)
        st.caption(f"총 {len(backups)}개 백업, {total / (1024*1024):.1f}MB")


def _render_system():
    st.subheader("시스템 정보")

    import sys
    cloud_mode = is_cloud()
    c1, c2, c3 = st.columns(3)
    c1.info(f"Python: {sys.version.split()[0]}")
    c2.info(f"프로젝트: {PROJECT_ROOT.name}")
    if cloud_mode:
        c3.warning("모드: Cloud")
    else:
        c3.success("모드: Local")

    # Disk usage
    st.divider()
    st.markdown("### 저장소 사용량")

    dirs_to_check = {
        "실행 로그": LOGS_DIR,
        "상태 파일": STATE_DIR,
        "실행 이력": HISTORY_DIR,
        "출력 파일": PROJECT_ROOT / "data" / "outputs",
    }

    for name, d in dirs_to_check.items():
        if d.exists():
            total = sum(f.stat().st_size for f in d.rglob("*") if f.is_file())
            total_mb = total / (1024 * 1024)
            file_count = sum(1 for f in d.rglob("*") if f.is_file())
            st.markdown(f"- **{name}**: {total_mb:.1f}MB ({file_count}개 파일)")

    # Maintenance
    st.divider()
    st.markdown("### 유지보수")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("오늘 진행 상태 초기화", key="btn_reset_state"):
            save_progress({})
            st.success("초기화 완료!")
            st.rerun()

    with col2:
        if st.button("7일 이전 로그 정리", key="btn_clean_logs"):
            import datetime
            cutoff = datetime.date.today() - datetime.timedelta(days=7)
            cutoff_str = cutoff.strftime("%Y%m%d")
            cleaned = 0
            for f in LOGS_DIR.glob("log_*.txt"):
                date_part = f.stem.replace("log_", "")
                if date_part < cutoff_str:
                    f.unlink()
                    cleaned += 1
            st.success(f"{cleaned}개 로그 파일 삭제됨")

    with col3:
        if st.button("캐시 초기화", key="btn_clear_cache"):
            st.cache_data.clear()
            st.success("캐시 초기화됨")
            st.rerun()
