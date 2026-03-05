"""
Dashboard shared utilities: state management, logging, script execution, dependencies.
"""
import os
import sys
import json
import subprocess
import datetime
import time
import traceback
from pathlib import Path
from typing import Optional, Tuple

from dashboard.drive_service import is_cloud

# Project root (parent of dashboard/)
PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXECUTION_DIR = PROJECT_ROOT / "execution"
DATA_DIR = PROJECT_ROOT / "data"
INPUTS_DIR = DATA_DIR / "inputs"
OUTPUTS_DIR = DATA_DIR / "outputs"
DASHBOARD_DIR = Path(__file__).resolve().parent
STATE_DIR = DASHBOARD_DIR / "state"
LOGS_DIR = DASHBOARD_DIR / "logs"
HISTORY_DIR = DASHBOARD_DIR / "history"
MAPPINGS_PATH = PROJECT_ROOT / "skills" / "order_processing" / "resources" / "mappings.json"

# Ensure dirs exist
for d in [STATE_DIR, LOGS_DIR, HISTORY_DIR, INPUTS_DIR, OUTPUTS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# Default master file path
DEFAULT_MASTER = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"


# --------------- Date utilities ---------------
def today_str(fmt: str = "%Y-%m-%d") -> str:
    return datetime.date.today().strftime(fmt)


def today_compact() -> str:
    return datetime.date.today().strftime("%Y%m%d")


def date_to_compact(date_str: str) -> str:
    """Convert YYYY-MM-DD to YYYYMMDD."""
    return date_str.replace("-", "")


def date_to_iso(compact_or_date) -> str:
    """Normalize any date input to YYYY-MM-DD string."""
    if isinstance(compact_or_date, datetime.date):
        return compact_or_date.strftime("%Y-%m-%d")
    s = str(compact_or_date).replace("-", "")
    if len(s) == 8:
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    return str(compact_or_date)


# --------------- Settings ---------------
SETTINGS_FILE = STATE_DIR / "settings.json"
DEFAULT_SETTINGS = {
    "master_file": DEFAULT_MASTER,
    "drive_folder_id": "1rD5u5OwwmawOSy_3Iu169ltGQwQV1DJW",
    "delivery_list_drive_path": r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\배송리스트",
    "timeouts": {
        "default": 300,
        "sikbom": 600,
        "auction": 120,
        "ledger": 600,
        "helo": 600,
    },
}


def load_settings() -> dict:
    if SETTINGS_FILE.exists():
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            merged = {**DEFAULT_SETTINGS, **saved}
            # Ensure nested dicts merge properly
            if "timeouts" in DEFAULT_SETTINGS and "timeouts" not in saved:
                merged["timeouts"] = DEFAULT_SETTINGS["timeouts"]
            return merged
        except (json.JSONDecodeError, Exception):
            return dict(DEFAULT_SETTINGS)
    return dict(DEFAULT_SETTINGS)


def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def get_timeout(key: str = "default") -> int:
    s = load_settings()
    return s.get("timeouts", {}).get(key, 300)


# --------------- Stage definitions with dependencies ---------------
STAGES = [
    {
        "id": "stage0_normalize",
        "name": "발주 정규화",
        "icon": "0",
        "requires": [],
        "script": "classify_and_learn.py",
        "page": "1. 발주 정규화",
    },
    {
        "id": "stage1",
        "name": "발주처리 & 배송리스트",
        "sub": "(1차출고 입력)",
        "icon": "1",
        "requires": [],
        "script": "process_orders.py",
        "page": "2. 발주 처리",
    },
    {
        "id": "stage2_order",
        "name": "발주시트 입력",
        "icon": "2",
        "requires": [],
        "script": "update_order_sheet.py",
        "page": "3. 발주시트 & 명세서",
    },
    {
        "id": "stage2_invoice",
        "name": "가명세서 생성",
        "icon": "3",
        "requires": [],
        "script": "generate_invoices.py",
        "page": "3. 발주시트 & 명세서",
    },
    {
        "id": "stage3_price",
        "name": "단가표 생성",
        "icon": "4",
        "requires": [],
        "script": "create_price_sheet.py",
        "page": "4. 단가표 관리",
    },
    {
        "id": "stage3_sikbom",
        "name": "식봄 가격 조회",
        "icon": "5",
        "requires": [],
        "script": "fill_sikbom_targeted.py",
        "page": "4. 단가표 관리",
    },
    {
        "id": "stage5_auction",
        "name": "경매가 조회",
        "icon": "6",
        "requires": [],
        "script": "fill_auction_from_api.py",
        "page": "4. 단가표 관리",
    },
    {
        "id": "stage4_helo",
        "name": "동원2차출고 입력",
        "sub": "(다음날짜 재고현황 생성)",
        "icon": "7",
        "requires": [],
        "script": "scrape_helo.py",
        "page": "7. 동원발주 스크래핑",
    },
    {
        "id": "stage4_second",
        "name": "2차 발주 메시지",
        "sub": "(개발중)",
        "icon": "8",
        "requires": [],
        "script": "second_order.py",
        "page": "7. 동원발주 스크래핑",
        "dev": True,
    },
    {
        "id": "stage_receipt",
        "name": "매입가 입력 (영수증)",
        "icon": "9",
        "requires": [],
        "script": "fill_receipt_prices.py",
        "page": "8. 매입가 입력 (영수증)",
    },
    {
        "id": "stage7_fill",
        "name": "매입가/판매가 계산",
        "icon": "10",
        "requires": [],
        "script": "fill_prices.py",
        "page": "5. 매입가 & 판매가",
    },
    {
        "id": "stage8_final",
        "name": "최종 명세서 생성",
        "icon": "11",
        "requires": [],
        "script": "generate_final_invoices.py",
        "page": "6. 최종 명세서",
    },
    {
        "id": "stage_report",
        "name": "일일보고서 생성",
        "icon": "12",
        "requires": [],
        "script": "generate_daily_report.py",
        "page": "13. 일일보고서",
    },
]

STAGE_MAP = {s["id"]: s for s in STAGES}


def check_prerequisites(stage_id: str, date_str: str = None) -> Tuple[bool, list]:
    """Check if all prerequisites for a stage are completed. Returns (ok, missing_list)."""
    stage = STAGE_MAP.get(stage_id)
    if not stage:
        return True, []
    missing = []
    for req_id in stage.get("requires", []):
        status = get_stage_status(req_id, date_str)
        if status != "completed":
            req = STAGE_MAP.get(req_id, {})
            missing.append(req.get("name", req_id))
    return len(missing) == 0, missing


# --------------- State (daily progress) ---------------
def _state_path(date_str: str = None) -> Path:
    ds = date_to_compact(date_str) if date_str else today_compact()
    return STATE_DIR / f"progress_{ds}.json"


def load_progress(date_str: str = None) -> dict:
    """Load progress, merging local state dir and Drive-synced source."""
    ds = date_to_compact(date_str) if date_str else today_compact()

    # Load from local state dir
    local_data = {}
    p = _state_path(date_str)
    if p.exists():
        try:
            with open(p, "r", encoding="utf-8") as f:
                local_data = json.load(f)
        except (json.JSONDecodeError, Exception):
            pass

    # Load from Drive
    drive_data = {}
    try:
        from dashboard.drive_service import load_progress_from_drive
        drive_data = load_progress_from_drive(ds)
    except Exception:
        pass

    if not drive_data:
        return local_data
    if not local_data:
        return drive_data

    # Merge: per stage, newer timestamp wins
    merged = {}
    for key in set(local_data.keys()) | set(drive_data.keys()):
        local_entry = local_data.get(key, {})
        drive_entry = drive_data.get(key, {})
        if not local_entry:
            merged[key] = drive_entry
        elif not drive_entry:
            merged[key] = local_entry
        else:
            local_ts = local_entry.get("timestamp", "")
            drive_ts = drive_entry.get("timestamp", "")
            merged[key] = drive_entry if drive_ts >= local_ts else local_entry
    return merged


def save_progress(progress: dict, date_str: str = None):
    """Save progress locally and sync to Drive."""
    ds = date_to_compact(date_str) if date_str else today_compact()

    # Save to local state dir
    p = _state_path(date_str)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)

    # Sync to Drive (best-effort)
    try:
        from dashboard.drive_service import sync_progress_to_drive
        sync_progress_to_drive(ds, progress)
    except Exception:
        pass


def mark_stage(stage_id: str, status: str = "completed", extra: dict = None, date_str: str = None):
    """Mark a stage as completed/failed/running with timestamp."""
    prog = load_progress(date_str)
    entry = {
        "status": status,
        "timestamp": datetime.datetime.now().isoformat(),
    }
    if extra:
        entry.update(extra)
    prog[stage_id] = entry
    save_progress(prog, date_str)
    # Also append to history
    _append_history(stage_id, status, extra, date_str)


def get_stage_status(stage_id: str, date_str: str = None) -> str:
    prog = load_progress(date_str)
    return prog.get(stage_id, {}).get("status", "pending")


def count_completed(date_str: str = None) -> int:
    prog = load_progress(date_str)
    return sum(1 for v in prog.values() if isinstance(v, dict) and v.get("status") == "completed")


# --------------- Execution History ---------------
def _history_path(date_str: str = None) -> Path:
    ds = date_str or today_compact()
    ds = date_to_compact(ds) if "-" in ds else ds
    return HISTORY_DIR / f"history_{ds}.jsonl"


def _append_history(stage_id: str, status: str, extra: dict = None, date_str: str = None):
    entry = {
        "stage_id": stage_id,
        "status": status,
        "timestamp": datetime.datetime.now().isoformat(),
        **(extra or {}),
    }
    try:
        with open(_history_path(date_str), "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


def load_history(date_str: str = None) -> list:
    p = _history_path(date_str)
    if not p.exists():
        return []
    entries = []
    try:
        with open(p, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    entries.append(json.loads(line))
    except Exception:
        pass
    return entries


# --------------- Logging ---------------
def _log_path(date_str: str = None) -> Path:
    ds = date_str or today_compact()
    ds = date_to_compact(ds) if "-" in ds else ds
    return LOGS_DIR / f"log_{ds}.txt"


def append_log(message: str, level: str = "INFO", date_str: str = None):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] [{level}] {message}\n"
    try:
        with open(_log_path(date_str), "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def read_log(date_str: str = None, tail: int = 200) -> str:
    p = _log_path(date_str)
    if not p.exists():
        return ""
    try:
        with open(p, "r", encoding="utf-8") as f:
            lines = f.readlines()
        return "".join(lines[-tail:])
    except Exception:
        return ""


# --------------- Script Execution ---------------
def _build_env() -> dict:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    py_path = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = str(PROJECT_ROOT) + (os.pathsep + py_path if py_path else "")
    return env


def run_script(
    script_name: str,
    args: list = None,
    cwd: str = None,
    timeout: int = None,
) -> Tuple[bool, str, str, float]:
    """
    Run an execution/ script and capture stdout/stderr.
    Returns (success, stdout, stderr, elapsed_seconds).
    """
    script_path = EXECUTION_DIR / script_name
    if not script_path.exists():
        msg = f"스크립트를 찾을 수 없습니다: {script_path}"
        append_log(msg, "ERROR")
        return False, "", msg, 0.0

    if timeout is None:
        timeout = get_timeout("default")

    cmd = [sys.executable, str(script_path)] + (args or [])
    work_dir = cwd or str(PROJECT_ROOT)
    env = _build_env()

    args_str = " ".join(args or [])
    append_log(f"START: {script_name} {args_str}")
    start_t = time.time()

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            cwd=work_dir,
            timeout=timeout,
            env=env,
            encoding="utf-8",
            errors="replace",
        )
        elapsed = time.time() - start_t
        success = result.returncode == 0
        status_str = "OK" if success else f"FAIL(rc={result.returncode})"
        append_log(f"END: {script_name} [{status_str}] ({elapsed:.1f}s)")

        # Log full output to file, truncate in memory
        if result.stdout:
            append_log(f"STDOUT({len(result.stdout)} chars): {result.stdout[-800:]}")
        if result.stderr and not success:
            append_log(f"STDERR: {result.stderr[-800:]}", "ERROR")

        return success, result.stdout, result.stderr, elapsed

    except subprocess.TimeoutExpired:
        elapsed = time.time() - start_t
        msg = f"시간 초과 ({timeout}초). 네트워크 또는 파일 접근 문제일 수 있습니다."
        append_log(f"TIMEOUT: {script_name} ({timeout}s)", "ERROR")
        return False, "", msg, elapsed
    except FileNotFoundError:
        elapsed = time.time() - start_t
        msg = f"Python 실행 파일을 찾을 수 없습니다: {sys.executable}"
        append_log(msg, "ERROR")
        return False, "", msg, elapsed
    except Exception as e:
        elapsed = time.time() - start_t
        msg = f"{type(e).__name__}: {e}"
        append_log(f"ERROR: {script_name} - {msg}", "ERROR")
        return False, "", msg, elapsed


def run_script_streaming(script_name: str, args: list = None, cwd: str = None):
    """Start script as Popen process for streaming output. Returns Popen object."""
    script_path = EXECUTION_DIR / script_name
    cmd = [sys.executable, "-u", str(script_path)] + (args or [])
    work_dir = cwd or str(PROJECT_ROOT)
    env = _build_env()

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd=work_dir,
        env=env,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    return proc


# --------------- File helpers ---------------
def check_master_file(settings: dict = None) -> Tuple[bool, str]:
    """Check if master file exists and is not locked by Excel."""
    s = settings or load_settings()

    # Cloud mode: check via Drive API
    if is_cloud():
        try:
            from dashboard.drive_service import check_drive_connection
            ok, msg = check_drive_connection()
            return ok, msg
        except Exception as e:
            return False, f"Drive 연결 확인 실패: {e}"

    # Local mode: file system check
    master = s.get("master_file", "")
    if not master:
        return False, "마스터 파일 경로가 설정되지 않았습니다."

    p = Path(master)
    if not p.exists():
        return False, f"마스터 파일을 찾을 수 없습니다: {p.name}"

    # Windows: check for Excel lock file (~$filename.xlsx)
    lock_file = p.parent / f"~${p.name}"
    if lock_file.exists():
        return False, "마스터 파일이 엑셀에서 열려 있습니다. 엑셀을 닫아주세요."

    try:
        # Try opening for write to verify not locked
        with open(p, "r+b") as f:
            f.seek(0)
            f.read(4)
        return True, "OK"
    except PermissionError:
        # Google Drive sync can cause transient locks — retry once
        import time
        time.sleep(2)
        lock_file2 = p.parent / f"~${p.name}"
        if lock_file2.exists():
            return False, "마스터 파일이 엑셀에서 열려 있습니다. 엑셀을 닫아주세요."
        try:
            with open(p, "r+b") as f:
                f.read(4)
            return True, "OK"
        except PermissionError:
            return False, "마스터 파일이 잠겨 있습니다. 엑셀 또는 Drive 동기화 완료 후 재시도하세요."
    except Exception as e:
        return False, f"파일 접근 오류: {e}"


def check_master_before_run(settings: dict = None) -> bool:
    """Streamlit-aware master check. Shows error and returns False if locked."""
    import streamlit as st
    ok, msg = check_master_file(settings)
    if not ok:
        st.error(f"⚠️ {msg}")
    return ok


def resolve_master_path(settings: dict = None) -> Tuple[str, bool]:
    """
    Resolve master file path for current environment.
    Returns (local_path, is_from_drive).
    - Local mode: returns settings["master_file"] as-is
    - Cloud mode: downloads from Drive to temp, returns temp path
    """
    s = settings or load_settings()

    if not is_cloud():
        return s.get("master_file", DEFAULT_MASTER), False

    # Cloud mode: download from Drive
    from dashboard.drive_service import download_master_file
    tmp_path = download_master_file()
    return tmp_path, True


def sync_master_back(local_path: str):
    """Upload master file back to Drive (only in cloud mode, no-op locally)."""
    if not is_cloud():
        return

    from dashboard.drive_service import upload_master_file
    upload_master_file(local_path)
    append_log("마스터 파일 Drive 업로드 완료 (cloud sync)")


def list_output_files(date_str: str = None, prefix: str = "") -> list:
    """List output files for a given date."""
    ds = date_to_compact(date_str) if date_str else today_compact()
    results = []
    if OUTPUTS_DIR.exists():
        for f in OUTPUTS_DIR.iterdir():
            if ds in f.name or (prefix and f.name.startswith(prefix)):
                results.append(f)
        for d in OUTPUTS_DIR.iterdir():
            if d.is_dir() and ds in d.name:
                for f in d.iterdir():
                    results.append(f)
    return sorted(results, key=lambda x: x.name)


def get_processed_orders_path(date_str: str = None) -> Path:
    ds = date_to_compact(date_str) if date_str else today_compact()
    return INPUTS_DIR / f"processed_orders_{ds}.txt"


def get_orders_path(date_str: str = None) -> Path:
    ds = date_to_compact(date_str) if date_str else today_compact()
    return INPUTS_DIR / f"orders_{ds}.txt"


# --------------- UI Helpers ---------------
def render_stage_badge(status: str, timestamp: str = "") -> str:
    """Return HTML badge with color dot for stage status (Toss-style)."""
    short_ts = timestamp[11:16] if len(timestamp) > 16 else ""
    if status == "completed":
        return f'<span class="badge"><span class="dot dot-ok"></span>완료 {short_ts}</span>'
    elif status == "failed":
        return '<span class="badge"><span class="dot dot-fail"></span>실패</span>'
    elif status == "running":
        return '<span class="badge"><span class="dot dot-run"></span>실행중</span>'
    return '<span class="badge"><span class="dot dot-wait"></span>대기</span>'


def format_elapsed(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}초"
    minutes = int(seconds // 60)
    secs = seconds % 60
    return f"{minutes}분 {secs:.0f}초"


# --------------- Execution Lock (prevent concurrent runs) ---------------
LOCK_FILE = STATE_DIR / ".execution_lock"


def acquire_lock(script_name: str) -> bool:
    """Try to acquire execution lock. Returns False if another script is running."""
    if LOCK_FILE.exists():
        try:
            with open(LOCK_FILE, "r") as f:
                lock_data = json.load(f)
            # Check if lock is stale (>15 min old)
            lock_time = datetime.datetime.fromisoformat(lock_data.get("timestamp", "2000-01-01"))
            if (datetime.datetime.now() - lock_time).total_seconds() > 900:
                # Stale lock, remove
                LOCK_FILE.unlink(missing_ok=True)
            else:
                return False
        except Exception:
            LOCK_FILE.unlink(missing_ok=True)

    try:
        with open(LOCK_FILE, "w") as f:
            json.dump({
                "script": script_name,
                "timestamp": datetime.datetime.now().isoformat(),
            }, f)
        return True
    except Exception:
        return False


def release_lock():
    """Release execution lock."""
    LOCK_FILE.unlink(missing_ok=True)


def get_lock_info() -> Optional[dict]:
    """Get current lock info or None if not locked."""
    if not LOCK_FILE.exists():
        return None
    try:
        with open(LOCK_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return None


# --------------- Master File Backup ---------------
BACKUP_DIR = DASHBOARD_DIR / "backups"
BACKUP_DIR.mkdir(exist_ok=True)


def backup_master_file(settings: dict = None) -> Optional[Path]:
    """Create a backup of the master file before modifications."""
    s = settings or load_settings()
    master = Path(s.get("master_file", ""))
    if not master.exists():
        return None

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"backup_{ts}_{master.name}"

    try:
        import shutil
        shutil.copy2(str(master), str(backup_path))
        append_log(f"백업 생성: {backup_path.name}")

        # Keep only last 5 backups
        backups = sorted(BACKUP_DIR.glob(f"backup_*_{master.name}"), reverse=True)
        for old in backups[5:]:
            old.unlink()
            append_log(f"오래된 백업 삭제: {old.name}")

        return backup_path
    except Exception as e:
        append_log(f"백업 실패: {e}", "ERROR")
        return None


def list_backups() -> list:
    """List available backup files."""
    if not BACKUP_DIR.exists():
        return []
    return sorted(BACKUP_DIR.glob("backup_*"), reverse=True)


def run_script_safe(
    script_name: str,
    args: list = None,
    cwd: str = None,
    timeout: int = None,
    backup: bool = False,
) -> Tuple[bool, str, str, float]:
    """
    Run script with execution lock and optional backup.
    Returns (success, stdout, stderr, elapsed).
    """
    # Check lock
    lock_info = get_lock_info()
    if lock_info:
        running = lock_info.get("script", "unknown")
        return False, "", f"다른 스크립트가 실행 중입니다: {running}", 0.0

    if not acquire_lock(script_name):
        return False, "", "실행 잠금 획득 실패", 0.0

    try:
        # Optional backup
        if backup:
            backup_master_file()

        # Run
        return run_script(script_name, args, cwd, timeout)
    finally:
        release_lock()
