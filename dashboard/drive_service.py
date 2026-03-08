"""
Google Drive API service for Streamlit Cloud deployment.
Provides dual-mode operation: local (G: File Stream) vs cloud (Drive API).
"""
import os
import tempfile
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Singleton Drive service instances
_drive_service = None      # Service Account (read/update)
_user_drive_service = None  # OAuth user (create/copy — SA has no storage quota)


def is_cloud() -> bool:
    """Detect if running in cloud environment (no local Google Drive mount)."""
    # Windows with G: drive mount = local mode
    if os.name == "nt":
        g_drive = Path(r"G:\내 드라이브")
        if g_drive.exists():
            return False
    # Everything else = cloud mode
    return True


def _get_st_secrets():
    """Access Streamlit st.secrets directly (AttrDict, not plain dict)."""
    try:
        import streamlit as st
        return st.secrets
    except Exception:
        return None


def get_drive_service():
    """Get singleton Google Drive API service client using service account credentials."""
    global _drive_service
    if _drive_service is not None:
        return _drive_service

    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        secrets = _get_st_secrets()
        if secrets is None:
            raise ValueError("Streamlit secrets not available")

        sa_info = dict(secrets["gcp_service_account"])

        creds = service_account.Credentials.from_service_account_info(
            sa_info,
            scopes=["https://www.googleapis.com/auth/drive"],
        )
        _drive_service = build("drive", "v3", credentials=creds)
        logger.info("Drive API service initialized (service account)")
        return _drive_service

    except Exception as e:
        logger.error(f"Failed to initialize Drive service: {e}")
        raise


def get_user_drive_service():
    """Get Drive service using OAuth user credentials (for file creation).

    Service Accounts have no storage quota → files().create/copy() fails.
    OAuth user credentials are needed for operations that create new files.

    Requires secrets.toml:
        [google_oauth]
        client_id = "..."
        client_secret = "..."
        refresh_token = "..."
    """
    global _user_drive_service
    if _user_drive_service is not None:
        return _user_drive_service

    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build

        secrets = _get_st_secrets()
        if secrets is None or "google_oauth" not in secrets:
            return None

        oauth = dict(secrets["google_oauth"])
        creds = Credentials(
            token=None,
            refresh_token=oauth["refresh_token"],
            client_id=oauth["client_id"],
            client_secret=oauth["client_secret"],
            token_uri="https://oauth2.googleapis.com/token",
        )

        _user_drive_service = build("drive", "v3", credentials=creds)
        logger.info("Drive API service initialized (OAuth user)")
        return _user_drive_service
    except Exception as e:
        logger.warning(f"OAuth Drive service not available: {e}")
        return None


def get_drive_file_ids() -> dict:
    """Get configured Drive file/folder IDs from Streamlit secrets."""
    try:
        secrets = _get_st_secrets()
        if secrets is None:
            return {}
        return dict(secrets["drive"])
    except Exception:
        return {}


def get_effective_master_file_id() -> str:
    """Get the active master file ID (override > secrets default)."""
    try:
        import streamlit as st
        override = st.session_state.get("master_file_id_override")
        if override:
            return override
    except Exception:
        pass

    file_ids = get_drive_file_ids()
    file_id = file_ids.get("master_file_id")
    if not file_id:
        raise ValueError("master_file_id not configured in secrets [drive] section")
    return file_id


def set_master_file_override(file_id: str, file_name: str = ""):
    """Set master file override (session state + Drive config for persistence)."""
    try:
        import streamlit as st
        st.session_state["master_file_id_override"] = file_id
        st.session_state["master_file_name_override"] = file_name
    except Exception:
        pass

    # Persist to Drive config
    try:
        _save_master_config(file_id, file_name)
    except Exception as e:
        logger.warning(f"Failed to persist master config: {e}")


def clear_master_file_override():
    """Clear override, revert to secrets default."""
    try:
        import streamlit as st
        st.session_state.pop("master_file_id_override", None)
        st.session_state.pop("master_file_name_override", None)
    except Exception:
        pass

    try:
        _save_master_config("", "")
    except Exception:
        pass


def load_master_override_on_start():
    """Load persisted master override into session state (call once on app start)."""
    try:
        import streamlit as st
        if "master_file_id_override" in st.session_state:
            return  # Already loaded

        config = _load_master_config()
        fid = config.get("master_file_id")
        if fid:
            st.session_state["master_file_id_override"] = fid
            st.session_state["master_file_name_override"] = config.get("master_file_name", "")
    except Exception:
        pass


def _save_master_config(file_id: str, file_name: str):
    """Save master file selection to Firestore (SA has no Drive storage quota)."""
    try:
        from dashboard.firebase_service import get_firestore_client
        db = get_firestore_client()
        db.collection("dashboardConfig").document("masterFile").set({
            "master_file_id": file_id,
            "master_file_name": file_name,
        })
    except Exception as e:
        logger.warning(f"Firestore config save failed: {e}")


def _load_master_config() -> dict:
    """Load master file selection from Firestore."""
    try:
        from dashboard.firebase_service import get_firestore_client
        db = get_firestore_client()
        doc = db.collection("dashboardConfig").document("masterFile").get()
        if doc.exists:
            data = doc.to_dict()
            if data.get("master_file_id"):
                return data
        return {}
    except Exception:
        return {}


def list_xlsx_in_master_folder() -> list:
    """List xlsx files in the 매입단가_자료 folder (price_data_folder_id)."""
    service = get_drive_service()
    file_ids = get_drive_file_ids()
    folder_id = file_ids.get("price_data_folder_id")
    if not folder_id:
        return []

    query = (
        f"'{folder_id}' in parents "
        f"and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
        f"and trashed = false"
    )
    results = service.files().list(
        q=query,
        fields="files(id, name, modifiedTime, size)",
        orderBy="modifiedTime desc",
        pageSize=20,
    ).execute()

    return results.get("files", [])


def download_master_file() -> str:
    """Download master file from Drive to a temp path. Returns the temp file path."""
    service = get_drive_service()
    file_id = get_effective_master_file_id()

    from googleapiclient.http import MediaIoBaseDownload
    import io

    # Get actual filename from Drive
    meta = service.files().get(fileId=file_id, fields="name").execute()
    filename = meta.get("name", "도크발주관리데이터.xlsx")

    request = service.files().get_media(fileId=file_id)
    tmp_dir = tempfile.gettempdir()
    tmp_path = os.path.join(tmp_dir, filename)

    fh = io.FileIO(tmp_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    fh.close()
    logger.info(f"Master file downloaded to {tmp_path}")
    return tmp_path


def upload_master_file(local_path: str) -> str:
    """Upload modified master file back to Drive. Returns the file ID."""
    service = get_drive_service()
    file_id = get_effective_master_file_id()

    from googleapiclient.http import MediaFileUpload

    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    updated = service.files().update(
        fileId=file_id,
        media_body=media,
    ).execute()

    logger.info(f"Master file uploaded: {updated.get('id')}")
    return updated.get("id", file_id)


def upload_file_to_folder(local_path: str, folder_id: str, filename: Optional[str] = None) -> str:
    """Upload a file to a specific Drive folder. Returns the new file ID.

    Uses OAuth user credentials if available (SA has no storage quota for create).
    """
    # Prefer OAuth user service for file creation
    service = get_user_drive_service() or get_drive_service()

    if filename is None:
        filename = os.path.basename(local_path)

    file_metadata = {
        "name": filename,
        "parents": [folder_id],
    }

    from googleapiclient.http import MediaFileUpload

    media = MediaFileUpload(local_path, resumable=True)
    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
    ).execute()

    logger.info(f"File uploaded: {filename} -> {created.get('id')}")
    return created.get("id", "")


def find_or_create_subfolder(parent_id: str, name: str) -> str:
    """Find or create a subfolder by name under parent. Returns folder ID."""
    service = get_drive_service()

    # Search for existing folder
    query = (
        f"'{parent_id}' in parents "
        f"and name = '{name}' "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )
    results = service.files().list(q=query, fields="files(id, name)", pageSize=1).execute()
    files = results.get("files", [])

    if files:
        return files[0]["id"]

    # Create new folder
    folder_metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = service.files().create(body=folder_metadata, fields="id").execute()
    logger.info(f"Created subfolder: {name} -> {folder.get('id')}")
    return folder.get("id", "")


def list_folder(folder_id: str) -> list:
    """List files in a Drive folder. Returns list of {id, name, mimeType, modifiedTime}."""
    service = get_drive_service()

    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(
        q=query,
        fields="files(id, name, mimeType, modifiedTime, size)",
        orderBy="modifiedTime desc",
        pageSize=100,
    ).execute()

    return results.get("files", [])


def download_file_by_id(file_id: str, local_path: str = None) -> str:
    """Drive 파일을 ID로 다운로드. local_path 미지정 시 temp 디렉토리."""
    service = get_drive_service()

    from googleapiclient.http import MediaIoBaseDownload
    import io

    # Get filename from Drive if no local_path specified
    if local_path is None:
        meta = service.files().get(fileId=file_id, fields="name").execute()
        tmp_dir = tempfile.gettempdir()
        local_path = os.path.join(tmp_dir, meta.get("name", f"drive_{file_id}"))

    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(local_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    fh.close()
    logger.info(f"File downloaded: {file_id} -> {local_path}")
    return local_path


def search_file_in_folder(folder_id: str, name_contains: str) -> list:
    """폴더 내에서 파일명 검색. 이전 보고서 찾기용.

    Returns:
        list of {id, name, modifiedTime}
    """
    service = get_drive_service()

    query = (
        f"'{folder_id}' in parents "
        f"and name contains '{name_contains}' "
        f"and trashed = false"
    )
    results = service.files().list(
        q=query,
        fields="files(id, name, modifiedTime)",
        orderBy="modifiedTime desc",
        pageSize=20,
    ).execute()

    return results.get("files", [])


# --------------- Progress Sync (Cloud ↔ Local) ---------------
import json as _json
import time as _time

PROGRESS_LOCAL_DRIVE = Path(
    r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\dashboard_progress"
)
_progress_folder_id = None
_progress_cache = {}  # {date_compact: (data, timestamp)}
_PROGRESS_CACHE_TTL = 30  # seconds


def _get_progress_folder_id() -> str:
    """Get or create dashboard_progress folder on Drive (next to master file)."""
    global _progress_folder_id
    if _progress_folder_id:
        return _progress_folder_id

    service = get_drive_service()
    file_ids = get_drive_file_ids()
    master_id = file_ids.get("master_file_id")
    if not master_id:
        raise ValueError("master_file_id not configured")

    meta = service.files().get(fileId=master_id, fields="parents").execute()
    parent_id = meta.get("parents", [None])[0]
    if not parent_id:
        raise ValueError("Cannot determine master file parent folder")

    _progress_folder_id = find_or_create_subfolder(parent_id, "dashboard_progress")
    return _progress_folder_id


def _upload_progress_api(date_compact: str, data: dict):
    """Upload progress JSON via Drive API."""
    folder_id = _get_progress_folder_id()
    filename = f"progress_{date_compact}.json"
    service = get_drive_service()

    query = f"'{folder_id}' in parents and name = '{filename}' and trashed = false"
    results = service.files().list(q=query, fields="files(id)", pageSize=1).execute()
    existing = results.get("files", [])

    tmp_path = os.path.join(tempfile.gettempdir(), filename)
    with open(tmp_path, "w", encoding="utf-8") as f:
        _json.dump(data, f, ensure_ascii=False, indent=2)

    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(tmp_path, mimetype="application/json")

    if existing:
        service.files().update(fileId=existing[0]["id"], media_body=media).execute()
    else:
        file_meta = {"name": filename, "parents": [folder_id]}
        service.files().create(body=file_meta, media_body=media, fields="id").execute()


def _download_progress_api(date_compact: str) -> dict:
    """Download progress JSON via Drive API."""
    try:
        folder_id = _get_progress_folder_id()
        filename = f"progress_{date_compact}.json"
        service = get_drive_service()

        query = f"'{folder_id}' in parents and name = '{filename}' and trashed = false"
        results = service.files().list(q=query, fields="files(id)", pageSize=1).execute()
        files = results.get("files", [])

        if not files:
            return {}

        from googleapiclient.http import MediaIoBaseDownload
        import io

        request = service.files().get_media(fileId=files[0]["id"])
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        buffer.seek(0)
        return _json.loads(buffer.read().decode("utf-8"))
    except Exception as e:
        logger.warning(f"Progress download failed: {e}")
        return {}


def sync_progress_to_drive(date_compact: str, data: dict):
    """Save progress. Local: file system, Cloud: Firestore."""
    try:
        if not is_cloud():
            PROGRESS_LOCAL_DRIVE.mkdir(parents=True, exist_ok=True)
            p = PROGRESS_LOCAL_DRIVE / f"progress_{date_compact}.json"
            with open(p, "w", encoding="utf-8") as f:
                _json.dump(data, f, ensure_ascii=False, indent=2)
        else:
            # Cloud: use Firestore (SA has no Drive storage quota for create)
            try:
                from dashboard.firebase_service import get_firestore_client
                db = get_firestore_client()
                db.collection("dashboardProgress").document(date_compact).set(data)
                _progress_cache[date_compact] = (data, _time.time())
            except Exception as e:
                logger.warning(f"Firestore progress sync failed: {e}")
                _progress_cache[date_compact] = (data, _time.time())
    except Exception as e:
        logger.warning(f"Progress sync failed: {e}")


def load_progress_from_drive(date_compact: str) -> dict:
    """Load progress. Local: file system, Cloud: Firestore with cache."""
    try:
        if not is_cloud():
            p = PROGRESS_LOCAL_DRIVE / f"progress_{date_compact}.json"
            if p.exists():
                with open(p, "r", encoding="utf-8") as f:
                    return _json.load(f)
            return {}
        else:
            now = _time.time()
            if date_compact in _progress_cache:
                cached_data, cached_ts = _progress_cache[date_compact]
                if now - cached_ts < _PROGRESS_CACHE_TTL:
                    return cached_data

            # Cloud: use Firestore
            try:
                from dashboard.firebase_service import get_firestore_client
                db = get_firestore_client()
                doc = db.collection("dashboardProgress").document(date_compact).get()
                data = doc.to_dict() if doc.exists else {}
            except Exception:
                data = {}
            _progress_cache[date_compact] = (data, _time.time())
            return data
    except Exception as e:
        logger.warning(f"Progress load failed: {e}")
        return {}


def check_drive_connection() -> tuple:
    """Test Drive API connection. Returns (ok: bool, message: str)."""
    if not is_cloud():
        return True, "로컬 모드 (Google Drive File Stream)"

    try:
        file_id = get_effective_master_file_id()

        service = get_drive_service()
        file_meta = service.files().get(
            fileId=file_id,
            fields="id, name, modifiedTime, size",
        ).execute()

        size_mb = int(file_meta.get("size", 0)) / (1024 * 1024)
        return True, f"연결됨: {file_meta['name']} ({size_mb:.1f}MB, {file_meta.get('modifiedTime', '')})"

    except Exception as e:
        return False, f"Drive 연결 실패: {e}"
