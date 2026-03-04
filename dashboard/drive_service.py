"""
Google Drive API service for Streamlit Cloud deployment.
Provides dual-mode operation: local (G:\ File Stream) vs cloud (Drive API).
"""
import os
import tempfile
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Singleton Drive service instance
_drive_service = None


def is_cloud() -> bool:
    """Detect if running in cloud environment (no local Google Drive mount)."""
    # Windows with G:\ mount = local mode
    if os.name == "nt":
        g_drive = Path(r"G:\내 드라이브")
        if g_drive.exists():
            return False
    # Everything else = cloud mode
    return True


def _get_secrets() -> dict:
    """Load secrets from Streamlit st.secrets (works both local .streamlit/secrets.toml and Cloud)."""
    try:
        import streamlit as st
        return dict(st.secrets)
    except Exception:
        return {}


def get_drive_service():
    """Get singleton Google Drive API service client using service account credentials."""
    global _drive_service
    if _drive_service is not None:
        return _drive_service

    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        secrets = _get_secrets()
        sa_info = secrets.get("gcp_service_account", {})

        if not sa_info:
            raise ValueError("gcp_service_account not found in Streamlit secrets")

        # Convert AttrDict to plain dict if needed
        if hasattr(sa_info, "to_dict"):
            sa_info = sa_info.to_dict()
        else:
            sa_info = dict(sa_info)

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


def get_drive_file_ids() -> dict:
    """Get configured Drive file/folder IDs from Streamlit secrets."""
    secrets = _get_secrets()
    drive_conf = secrets.get("drive", {})
    if hasattr(drive_conf, "to_dict"):
        drive_conf = drive_conf.to_dict()
    else:
        drive_conf = dict(drive_conf)
    return drive_conf


def download_master_file() -> str:
    """Download master file from Drive to a temp path. Returns the temp file path."""
    service = get_drive_service()
    file_ids = get_drive_file_ids()
    file_id = file_ids.get("master_file_id")

    if not file_id:
        raise ValueError("master_file_id not configured in secrets [drive] section")

    from googleapiclient.http import MediaIoBaseDownload
    import io

    request = service.files().get_media(fileId=file_id)
    tmp_dir = tempfile.gettempdir()
    tmp_path = os.path.join(tmp_dir, "도크발주관리데이터.xlsx")

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
    file_ids = get_drive_file_ids()
    file_id = file_ids.get("master_file_id")

    if not file_id:
        raise ValueError("master_file_id not configured in secrets [drive] section")

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
    """Upload a file to a specific Drive folder. Returns the new file ID."""
    service = get_drive_service()

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


def check_drive_connection() -> tuple:
    """Test Drive API connection. Returns (ok: bool, message: str)."""
    if not is_cloud():
        return True, "로컬 모드 (Google Drive File Stream)"

    try:
        file_ids = get_drive_file_ids()
        if not file_ids.get("master_file_id"):
            return False, "master_file_id가 secrets에 설정되지 않았습니다."

        service = get_drive_service()
        # Try to get master file metadata
        file_meta = service.files().get(
            fileId=file_ids["master_file_id"],
            fields="id, name, modifiedTime, size",
        ).execute()

        size_mb = int(file_meta.get("size", 0)) / (1024 * 1024)
        return True, f"연결됨: {file_meta['name']} ({size_mb:.1f}MB, {file_meta.get('modifiedTime', '')})"

    except Exception as e:
        return False, f"Drive 연결 실패: {e}"
