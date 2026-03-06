"""
Firebase Admin SDK service layer for Streamlit dashboard.
Singleton initialization + Firestore/RTDB helpers.

Supports:
  - Local: serviceAccountKey.json (project root)
  - Cloud: st.secrets["firebase"] section
  - Coexists with upload_deliveries.py (firebase_admin._apps check)
"""

import logging
from pathlib import Path
from datetime import datetime, timezone, timedelta

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RTDB_URL = "https://dok-delivery-default-rtdb.firebaseio.com"
KST = timezone(timedelta(hours=9))

# ─── Singleton ────────────────────────────────────────────────

_firestore_client = None
_initialized = False


def _ensure_initialized():
    """Firebase Admin SDK 초기화 (한 번만)."""
    global _initialized
    if _initialized:
        return

    import firebase_admin
    from firebase_admin import credentials

    if firebase_admin._apps:
        _initialized = True
        return

    # 1) Try Streamlit secrets
    try:
        import streamlit as st
        if "firebase" in st.secrets:
            sa_info = dict(st.secrets["firebase"])
            cred = credentials.Certificate(sa_info)
            firebase_admin.initialize_app(cred, {"databaseURL": RTDB_URL})
            _initialized = True
            logger.info("Firebase initialized from st.secrets")
            return
    except Exception:
        pass

    # 2) Try local service account key
    key_path = PROJECT_ROOT / "serviceAccountKey.json"
    if key_path.exists():
        cred = credentials.Certificate(str(key_path))
        firebase_admin.initialize_app(cred, {"databaseURL": RTDB_URL})
        _initialized = True
        logger.info("Firebase initialized from serviceAccountKey.json")
        return

    raise FileNotFoundError(
        "Firebase 인증 정보를 찾을 수 없습니다. "
        "serviceAccountKey.json 또는 st.secrets['firebase'] 필요"
    )


def get_firestore_client():
    """싱글톤 Firestore 클라이언트."""
    global _firestore_client
    if _firestore_client is not None:
        return _firestore_client

    _ensure_initialized()
    from firebase_admin import firestore
    _firestore_client = firestore.client()
    return _firestore_client


def get_rtdb():
    """Realtime Database 모듈 반환."""
    _ensure_initialized()
    from firebase_admin import db
    return db


# ─── Firestore Helpers ────────────────────────────────────────

def get_deliveries(date_compact: str) -> list[dict]:
    """deliveryDays/{date}/deliveries 전체 조회 (order 순)."""
    db = get_firestore_client()
    ref = db.collection("deliveryDays").document(date_compact).collection("deliveries")
    docs = ref.order_by("order").stream()

    results = []
    for doc in docs:
        d = doc.to_dict()
        d["id"] = doc.id
        results.append(d)
    return results


def get_day_summary(date_compact: str) -> dict | None:
    """deliveryDays/{date} 문서 조회."""
    db = get_firestore_client()
    doc = db.collection("deliveryDays").document(date_compact).get()
    return doc.to_dict() if doc.exists else None


def get_active_drivers() -> list[dict]:
    """users 컬렉션에서 role=driver인 유저 조회.

    Returns list of {uid, name, displayName, email, phone, active, ...}
    """
    from google.cloud.firestore_v1.base_query import FieldFilter

    db = get_firestore_client()
    docs = db.collection("users").where(filter=FieldFilter("role", "==", "driver")).stream()

    drivers = []
    for doc in docs:
        d = doc.to_dict()
        d["uid"] = doc.id
        # Normalize: ensure 'name' key exists (app uses displayName)
        d["name"] = d.get("displayName") or d.get("name") or doc.id[:6]
        drivers.append(d)
    return drivers


def assign_driver(date_compact: str, uid: str, name: str, store_ids: list[str]):
    """기사에게 매장 배정 (배치 쓰기).

    1) assignments/{date}/drivers/{uid} 에 배정 매장 기록
    2) 각 delivery 문서에 assignedTo/driverName 업데이트
    """
    db = get_firestore_client()
    batch = db.batch()

    # Assignment record
    assign_ref = (
        db.collection("assignments")
        .document(date_compact)
        .collection("drivers")
        .document(uid)
    )
    batch.set(assign_ref, {
        "driverName": name,
        "storeIds": store_ids,
        "assignedAt": datetime.now(KST).isoformat(),
    })

    # Update each delivery
    day_ref = db.collection("deliveryDays").document(date_compact)
    for sid in store_ids:
        doc_ref = day_ref.collection("deliveries").document(sid)
        batch.update(doc_ref, {
            "assignedTo": uid,
            "driverName": name,
        })

    batch.commit()
    logger.info(f"Assigned {name} ({uid}) → {len(store_ids)} stores on {date_compact}")


def get_assignments(date_compact: str) -> dict[str, dict]:
    """assignments/{date}/drivers 전체 조회. {uid: {driverName, storeIds, ...}}"""
    db = get_firestore_client()
    ref = (
        db.collection("assignments")
        .document(date_compact)
        .collection("drivers")
    )
    docs = ref.stream()
    return {doc.id: doc.to_dict() for doc in docs}


# ─── RTDB Helpers (실시간 추적) ────────────────────────────────

def get_all_driver_locations(uids: list[str]) -> dict[str, dict]:
    """RTDB에서 기사 GPS 위치 일괄 조회.

    Returns: {uid: {latitude, longitude, speed, heading, updatedAt, ...}}
    """
    rtdb = get_rtdb()
    locations = {}
    for uid in uids:
        ref = rtdb.reference(f"driverLocations/{uid}")
        data = ref.get()
        if data:
            locations[uid] = data
    return locations


# ─── 매장 좌표 (mapService.ts에서 포팅) ──────────────────────

STORE_COORDS: list[tuple[str, float, float]] = [
    ("봄날", 37.5260, 127.0390),
    ("송파", 37.4983, 127.0641),
    ("육회", 37.4992, 127.1055),
    ("도봉", 37.6693, 127.0472),
    ("고른햇살", 37.5270, 127.0435),
    ("소유", 37.5245, 127.0380),
    ("선데이", 37.5240, 127.0385),
    ("압구정", 37.5266, 127.0406),
    ("강남점", 37.5006, 127.0365),
    ("역삼", 37.5010, 127.0367),
    ("마곡", 37.5614, 126.8285),
    ("이너프유", 37.5270, 127.0440),
    ("주안", 37.4546, 126.7033),
    ("구월", 37.4568, 126.7215),
    ("광교", 37.2893, 127.0435),
    ("동원", 37.4929, 127.1162),
    ("분당", 37.3591, 127.1053),
    ("평택", 37.0062, 127.1278),
    ("동탄", 37.2001, 127.0702),
    ("브럭시", 37.5131, 127.1025),
    ("하남", 37.5455, 127.2043),
]

DEPOT_COORD = (37.4929, 127.1162)  # 가락동 동원1차


def get_store_coords(store_name: str) -> tuple[float, float] | None:
    """매장명 → (lat, lng). 부분 문자열 매칭."""
    name_norm = store_name.replace(" ", "")
    for keyword, lat, lng in STORE_COORDS:
        if keyword in name_norm:
            return (lat, lng)
    return None


# ─── 통계 ─────────────────────────────────────────────────────

def compute_delivery_stats(deliveries: list[dict]) -> dict:
    """상태별 카운트 계산."""
    stats = {
        "total": len(deliveries),
        "pending": 0,
        "in_transit": 0,
        "arrived": 0,
        "delivered": 0,
        "issue": 0,
    }
    for d in deliveries:
        s = d.get("status", "pending")
        if s in stats:
            stats[s] += 1
    return stats


# ─── 연결 테스트 ──────────────────────────────────────────────

def check_firebase_connection() -> tuple[bool, str]:
    """Firebase 연결 테스트."""
    try:
        _ensure_initialized()
        db = get_firestore_client()
        # Simple read test
        db.collection("deliveryDays").limit(1).get()
        return True, "Firebase 연결 성공"
    except Exception as e:
        return False, f"Firebase 연결 실패: {e}"
