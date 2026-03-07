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

# 매장 좌표 + 주소 (mapService.ts와 동기화)
STORE_COORDS: list[tuple[str, float, float, str]] = [
    ("육회", 37.487570, 127.132157, "서울 송파구 송이로 106 2층"),
    ("봄날", 37.492947, 127.124287, "서울 송파구 중대로12길 18-21 2층"),
    ("고른햇살", 37.590068, 127.030305, "서울 성북구 개운사길 14"),
    ("소유", 37.538620, 127.056875, "서울 성동구 뚝섬로9길 16"),
    ("선데이", 37.526549, 127.036813, "서울 강남구 언주로170길 37 201호"),
    ("이너프유", 37.502248, 126.789145, "경기 부천시 부천로198번길 27 서림테크노파크1차 404호"),
    ("브럭시", 37.513742, 127.104247, "서울 송파구 올림픽로 300 롯데월드몰 B3층"),
    ("압구정", 37.527328, 127.037997, "서울 강남구 압구정로50길 8 1층"),
    ("하남", 37.545445, 127.224057, "경기 하남시 미사대로 750 스타필드하남"),
    ("강남점", 37.500078, 127.027516, "서울 강남구 테헤란로1길 28-11 1층"),
    ("마곡", 37.559229, 126.835888, "서울 강서구 공항대로 247"),
    ("구월", 37.445662, 126.703161, "인천 남동구 인하로507번길 23-1 1층 101호"),
    ("주안", 37.461984, 126.671774, "인천 미추홀구 경인로 372"),
    ("분당", 37.349277, 127.108402, "경기 성남시 분당구 성남대로 151 분당엠코헤리츠 2층"),
    ("평택", 37.033484, 127.013226, "경기 평택시 고덕면 방축3길 57-8 2층"),
    ("동탄", 37.200359, 127.095576, "경기 화성시 동탄대로5길 21 B1층"),
    ("도봉", 37.648923, 127.035000, "서울 도봉구 도봉로 621 롯데캐슬 102동 2층"),
    ("광교", 37.287526, 127.055817, "경기 수원시 영통구 도청로18번길 26 지하1층"),
    ("역삼", 37.499555, 127.047845, "서울 강남구 역삼로 161 지하1층"),
    ("송파", 37.493820, 127.122433, "서울 송파구 중대로10길 17 1층"),
    ("동원", 37.503860, 127.113830, "서울 송파구 가락로 2"),
]

DEPOT_COORD = (37.503860, 127.113830)  # 가락시장

# 매장 출입방법 (mapService.ts와 동기화)
STORE_ENTRY_MAP: dict[str, str] = {
    "육회": "5211#",
    "봄날": "1190702*",
    "고른햇살": "1004*",
    "소유": "102030*",
    "선데이": "자율출입",
    "이너프유": "보안키",
    "브럭시": "자율출입 (롯데월드몰 B3층 R3기둥)",
    "압구정": "자율출입",
    "하남": "자율출입 (스타필드하남)",
    "강남점": "보안키",
    "마곡": "냉장고872",
    "구월점": "보안키",
    "주안": "보안키",
    "분당": "입구배송",
    "평택": "보안키",
    "동탄": "자율출입",
    "도봉": "",
    "광교": "예정",
    "역삼": "160309*",
    "송파": "1234*",
}


def get_store_coords(store_name: str) -> tuple[float, float] | None:
    """매장명 → (lat, lng). 부분 문자열 매칭."""
    name_norm = store_name.replace(" ", "")
    for keyword, lat, lng, _addr in STORE_COORDS:
        if keyword in name_norm:
            return (lat, lng)
    return None


def get_store_address(store_name: str) -> str | None:
    """매장명 → 주소."""
    name_norm = store_name.replace(" ", "")
    for keyword, _lat, _lng, addr in STORE_COORDS:
        if keyword in name_norm:
            return addr
    return None


def get_store_entry(store_name: str) -> str | None:
    """매장명 → 출입방법."""
    name_norm = store_name.replace(" ", "")
    for keyword, entry in STORE_ENTRY_MAP.items():
        if keyword in name_norm:
            return entry if entry else None
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
