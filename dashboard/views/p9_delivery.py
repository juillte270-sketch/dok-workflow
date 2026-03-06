"""
p9_delivery.py — 배송 관리 대시보드 페이지

탭:
  - 현황: Firebase에서 배송 상태 실시간 조회 (로컬 JSON 폴백)
  - 기사 배정: 기사 선택 → 매장 배정 (Firebase 배치 쓰기)
  - 실시간 추적: folium 지도 + 기사 위치 (RTDB)
  - 업로드: processed_orders → Firebase 업로드
"""

import streamlit as st
import json
import subprocess
import sys
from pathlib import Path
from datetime import datetime, timezone, timedelta

# Paths
DASHBOARD_DIR = Path(__file__).resolve().parent.parent
PROJECT_ROOT = DASHBOARD_DIR.parent
EXECUTION_DIR = PROJECT_ROOT / "execution"
OUTPUTS_DIR = PROJECT_ROOT / "data" / "outputs"
INPUTS_DIR = PROJECT_ROOT / "data" / "inputs"

KST = timezone(timedelta(hours=9))

# Status display
STATUS_LABEL = {
    "pending": "⏳ 대기",
    "in_transit": "🚛 배송출발",
    "arrived": "📦 입고중",
    "delivered": "✅ 배송완료",
    "issue": "⚠️ 문제",
}

STATUS_COLOR_MAP = {
    "pending": "gray",
    "in_transit": "blue",
    "arrived": "orange",
    "delivered": "green",
    "issue": "red",
}

DRIVER_COLORS = ["#E74C3C", "#3498DB", "#2ECC71", "#9B59B6", "#F39C12", "#1ABC9C"]


# ─── Firebase Service (lazy import) ──────────────────────────

def _get_firebase():
    """Firebase service 모듈 lazy import. 실패 시 None."""
    try:
        from dashboard import firebase_service as fb
        return fb
    except Exception as e:
        st.warning(f"Firebase 모듈 로드 실패: {e}")
        return None


def _load_deliveries_firebase(date_compact: str) -> list[dict] | None:
    """Firebase에서 배송 데이터 로드. 실패 시 None."""
    fb = _get_firebase()
    if fb is None:
        return None
    try:
        deliveries = fb.get_deliveries(date_compact)
        return deliveries if deliveries else None
    except Exception as e:
        st.warning(f"Firebase 데이터 로드 실패: {e}")
        return None


# ─── Local JSON fallback ─────────────────────────────────────

def _load_delivery_json(date_compact: str) -> dict | None:
    """deliveries_YYYYMMDD.json 로드 (Firebase 없이 로컬 확인용)"""
    json_path = OUTPUTS_DIR / f"deliveries_{date_compact}.json"
    if json_path.exists():
        with open(json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def _run_upload(date_compact: str, dry_run: bool = False, json_only: bool = False) -> tuple[bool, str]:
    """upload_deliveries.py 실행"""
    cmd = [sys.executable, str(EXECUTION_DIR / "upload_deliveries.py"), "--date", date_compact]
    if dry_run:
        cmd.append("--dry-run")
    if json_only:
        cmd.append("--json-only")

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=60,
            cwd=str(PROJECT_ROOT),
        )
        output = result.stdout + (result.stderr or "")
        return result.returncode == 0, output
    except Exception as e:
        return False, str(e)


# ─── Tab 1: 현황 ─────────────────────────────────────────────

def _render_status_tab(date_compact: str):
    """현황 탭: Firebase 우선, 로컬 JSON 폴백."""

    # 새로고침 버튼
    if st.button("🔄 새로고침", key="refresh_status"):
        st.rerun()

    # Firebase에서 로드 시도
    deliveries = _load_deliveries_firebase(date_compact)
    source = "Firebase"

    if deliveries is None:
        # 로컬 JSON 폴백
        data = _load_delivery_json(date_compact)
        if data is None:
            st.info(f"{date_compact} 배송 데이터가 없습니다. '업로드' 탭에서 생성하세요.")
            return
        deliveries = data.get("deliveries", [])
        source = "로컬 JSON"

    st.caption(f"데이터 소스: {source}")

    # 통계
    fb = _get_firebase()
    if fb:
        stats = fb.compute_delivery_stats(deliveries)
    else:
        stats = {
            "total": len(deliveries),
            "pending": sum(1 for d in deliveries if d.get("status") == "pending"),
            "in_transit": sum(1 for d in deliveries if d.get("status") == "in_transit"),
            "arrived": sum(1 for d in deliveries if d.get("status") == "arrived"),
            "delivered": sum(1 for d in deliveries if d.get("status") == "delivered"),
            "issue": sum(1 for d in deliveries if d.get("status") == "issue"),
        }

    # 6개 메트릭 카드
    cols = st.columns(6)
    cols[0].metric("전체", stats["total"])
    cols[1].metric("대기", stats["pending"])
    cols[2].metric("이동중", stats["in_transit"])
    cols[3].metric("입고중", stats["arrived"])
    cols[4].metric("완료", stats["delivered"])
    cols[5].metric("문제", stats["issue"])

    # 진행률
    total = stats["total"]
    delivered = stats["delivered"]
    progress = delivered / total if total > 0 else 0
    st.progress(progress, text=f"진행률: {delivered}/{total} ({int(progress * 100)}%)")

    # 매장별 상태 테이블
    st.subheader("매장별 상태")

    table_data = []
    for d in deliveries:
        items = d.get("items", [])
        item_count = len(items)
        red_count = sum(1 for i in items if i.get("color") == "red")
        blue_count = sum(1 for i in items if i.get("color") == "blue")

        color_info = ""
        if red_count or blue_count:
            parts = []
            if red_count:
                parts.append(f"재고{red_count}")
            if blue_count:
                parts.append(f"시장{blue_count}")
            color_info = f" ({', '.join(parts)})"

        table_data.append({
            "순서": d.get("order", 0),
            "매장": d.get("storeName", ""),
            "품목 수": f"{item_count}{color_info}",
            "기사": d.get("driverName") or "-",
            "상태": STATUS_LABEL.get(d.get("status", ""), d.get("status", "")),
        })

    st.dataframe(table_data, use_container_width=True, hide_index=True)

    # 매장 상세 expander
    st.subheader("매장 상세")
    for d in deliveries:
        items = d.get("items", [])
        with st.expander(f"#{d.get('order', 0)} {d.get('storeName', '')} ({len(items)}개 품목)"):
            for item in items:
                color_tag = ""
                if item.get("color") == "red":
                    color_tag = " 🔴"
                elif item.get("color") == "blue":
                    color_tag = " 🔵"
                st.text(f"  {item.get('name', '')}  {item.get('qty', '')}{color_tag}")


# ─── Tab 2: 기사 배정 ────────────────────────────────────────

def _render_assign_tab(date_compact: str):
    """기사 배정 탭: Firebase에서 기사/배송 조회 → 배정."""

    fb = _get_firebase()
    if fb is None:
        st.warning("Firebase 연결이 필요합니다. serviceAccountKey.json을 확인하세요.")
        return

    # 배송 데이터 로드
    try:
        deliveries = fb.get_deliveries(date_compact)
    except Exception as e:
        st.error(f"배송 데이터 로드 실패: {e}")
        return

    if not deliveries:
        st.info(f"{date_compact} 배송 데이터가 없습니다. 먼저 업로드하세요.")
        return

    # 기사 목록
    try:
        drivers = fb.get_active_drivers()
    except Exception as e:
        st.error(f"기사 목록 로드 실패: {e}")
        return

    if not drivers:
        st.warning("등록된 기사가 없습니다. 모바일 앱에서 기사를 등록하세요.")
        return

    # 기사 선택
    st.subheader("기사 선택")
    driver_options = {d.get("name", d["uid"]): d for d in drivers}
    selected_name = st.radio(
        "배정할 기사",
        options=list(driver_options.keys()),
        horizontal=True,
    )
    selected_driver = driver_options[selected_name]

    # 미배정 / 배정완료 분류
    unassigned = [d for d in deliveries if not d.get("assignedTo")]
    assigned = [d for d in deliveries if d.get("assignedTo")]

    # 미배정 매장 선택
    st.subheader(f"미배정 매장 ({len(unassigned)}개)")
    if unassigned:
        store_options = {
            f"#{d['order']} {d['storeName']}": d["id"]
            for d in unassigned
        }
        selected_stores = st.multiselect(
            "배정할 매장 선택",
            options=list(store_options.keys()),
        )

        if selected_stores:
            store_ids = [store_options[s] for s in selected_stores]
            btn_label = f"{len(selected_stores)}개 매장 → {selected_name} 배정"
            if st.button(btn_label, type="primary"):
                with st.spinner("배정 중..."):
                    try:
                        fb.assign_driver(
                            date_compact,
                            selected_driver["uid"],
                            selected_name,
                            store_ids,
                        )
                        st.success(f"{selected_name}에게 {len(store_ids)}개 매장 배정 완료!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"배정 실패: {e}")
    else:
        st.success("모든 매장이 배정되었습니다.")

    # 배정 완료 현황
    if assigned:
        st.subheader("배정 현황")
        # 기사별 그룹핑
        by_driver: dict[str, list] = {}
        for d in assigned:
            dname = d.get("driverName", "미지정")
            by_driver.setdefault(dname, []).append(d)

        for dname, stores in by_driver.items():
            with st.expander(f"🚛 {dname} ({len(stores)}개 매장)"):
                for d in stores:
                    status = STATUS_LABEL.get(d.get("status", ""), d.get("status", ""))
                    st.text(f"  #{d.get('order', 0)} {d.get('storeName', '')} — {status}")


# ─── Tab 3: 실시간 추적 ──────────────────────────────────────

def _render_tracking_tab(date_compact: str):
    """실시간 추적 탭: folium 지도 + 기사 위치."""

    fb = _get_firebase()
    if fb is None:
        st.warning("Firebase 연결이 필요합니다.")
        return

    # folium 체크
    try:
        import folium
        from streamlit_folium import st_folium
    except ImportError:
        st.error("folium / streamlit-folium 패키지가 필요합니다.")
        st.code("pip install folium streamlit-folium")
        return

    if st.button("🔄 새로고침", key="refresh_tracking"):
        st.rerun()

    # 배송 데이터
    try:
        deliveries = fb.get_deliveries(date_compact)
    except Exception as e:
        st.error(f"배송 데이터 로드 실패: {e}")
        return

    if not deliveries:
        st.info(f"{date_compact} 배송 데이터가 없습니다.")
        return

    # 기사 위치 조회
    assigned_uids = list({d["assignedTo"] for d in deliveries if d.get("assignedTo")})
    driver_locations = {}
    if assigned_uids:
        try:
            driver_locations = fb.get_all_driver_locations(assigned_uids)
        except Exception:
            pass

    # 요약 메트릭
    stats = fb.compute_delivery_stats(deliveries)
    assigned_count = sum(1 for d in deliveries if d.get("assignedTo"))
    unassigned_count = stats["total"] - assigned_count

    cols = st.columns(4)
    cols[0].metric("전체", stats["total"])
    cols[1].metric("배정", assigned_count)
    cols[2].metric("완료", stats["delivered"])
    cols[3].metric("미배정", unassigned_count)

    # 지도 생성
    m = folium.Map(
        location=[37.50, 127.05],
        zoom_start=11,
        tiles="CartoDB positron",
    )

    # 매장 마커
    for d in deliveries:
        coords = fb.get_store_coords(d.get("storeName", ""))
        if coords is None:
            continue

        lat, lng = coords
        status = d.get("status", "pending")
        color = STATUS_COLOR_MAP.get(status, "gray")
        driver_info = f"<br>기사: {d['driverName']}" if d.get("driverName") else ""
        items_count = len(d.get("items", []))

        popup_html = (
            f"<b>#{d.get('order', 0)} {d.get('storeName', '')}</b><br>"
            f"상태: {STATUS_LABEL.get(status, status)}<br>"
            f"품목: {items_count}개"
            f"{driver_info}"
        )

        folium.CircleMarker(
            location=[lat, lng],
            radius=10,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.7,
            popup=folium.Popup(popup_html, max_width=200),
            tooltip=d.get("storeName", ""),
        ).add_to(m)

    # 기사 마커
    now_kst = datetime.now(KST)
    driver_names_map = {}

    # 배송 데이터에서 uid→name 매핑
    for d in deliveries:
        if d.get("assignedTo") and d.get("driverName"):
            driver_names_map[d["assignedTo"]] = d["driverName"]

    for idx, (uid, loc) in enumerate(driver_locations.items()):
        lat = loc.get("latitude")
        lng = loc.get("longitude")
        if not lat or not lng:
            continue

        name = driver_names_map.get(uid, uid[:4])
        initial = name[0] if name else "?"
        color = DRIVER_COLORS[idx % len(DRIVER_COLORS)]

        # 온라인 판정: updatedAt이 10분 이내
        online = False
        updated_at = loc.get("updatedAt")
        if updated_at:
            try:
                if isinstance(updated_at, (int, float)):
                    updated_dt = datetime.fromtimestamp(updated_at / 1000, tz=KST)
                else:
                    updated_dt = datetime.fromisoformat(str(updated_at))
                    if updated_dt.tzinfo is None:
                        updated_dt = updated_dt.replace(tzinfo=KST)
                online = (now_kst - updated_dt).total_seconds() < 600
            except Exception:
                pass

        border_color = color if online else "#999"

        icon_html = (
            f'<div style="'
            f"background:{color};"
            f"color:white;"
            f"border:3px solid {border_color};"
            f"border-radius:50%;"
            f"width:32px;height:32px;"
            f"display:flex;align-items:center;justify-content:center;"
            f"font-weight:bold;font-size:14px;"
            f'">{initial}</div>'
        )

        folium.Marker(
            location=[lat, lng],
            icon=folium.DivIcon(
                html=icon_html,
                icon_size=(32, 32),
                icon_anchor=(16, 16),
            ),
            tooltip=f"{name} ({'온라인' if online else '오프라인'})",
        ).add_to(m)

    # 지도 렌더링
    st_folium(m, width=700, height=500, returned_objects=[])

    # 기사별 상세 카드
    if driver_locations:
        st.subheader("기사 상세")
        for idx, (uid, loc) in enumerate(driver_locations.items()):
            name = driver_names_map.get(uid, uid[:4])
            color = DRIVER_COLORS[idx % len(DRIVER_COLORS)]

            # 온라인 상태
            online = False
            updated_at = loc.get("updatedAt")
            time_str = "-"
            if updated_at:
                try:
                    if isinstance(updated_at, (int, float)):
                        updated_dt = datetime.fromtimestamp(updated_at / 1000, tz=KST)
                    else:
                        updated_dt = datetime.fromisoformat(str(updated_at))
                        if updated_dt.tzinfo is None:
                            updated_dt = updated_dt.replace(tzinfo=KST)
                    online = (now_kst - updated_dt).total_seconds() < 600
                    time_str = updated_dt.strftime("%H:%M:%S")
                except Exception:
                    pass

            status_dot = f"🟢 온라인" if online else "⚪ 오프라인"
            speed = loc.get("speed", 0) or 0

            # 이 기사에게 배정된 매장
            my_deliveries = [d for d in deliveries if d.get("assignedTo") == uid]
            my_done = sum(1 for d in my_deliveries if d.get("status") == "delivered")
            my_total = len(my_deliveries)
            my_progress = my_done / my_total if my_total > 0 else 0

            with st.expander(f"🚛 {name} — {status_dot} | {my_done}/{my_total} 완료"):
                st.markdown(f"**마지막 업데이트:** {time_str}")
                st.markdown(f"**속도:** {speed:.0f} km/h")
                st.progress(my_progress, text=f"진행률: {int(my_progress * 100)}%")

                # 배송 타임라인
                if my_deliveries:
                    for d in my_deliveries:
                        s = d.get("status", "pending")
                        label = STATUS_LABEL.get(s, s)
                        st.text(f"  #{d.get('order', 0)} {d.get('storeName', '')} — {label}")


# ─── Tab 4: 업로드 (기존 유지) ───────────────────────────────

def _render_upload_tab(date_compact: str):
    """업로드 탭: processed_orders → Firebase."""

    st.subheader("배송 데이터 업로드")

    processed_path = INPUTS_DIR / f"processed_orders_{date_compact}.txt"
    file_exists = processed_path.exists()

    if file_exists:
        st.success(f"발주 파일 발견: `processed_orders_{date_compact}.txt`")
    else:
        st.warning(f"발주 파일이 없습니다: `processed_orders_{date_compact}.txt`")
        st.caption("먼저 '2. 발주 처리' 페이지에서 발주를 처리하세요.")

    col1, col2, col3 = st.columns(3)

    # JSON 생성 (로컬)
    with col1:
        if st.button("JSON 미리보기", disabled=not file_exists, use_container_width=True):
            with st.spinner("배송 데이터 생성 중..."):
                ok, output = _run_upload(date_compact, json_only=True)
            if ok:
                st.success("JSON 생성 완료!")
                data = _load_delivery_json(date_compact)
                if data:
                    st.json(data)
            else:
                st.error(f"오류:\n{output}")

    # Dry-run
    with col2:
        if st.button("Dry-Run", disabled=not file_exists, use_container_width=True):
            with st.spinner("미리보기 중..."):
                ok, output = _run_upload(date_compact, dry_run=True)
            st.code(output)

    # Firebase 업로드
    with col3:
        if st.button("Firebase 업로드", disabled=not file_exists, type="primary", use_container_width=True):
            with st.spinner("Firebase에 업로드 중..."):
                ok, output = _run_upload(date_compact)
            if ok:
                st.success("Firebase 업로드 완료!")
            else:
                st.error(f"업로드 실패:\n{output}")
            st.code(output)


# ─── Main render ─────────────────────────────────────────────

def render(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    st.header("배송 관리")

    tab1, tab2, tab3, tab4 = st.tabs(["현황", "기사 배정", "실시간 추적", "업로드"])

    with tab1:
        _render_status_tab(date_compact)

    with tab2:
        _render_assign_tab(date_compact)

    with tab3:
        _render_tracking_tab(date_compact)

    with tab4:
        _render_upload_tab(date_compact)
