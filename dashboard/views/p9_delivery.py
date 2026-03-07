"""
p9_delivery.py — 배송 관리 대시보드 페이지

탭:
  - 현황: Firebase에서 배송 상태 실시간 조회 (로컬 JSON 폴백)
  - 기사 배정: 기사 선택 → 매장 배정 (Firebase 배치 쓰기)
  - 실시간 추적: folium 지도 + 기사 위치 (RTDB)
  - 리포트: 기사별 배송 성과 분석 (소요시간, 완료율, 타임라인)
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


def _fmt_timestamp(ts) -> str:
    """Firebase Timestamp → HH:MM 문자열."""
    try:
        if hasattr(ts, "timestamp"):
            # Firestore Timestamp
            dt = datetime.fromtimestamp(ts.timestamp(), tz=KST)
        elif isinstance(ts, (int, float)):
            dt = datetime.fromtimestamp(ts / 1000, tz=KST)
        elif isinstance(ts, dict) and "_seconds" in ts:
            dt = datetime.fromtimestamp(ts["_seconds"], tz=KST)
        elif isinstance(ts, str):
            dt = datetime.fromisoformat(ts)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=KST)
        else:
            return str(ts)
        return dt.strftime("%H:%M")
    except Exception:
        return str(ts)

# Status display
STATUS_LABEL = {
    "pending": "🚛 배송출발",
    "in_transit": "🚛 배송출발",
    "arrived": "📦 입고중",
    "delivered": "✅ 배송완료",
    "issue": "⚠️ 문제",
}

STATUS_COLOR_MAP = {
    "pending": "#3B82F6",     # 파랑 — 배송출발
    "in_transit": "#3B82F6",  # 파랑 — 배송출발
    "arrived": "#F97316",     # 주황 — 입고중
    "delivered": "#10B981",   # 초록 — 배송완료
    "issue": "#DC2626",       # 진빨강 — 문제
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

    # 메트릭 + 진행률 (모바일 대응 HTML)
    total = stats["total"]
    delivered = stats["delivered"]
    pct = int(delivered / total * 100) if total > 0 else 0
    bar_color = "#10B981" if pct == 100 else "#3B82F6"
    issue_html = (
        f'<span style="background:#7F1D1D;color:#FCA5A5;padding:2px 8px;border-radius:12px">'
        f'문제 <b>{stats["issue"]}</b></span>'
        if stats["issue"] else ""
    )

    st.markdown(
        f'<div style="display:flex;flex-wrap:wrap;gap:6px;margin:8px 0">'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem">'
        f'전체 <b>{total}</b></span>'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem;color:#3B82F6">'
        f'배송출발 <b>{stats["pending"] + stats["in_transit"]}</b></span>'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem;color:#F97316">'
        f'입고중 <b>{stats["arrived"]}</b></span>'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem;color:#10B981">'
        f'배송완료 <b>{delivered}</b></span>'
        f'{issue_html}'
        f'</div>'
        f'<div style="background:#1E293B;border-radius:4px;height:6px;margin:4px 0;overflow:hidden">'
        f'<div style="background:{bar_color};width:{pct}%;height:100%;border-radius:4px"></div></div>'
        f'<div style="text-align:right;font-size:0.72rem;color:#64748B">{delivered}/{total} ({pct}%)</div>',
        unsafe_allow_html=True,
    )

    # 매장별 상태 리스트 (모바일 대응)
    rows_html = ""
    for d in deliveries:
        items = d.get("items", [])
        item_count = len(items)
        red_count = sum(1 for i in items if i.get("color") == "red")
        blue_count = sum(1 for i in items if i.get("color") == "blue")

        color_tags = ""
        if red_count:
            color_tags += f'<span style="color:#F87171;font-size:0.7rem">재고{red_count}</span> '
        if blue_count:
            color_tags += f'<span style="color:#60A5FA;font-size:0.7rem">시장{blue_count}</span>'

        status = d.get("status", "pending")
        s_color = {"pending": "#3B82F6", "in_transit": "#3B82F6", "arrived": "#F97316",
                    "delivered": "#10B981", "issue": "#DC2626"}.get(status, "#94A3B8")
        s_label = {"pending": "배송출발", "in_transit": "배송출발", "arrived": "입고중",
                    "delivered": "배송완료", "issue": "문제"}.get(status, status)
        driver = d.get("driverName") or ""
        driver_html = f'<span style="color:#94A3B8;font-size:0.75rem">{driver}</span>' if driver else ""

        rows_html += (
            f'<div style="display:flex;align-items:center;gap:8px;padding:7px 10px;'
            f'border-bottom:1px solid #1E293B">'
            f'<span style="color:#64748B;font-size:0.75rem;min-width:18px">{d.get("order", 0)}</span>'
            f'<span style="flex:1;font-size:0.85rem;line-height:1.3">'
            f'<b>{d.get("storeName", "")}</b> '
            f'<span style="color:#64748B;font-size:0.75rem">{item_count}개</span> '
            f'{color_tags}</span>'
            f'<span style="display:flex;flex-direction:column;align-items:flex-end;gap:1px">'
            f'<span style="color:{s_color};font-size:0.75rem;font-weight:600">{s_label}</span>'
            f'{driver_html}</span>'
            f'</div>'
        )

    st.markdown(
        f'<div style="border:1px solid #334155;border-radius:10px;overflow:hidden;margin:8px 0">'
        f'{rows_html}</div>',
        unsafe_allow_html=True,
    )

    # 매장 상세 expander
    for d in deliveries:
        items = d.get("items", [])
        store_name = d.get("storeName", "")
        with st.expander(f"#{d.get('order', 0)} {store_name} ({len(items)}개)"):
            # 전체를 하나의 HTML 블록으로 렌더링 (Streamlit 마진 겹침 방지)
            html = '<div style="font-size:0.82rem;line-height:1.7">'

            # 주소 + 출입정보
            addr = fb.get_store_address(store_name) if fb else None
            entry = fb.get_store_entry(store_name) if fb else None
            if addr:
                html += f'<div style="color:#94A3B8;font-size:0.78rem;word-break:break-all">📍 {addr}</div>'
            if entry:
                html += f'<div style="color:#94A3B8;font-size:0.78rem">🔑 {entry}</div>'

            # 타임라인
            started = d.get("startedAt")
            arrived_at = d.get("arrivedAt")
            completed_at = d.get("completedAt")
            if started or arrived_at or completed_at:
                html += '<div style="margin:6px 0;font-size:0.78rem;color:#94A3B8">'
                if started:
                    html += f'<div>🚛 출발 <b style="color:#CBD5E1">{_fmt_timestamp(started)}</b></div>'
                if arrived_at:
                    html += f'<div>📦 도착 <b style="color:#CBD5E1">{_fmt_timestamp(arrived_at)}</b></div>'
                if completed_at:
                    html += f'<div>✅ 완료 <b style="color:#CBD5E1">{_fmt_timestamp(completed_at)}</b></div>'
                html += '</div>'

            # 품목 목록
            if items:
                html += '<div style="margin-top:6px;border-top:1px solid #334155;padding-top:4px">'
                for item in items:
                    color_dot = ""
                    if item.get("color") == "red":
                        color_dot = ' <span style="color:#EF4444">●</span>'
                    elif item.get("color") == "blue":
                        color_dot = ' <span style="color:#3B82F6">●</span>'
                    html += (
                        f'<div>{item.get("name", "")} '
                        f'<span style="color:#64748B">{item.get("qty", "")}</span>'
                        f'{color_dot}</div>'
                    )
                html += '</div>'

            # 메모
            notes = d.get("notes")
            if notes:
                html += f'<div style="color:#FBBF24;margin-top:6px;font-size:0.78rem">📝 {notes}</div>'

            html += '</div>'
            st.markdown(html, unsafe_allow_html=True)

            # 배송 사진 (이미지는 st.image 필요)
            photo_url = d.get("photoUrl")
            if photo_url:
                st.image(photo_url, caption="배송 사진", width=250)


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
        key=f"assign_driver_{date_compact}",
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

    # 배정 완료 현황 + 재배정
    if assigned:
        st.subheader("배정 현황")
        by_driver: dict[str, list] = {}
        for d in assigned:
            dname = d.get("driverName", "미지정")
            by_driver.setdefault(dname, []).append(d)

        other_drivers = [n for n in driver_options if n != selected_name]

        for dname, stores in by_driver.items():
            with st.expander(f"🚛 {dname} ({len(stores)}개 매장)"):
                for d in stores:
                    status = STATUS_LABEL.get(d.get("status", ""), d.get("status", ""))
                    st.text(f"  #{d.get('order', 0)} {d.get('storeName', '')} — {status}")

                # 재배정 UI
                if other_drivers:
                    st.divider()
                    reassign_options = {
                        f"#{d['order']} {d['storeName']}": d["id"]
                        for d in stores
                    }
                    reassign_key = f"reassign_{dname}"
                    picked = st.multiselect(
                        "재배정할 매장",
                        options=list(reassign_options.keys()),
                        key=reassign_key,
                    )
                    target_key = f"target_{dname}"
                    target_name = st.selectbox(
                        "이동할 기사",
                        options=other_drivers,
                        key=target_key,
                    )
                    if picked and target_name:
                        ids = [reassign_options[p] for p in picked]
                        target_driver = driver_options[target_name]
                        btn_key = f"btn_reassign_{dname}"
                        if st.button(
                            f"{len(picked)}개 → {target_name} 재배정",
                            key=btn_key,
                        ):
                            with st.spinner("재배정 중..."):
                                try:
                                    fb.assign_driver(
                                        date_compact,
                                        target_driver["uid"],
                                        target_name,
                                        ids,
                                    )
                                    st.success(f"{target_name}에게 {len(ids)}개 매장 재배정 완료!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"재배정 실패: {e}")


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

    # 요약 메트릭 (모바일 대응)
    stats = fb.compute_delivery_stats(deliveries)
    assigned_count = sum(1 for d in deliveries if d.get("assignedTo"))
    unassigned_count = stats["total"] - assigned_count

    st.markdown(
        f'<div style="display:flex;flex-wrap:wrap;gap:6px;margin:8px 0">'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem">'
        f'전체 <b>{stats["total"]}</b></span>'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem;color:#60A5FA">'
        f'배정 <b>{assigned_count}</b></span>'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem;color:#4ADE80">'
        f'완료 <b>{stats["delivered"]}</b></span>'
        f'<span style="background:#1E293B;padding:4px 10px;border-radius:8px;font-size:0.85rem;color:#94A3B8">'
        f'미배정 <b>{unassigned_count}</b></span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # 지도 생성 (Vworld 한국어 타일)
    m = folium.Map(
        location=[37.50, 127.05],
        zoom_start=11,
        tiles=None,
    )
    folium.TileLayer(
        tiles="https://tile.openstreetmap.org/{z}/{x}/{y}.png",
        attr="OpenStreetMap",
        name="OpenStreetMap",
    ).add_to(m)
    folium.TileLayer(
        tiles="http://api.vworld.kr/req/wmts/1.0.0/EEAB40D2-8498-3E92-8E60-40B0B5010485/Base/{z}/{y}/{x}.png",
        attr="Vworld",
        name="Vworld 한국어",
    ).add_to(m)
    folium.LayerControl().add_to(m)

    # 범례
    legend_html = (
        '<div style="position:fixed;bottom:30px;left:10px;z-index:1000;'
        'background:rgba(0,0,0,0.75);padding:8px 12px;border-radius:8px;'
        'font-size:12px;color:white;line-height:1.6">'
        '<span style="color:#3B82F6">●</span> 배송출발 &nbsp;'
        '<span style="color:#F97316">●</span> 입고중 &nbsp;'
        '<span style="color:#10B981">●</span> 완료 &nbsp;'
        '<span style="color:#8B5CF6">●</span> 기사'
        '</div>'
    )
    m.get_root().html.add_child(folium.Element(legend_html))

    # 매장 마커
    for d in deliveries:
        coords = fb.get_store_coords(d.get("storeName", ""))
        if coords is None:
            continue

        lat, lng = coords
        status = d.get("status", "pending")
        color = STATUS_COLOR_MAP.get(status, "#EF4444")
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
            fill_opacity=0.8,
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
        color = "#8B5CF6"  # 기사 마커: 보라

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

    # 지도 렌더링 (모바일 대응: 폭 100%)
    st_folium(m, use_container_width=True, height=450, returned_objects=[])

    # 기사별 상세 카드
    if driver_locations:
        st.subheader("기사 상세")
        for idx, (uid, loc) in enumerate(driver_locations.items()):
            name = driver_names_map.get(uid, uid[:4])

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


# ─── Tab 4: 리포트 ────────────────────────────────────────────

def _ts_to_datetime(ts):
    """Firebase Timestamp → datetime (KST). 실패 시 None."""
    try:
        if hasattr(ts, "timestamp"):
            return datetime.fromtimestamp(ts.timestamp(), tz=KST)
        elif isinstance(ts, (int, float)):
            return datetime.fromtimestamp(ts / 1000, tz=KST)
        elif isinstance(ts, dict) and "_seconds" in ts:
            return datetime.fromtimestamp(ts["_seconds"], tz=KST)
        elif isinstance(ts, str):
            dt = datetime.fromisoformat(ts)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=KST)
            return dt
    except Exception:
        pass
    return None


def _render_report_tab(date_compact: str):
    """리포트 탭: 기사별 배송 성과 분석."""

    fb = _get_firebase()
    if fb is None:
        st.warning("Firebase 연결이 필요합니다.")
        return

    try:
        deliveries = fb.get_deliveries(date_compact)
    except Exception as e:
        st.error(f"배송 데이터 로드 실패: {e}")
        return

    if not deliveries:
        st.info(f"{date_compact} 배송 데이터가 없습니다.")
        return

    if st.button("🔄 새로고침", key="refresh_report"):
        st.rerun()

    # 전체 요약
    stats = fb.compute_delivery_stats(deliveries)
    total = stats["total"]
    delivered = stats["delivered"]
    pct = int(delivered / total * 100) if total > 0 else 0

    # 전체 소요시간 계산
    all_travel_mins = []
    all_unload_mins = []
    all_total_mins = []

    for d in deliveries:
        started = _ts_to_datetime(d.get("startedAt"))
        arrived = _ts_to_datetime(d.get("arrivedAt"))
        completed = _ts_to_datetime(d.get("completedAt"))

        if started and arrived:
            travel = (arrived - started).total_seconds() / 60
            if 0 < travel < 300:
                all_travel_mins.append(travel)
        if arrived and completed:
            unload = (completed - arrived).total_seconds() / 60
            if 0 < unload < 300:
                all_unload_mins.append(unload)
        if started and completed:
            total_t = (completed - started).total_seconds() / 60
            if 0 < total_t < 600:
                all_total_mins.append(total_t)

    avg_travel = sum(all_travel_mins) / len(all_travel_mins) if all_travel_mins else 0
    avg_unload = sum(all_unload_mins) / len(all_unload_mins) if all_unload_mins else 0
    avg_total = sum(all_total_mins) / len(all_total_mins) if all_total_mins else 0

    # 전체 요약 카드
    st.markdown(
        f'<div style="background:#1E293B;border-radius:10px;padding:14px;margin:8px 0">'
        f'<div style="font-size:0.9rem;font-weight:bold;margin-bottom:8px">📊 전체 요약</div>'
        f'<div style="display:flex;flex-wrap:wrap;gap:10px;font-size:0.82rem">'
        f'<div style="flex:1;min-width:100px;text-align:center">'
        f'<div style="color:#64748B">완료율</div>'
        f'<div style="font-size:1.3rem;font-weight:bold;color:#10B981">{pct}%</div>'
        f'<div style="color:#64748B;font-size:0.72rem">{delivered}/{total}</div></div>'
        f'<div style="flex:1;min-width:100px;text-align:center">'
        f'<div style="color:#64748B">평균 이동</div>'
        f'<div style="font-size:1.3rem;font-weight:bold;color:#3B82F6">{avg_travel:.0f}분</div></div>'
        f'<div style="flex:1;min-width:100px;text-align:center">'
        f'<div style="color:#64748B">평균 하차</div>'
        f'<div style="font-size:1.3rem;font-weight:bold;color:#F97316">{avg_unload:.0f}분</div></div>'
        f'<div style="flex:1;min-width:100px;text-align:center">'
        f'<div style="color:#64748B">평균 총소요</div>'
        f'<div style="font-size:1.3rem;font-weight:bold;color:#8B5CF6">{avg_total:.0f}분</div></div>'
        f'</div></div>',
        unsafe_allow_html=True,
    )

    # 기사별 분류
    by_driver: dict[str, list] = {}
    for d in deliveries:
        if d.get("assignedTo"):
            dname = d.get("driverName", "미지정")
            by_driver.setdefault(dname, []).append(d)

    unassigned = [d for d in deliveries if not d.get("assignedTo")]

    if not by_driver:
        st.info("배정된 기사가 없습니다.")
        return

    # 기사별 성과 카드
    for dname, driver_dels in by_driver.items():
        d_total = len(driver_dels)
        d_done = sum(1 for d in driver_dels if d.get("status") == "delivered")
        d_pct = int(d_done / d_total * 100) if d_total > 0 else 0

        # 기사별 시간 계산
        d_travel = []
        d_unload = []
        d_total_mins = []

        for d in driver_dels:
            started = _ts_to_datetime(d.get("startedAt"))
            arrived = _ts_to_datetime(d.get("arrivedAt"))
            completed = _ts_to_datetime(d.get("completedAt"))

            if started and arrived:
                t = (arrived - started).total_seconds() / 60
                if 0 < t < 300:
                    d_travel.append(t)
            if arrived and completed:
                u = (completed - arrived).total_seconds() / 60
                if 0 < u < 300:
                    d_unload.append(u)
            if started and completed:
                tt = (completed - started).total_seconds() / 60
                if 0 < tt < 600:
                    d_total_mins.append(tt)

        d_avg_travel = sum(d_travel) / len(d_travel) if d_travel else 0
        d_avg_unload = sum(d_unload) / len(d_unload) if d_unload else 0

        # 첫 출발 ~ 마지막 완료 총 근무시간
        started_times = []
        completed_times = []
        for d in driver_dels:
            s = _ts_to_datetime(d.get("startedAt"))
            c = _ts_to_datetime(d.get("completedAt"))
            if s:
                started_times.append(s)
            if c:
                completed_times.append(c)

        work_duration = ""
        if started_times and completed_times:
            first_start = min(started_times)
            last_complete = max(completed_times)
            work_mins = (last_complete - first_start).total_seconds() / 60
            if 0 < work_mins < 720:
                work_h = int(work_mins // 60)
                work_m = int(work_mins % 60)
                work_duration = f"{work_h}시간 {work_m}분" if work_h else f"{work_m}분"

        with st.expander(f"🚛 {dname} — {d_done}/{d_total} 완료 ({d_pct}%)"):
            # 성과 요약 HTML
            html = '<div style="font-size:0.82rem;line-height:1.7">'

            # 메트릭 행
            html += (
                f'<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px">'
                f'<span style="background:#0F172A;padding:3px 8px;border-radius:6px">'
                f'완료율 <b style="color:#10B981">{d_pct}%</b></span>'
                f'<span style="background:#0F172A;padding:3px 8px;border-radius:6px">'
                f'평균이동 <b style="color:#3B82F6">{d_avg_travel:.0f}분</b></span>'
                f'<span style="background:#0F172A;padding:3px 8px;border-radius:6px">'
                f'평균하차 <b style="color:#F97316">{d_avg_unload:.0f}분</b></span>'
            )
            if work_duration:
                html += (
                    f'<span style="background:#0F172A;padding:3px 8px;border-radius:6px">'
                    f'총근무 <b style="color:#8B5CF6">{work_duration}</b></span>'
                )
            html += '</div>'

            # 배송 타임라인
            html += '<div style="border-top:1px solid #334155;padding-top:8px">'
            for d in sorted(driver_dels, key=lambda x: x.get("order", 0)):
                s = d.get("status", "pending")
                s_color = STATUS_COLOR_MAP.get(s, "#94A3B8")
                s_label = {"pending": "배송출발", "in_transit": "배송출발", "arrived": "입고중",
                           "delivered": "완료", "issue": "문제"}.get(s, s)

                started = _ts_to_datetime(d.get("startedAt"))
                arrived = _ts_to_datetime(d.get("arrivedAt"))
                completed = _ts_to_datetime(d.get("completedAt"))

                # 시간 정보
                time_parts = []
                if started:
                    time_parts.append(f'출발 {started.strftime("%H:%M")}')
                if arrived:
                    time_parts.append(f'도착 {arrived.strftime("%H:%M")}')
                if completed:
                    time_parts.append(f'완료 {completed.strftime("%H:%M")}')
                time_str = " → ".join(time_parts) if time_parts else ""

                # 소요시간
                dur_str = ""
                if started and completed:
                    dur = (completed - started).total_seconds() / 60
                    if 0 < dur < 600:
                        dur_str = f' ({dur:.0f}분)'

                html += (
                    f'<div style="padding:4px 0;border-bottom:1px solid #1E293B;display:flex;align-items:center;gap:6px">'
                    f'<span style="color:#64748B;min-width:22px">#{d.get("order", 0)}</span>'
                    f'<span style="flex:1"><b>{d.get("storeName", "")}</b></span>'
                    f'<span style="color:{s_color};font-size:0.75rem;font-weight:600">{s_label}{dur_str}</span>'
                    f'</div>'
                )
                if time_str:
                    html += f'<div style="color:#64748B;font-size:0.72rem;padding:0 0 4px 28px">{time_str}</div>'

            html += '</div></div>'
            st.markdown(html, unsafe_allow_html=True)

    # 미배정 매장
    if unassigned:
        st.markdown(
            f'<div style="background:#7F1D1D;border-radius:8px;padding:10px;margin-top:10px;'
            f'font-size:0.82rem">'
            f'⚠️ 미배정 매장 <b>{len(unassigned)}개</b>: '
            + ", ".join(f'#{d.get("order", 0)} {d.get("storeName", "")}' for d in unassigned)
            + '</div>',
            unsafe_allow_html=True,
        )


# ─── Tab 5: 업로드 (기존 유지) ───────────────────────────────

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

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["현황", "기사 배정", "실시간 추적", "리포트", "업로드"])

    with tab1:
        _render_status_tab(date_compact)

    with tab2:
        _render_assign_tab(date_compact)

    with tab3:
        _render_tracking_tab(date_compact)

    with tab4:
        _render_report_tab(date_compact)

    with tab5:
        _render_upload_tab(date_compact)
