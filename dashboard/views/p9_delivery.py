"""
p9_delivery.py — 배송 관리 대시보드 페이지

탭:
  - 현황: 전체/완료/진행/대기 metric + 매장별 상태 테이블
  - 기사 배정: 기사 선택 → 매장 배정
  - 실시간 추적: 기사별 GPS 위치 (Phase 2)
  - 업로드: processed_orders → Firebase 업로드
"""

import streamlit as st
import json
import subprocess
import sys
from pathlib import Path
from datetime import datetime

# Paths
DASHBOARD_DIR = Path(__file__).resolve().parent.parent
PROJECT_ROOT = DASHBOARD_DIR.parent
EXECUTION_DIR = PROJECT_ROOT / "execution"
OUTPUTS_DIR = PROJECT_ROOT / "data" / "outputs"
INPUTS_DIR = PROJECT_ROOT / "data" / "inputs"


def _load_delivery_json(date_compact: str) -> dict | None:
    """deliveries_YYYYMMDD.json 로드 (Firebase 없이 로컬 확인용)"""
    json_path = OUTPUTS_DIR / f"deliveries_{date_compact}.json"
    if json_path.exists():
        with open(json_path, 'r', encoding='utf-8') as f:
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
            encoding='utf-8',
            timeout=60,
            cwd=str(PROJECT_ROOT),
        )
        output = result.stdout + (result.stderr or '')
        return result.returncode == 0, output
    except Exception as e:
        return False, str(e)


def render(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    st.header("배송 관리")

    tab1, tab2, tab3, tab4 = st.tabs(["현황", "기사 배정", "실시간 추적", "업로드"])

    # ─── Tab 1: 현황 ───────────────────────────────────────────
    with tab1:
        data = _load_delivery_json(date_compact)

        if data is None:
            st.info(f"{date_compact} 배송 데이터가 없습니다. '업로드' 탭에서 생성하세요.")
            return

        deliveries = data.get('deliveries', [])
        total = len(deliveries)
        pending = sum(1 for d in deliveries if d['status'] == 'pending')
        in_transit = sum(1 for d in deliveries if d['status'] == 'in_transit')
        delivered = sum(1 for d in deliveries if d['status'] == 'delivered')
        issue = sum(1 for d in deliveries if d['status'] == 'issue')

        # Metrics
        cols = st.columns(4)
        cols[0].metric("전체", total)
        cols[1].metric("대기", pending)
        cols[2].metric("배송중", in_transit)
        cols[3].metric("완료", delivered)

        if issue > 0:
            st.error(f"문제 발생: {issue}건")

        # Progress bar
        progress = delivered / total if total > 0 else 0
        st.progress(progress, text=f"진행률: {int(progress * 100)}%")

        # Store status table
        st.subheader("매장별 상태")

        status_emoji = {
            'pending': '⏳ 대기',
            'in_transit': '🚛 배송중',
            'delivered': '✅ 완료',
            'issue': '⚠️ 문제',
        }

        table_data = []
        for d in deliveries:
            item_count = len(d.get('items', []))
            red_count = sum(1 for i in d.get('items', []) if i.get('color') == 'red')
            blue_count = sum(1 for i in d.get('items', []) if i.get('color') == 'blue')

            color_info = ""
            if red_count or blue_count:
                parts = []
                if red_count:
                    parts.append(f"재고{red_count}")
                if blue_count:
                    parts.append(f"시장{blue_count}")
                color_info = f" ({', '.join(parts)})"

            table_data.append({
                "순서": d['order'],
                "매장": d['storeName'],
                "품목 수": f"{item_count}{color_info}",
                "기사": d.get('driverName') or '-',
                "상태": status_emoji.get(d['status'], d['status']),
            })

        st.dataframe(table_data, use_container_width=True, hide_index=True)

        # Expandable: store detail
        st.subheader("매장 상세")
        for d in deliveries:
            with st.expander(f"#{d['order']} {d['storeName']} ({len(d.get('items', []))}개 품목)"):
                for item in d.get('items', []):
                    color_tag = ""
                    if item.get('color') == 'red':
                        color_tag = " 🔴"
                    elif item.get('color') == 'blue':
                        color_tag = " 🔵"
                    st.text(f"  {item['name']}  {item['qty']}{color_tag}")

    # ─── Tab 2: 기사 배정 ──────────────────────────────────────
    with tab2:
        st.info("기사 배정은 모바일 앱(관리자 모드) 또는 Firebase Console에서 수행합니다.")
        st.markdown("""
        **배정 방법:**
        1. DoK Delivery 앱 → 관리자 로그인
        2. '기사 배정' 탭에서 기사 선택 + 매장 배정
        3. 기사에게 실시간 알림 전송

        **또는** Firebase Console에서 직접 `deliveryDays/{날짜}/deliveries/{매장}`의
        `assignedTo`, `driverName` 필드를 수정합니다.
        """)

    # ─── Tab 3: 실시간 추적 ────────────────────────────────────
    with tab3:
        st.info("실시간 위치 추적은 Phase 2에서 구현 예정입니다.")
        st.markdown("""
        **계획:**
        - 기사 앱에서 GPS 위치를 Firebase Realtime Database에 전송
        - 이 페이지에서 folium 지도로 실시간 위치 표시
        - 각 기사의 마지막 위치, 속도, 방향 표시
        """)

    # ─── Tab 4: 업로드 ─────────────────────────────────────────
    with tab4:
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
                    # Reload data
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
