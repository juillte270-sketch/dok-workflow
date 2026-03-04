"""Page 0b: 발주 정규화 / 공급처 분류 — study04 엔진 직접 호출."""
import sys
import json
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

# study04 모듈 직접 임포트
STUDY04_ROOT = Path(r"C:\Users\DoKH_D\OneDrive\Desktop\study04")
if str(STUDY04_ROOT) not in sys.path:
    sys.path.insert(0, str(STUDY04_ROOT))


# ──────────────────── Engine Loading ────────────────────

def _load_engine():
    """study04 모듈 전체를 lazy-load하여 반환.

    모듈 전체를 반환하여 convert_order, convert_stage2, parse_item_line 뿐 아니라
    atomic_json_update, _add_or_update_learned_rule, _data_cache 등 학습 함수에도 접근.
    """
    import app_optimized as engine
    return engine


# ──────────────────── Conversion ────────────────────

def _convert_full(text: str) -> dict:
    """convert_full API 로직을 직접 실행."""
    engine = _load_engine()

    stage1_result = engine.convert_order(text)

    items_for_stage2 = []
    for item in stage1_result.items:
        if item.rule_type == "customer":
            continue
        parsed = engine.parse_item_line(item.converted)
        if parsed:
            parsed["source"] = item.customer_alias or ""
            parsed["customer_group"] = item.customer_group or ""
            items_for_stage2.append(parsed)

    stage2_result = engine.convert_stage2(items_for_stage2)

    s1 = stage1_result.dict() if hasattr(stage1_result, "dict") else stage1_result.model_dump()
    s2 = stage2_result.dict() if hasattr(stage2_result, "dict") else stage2_result.model_dump()
    return {"stage1": s1, "stage2": s2}


# ──────────────────── Supplier List ────────────────────

def _get_all_suppliers() -> list:
    """suppliers.json에서 전체 공급처 이름 목록 로드."""
    suppliers_file = STUDY04_ROOT / "data" / "suppliers.json"
    try:
        with open(suppliers_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        names = [s["name"] for s in data.get("suppliers", [])]
        names.extend(["재고, 창고소분", "시장구매, 시장소분"])
        return names
    except Exception:
        return []


# ──────────────────── Learning Functions ────────────────────

def _learn_stage1(orig: str, corrected: str, group: str, store: str):
    """Stage1 학습: mappings.json learned[] 에 규칙 추가."""
    engine = _load_engine()
    parsed = engine.parse_item_line(corrected)
    default_unit = parsed.get("unit") if parsed else None

    with engine.atomic_json_update(engine.MAPPINGS_FILE) as mappings:
        engine._add_or_update_learned_rule(
            mappings, orig, corrected, group, store, default_unit
        )
    engine._data_cache.invalidate()


def _learn_stage2(item: str, new_supplier: str, group: str, store: str = None):
    """Stage2 학습: suppliers.json learned_suppliers[] 에 규칙 추가."""
    engine = _load_engine()
    with engine.atomic_json_update(engine.SUPPLIERS_FILE) as data:
        learned = data.setdefault("learned_suppliers", [])
        # 기존 규칙 업데이트 or 추가
        for entry in learned:
            if entry.get("item") == item and entry.get("customer_group") == group:
                entry["supplier"] = new_supplier
                break
        else:
            learned.append({
                "item": item,
                "supplier": new_supplier,
                "customer_group": group,
                "customer_store": store,
            })
    engine._data_cache.invalidate()


# ──────────────────── Copy Box ────────────────────

def _copy_box(text: str, box_id: str, height: int = 300):
    """study04 스타일 텍스트 박스 + 클립보드 복사 버튼."""
    escaped = (
        text.replace("&", "&amp;").replace("<", "&lt;")
        .replace(">", "&gt;").replace('"', "&quot;")
    )
    html = f"""
    <div style="position:relative; font-family:'Consolas','SF Mono','Fira Code',monospace;
                background:#1E293B; border:1px solid #334155; border-radius:8px;
                padding:12px 16px; padding-top:40px; color:#E2E8F0;
                font-size:14px; line-height:1.6; white-space:pre-wrap;
                max-height:{height}px; overflow-y:auto;">
        <button onclick="
            navigator.clipboard.writeText(document.getElementById('{box_id}').textContent).then(()=>{{
                this.textContent='Copied!';
                setTimeout(()=>this.textContent='Copy', 1500);
            }})
        " style="position:absolute; top:8px; right:8px; background:#334155; color:#E2E8F0;
                 border:1px solid #475569; border-radius:6px; padding:4px 12px;
                 cursor:pointer; font-size:12px; font-family:sans-serif;
                 transition:all 0.15s ease;"
           onmouseover="this.style.background='#818CF8'; this.style.borderColor='#818CF8'"
           onmouseout="this.style.background='#334155'; this.style.borderColor='#475569'">
            Copy
        </button>
        <div id="{box_id}">{escaped}</div>
    </div>
    """
    components.html(html, height=min(height + 60, 500), scrolling=True)


# ──────────────────── Main Render ────────────────────

def render(get_date_str, get_date_compact):
    st.header("발주 정규화 / 공급처 분류")
    st.caption("study04 발주변환 엔진 — Stage1(정규화) + Stage2(공급처 분류)")

    # --- 엔진 로드 확인 ---
    try:
        _load_engine()
        engine_ok = True
    except Exception as e:
        engine_ok = False
        st.error(f"study04 엔진 로드 실패: {e}")

    st.divider()

    # --- 입력 ---
    st.subheader("발주 텍스트 입력")
    order_text = st.text_area(
        "카카오톡 발주 내용을 붙여넣으세요",
        height=250,
        placeholder=(
            "-샤브야키 동탄점\n"
            "감자 10kg\n"
            "청경채 2박스\n"
            "깐대파 5kg\n"
            "청양고추 1박스\n"
            "팽이 3개"
        ),
        key="normalize_input",
    )

    col_run, col_reconvert, _ = st.columns([1, 1, 2])
    with col_run:
        run_clicked = st.button(
            "변환 실행", type="primary",
            disabled=(not order_text or not engine_ok),
            key="btn_normalize",
        )
    with col_reconvert:
        reconvert_clicked = st.button(
            "재변환 (학습 반영)",
            disabled=(
                not order_text
                or not engine_ok
                or "normalize_result" not in st.session_state
            ),
            key="btn_reconvert",
        )

    # --- 변환 실행 ---
    if (run_clicked or reconvert_clicked) and order_text:
        with st.spinner("Stage1 + Stage2 변환 중..."):
            try:
                result = _convert_full(order_text)
                st.session_state["normalize_result"] = result
                if reconvert_clicked:
                    st.toast("재변환 완료 — 학습 결과가 반영되었습니다")
            except Exception as e:
                st.error(f"변환 오류: {e}")
                return

    # --- 결과 표시 ---
    result = st.session_state.get("normalize_result")
    if not result:
        return

    st.divider()

    stage1 = result.get("stage1", {})
    stage2 = result.get("stage2", {})

    tab1, tab2, tab3 = st.tabs(["Stage 1: 정규화", "Stage 2: 공급처 분류", "원본 JSON"])

    # ── Stage 1: 거래처별 그룹 + 인라인 수정 ──
    with tab1:
        cg = stage1.get("customer_group", "—")
        cs = stage1.get("customer_store", "—")
        st.markdown(f"**거래처:** {cg}  /  **매장:** {cs}")

        items = stage1.get("items", [])
        sub1a, sub1b = st.tabs(["품목별 상세 (수정 가능)", "텍스트 (복사용)"])

        with sub1a:
            _render_stage1_editable(items, stage1)

        with sub1b:
            full = stage1.get("full_text", "")
            if full:
                _copy_box(full, "s1_text", height=300)

    # ── Stage 2: 공급처별 + 인라인 수정 ──
    with tab2:
        sub2a, sub2b = st.tabs(["공급처별 상세 (수정 가능)", "텍스트 (복사용)"])

        with sub2a:
            _render_stage2_editable(stage2)

        with sub2b:
            full2 = stage2.get("full_text", "")
            if full2:
                _copy_box(full2, "s2_text", height=400)

    # ── 원본 JSON ──
    with tab3:
        st.json(result)


# ──────────────────── Stage 1: Editable ────────────────────

_BADGE = {
    "learned": ("#22C55E", "learned"),
    "pattern": ("#3B82F6", "pattern"),
    "exact":   ("#A855F7", "exact"),
}


def _render_stage1_editable(items: list, stage1_data: dict):
    """Stage1 결과를 거래처별 그룹으로 표시 + 인라인 수정 & 학습."""
    if not items:
        st.info("변환 결과가 없습니다.")
        return

    current_group = stage1_data.get("customer_group", "") or ""
    current_store = stage1_data.get("customer_store", "") or ""

    for idx, it in enumerate(items):
        rule = it.get("rule_type", "unknown")

        # 거래처 라인 → 그룹 헤더로 표시
        if rule == "customer":
            name = it.get("converted", it.get("original", "")).lstrip("-").strip()
            current_group = it.get("customer_group", "") or current_group
            current_store = it.get("customer_alias", "") or current_store
            st.markdown(f"---\n#### {name}")
            continue

        # 품목 행
        orig = it.get("original", "")
        conv = it.get("converted", "")
        conf = it.get("confidence", 0)

        color, badge_text = _BADGE.get(rule, ("#94A3B8", "unknown"))

        # 레이아웃: 원본 | 변환결과(수정가능) | 뱃지 | 학습버튼
        c1, c2, c3, c4 = st.columns([2, 3, 1.5, 1])

        with c1:
            st.markdown(f"`{orig}`")

        with c2:
            edited = st.text_input(
                "변환", value=conv,
                key=f"s1_edit_{idx}",
                label_visibility="collapsed",
            )

        with c3:
            st.markdown(
                f"<span style='background:{color}; color:white; "
                f"padding:2px 8px; border-radius:4px; font-size:12px;'>"
                f"{badge_text}</span> "
                f"<small style='color:#64748B'>{conf:.0%}</small>",
                unsafe_allow_html=True,
            )

        with c4:
            is_modified = edited != conv
            grp = it.get("customer_group", "") or current_group
            sto = it.get("customer_alias", "") or current_store

            if st.button(
                "학습", key=f"s1_learn_{idx}",
                disabled=not is_modified,
                type="primary" if is_modified else "secondary",
            ):
                try:
                    _learn_stage1(orig, edited, grp, sto)
                    st.toast(f"학습 완료: {orig} → {edited}")
                except Exception as e:
                    st.error(f"학습 실패: {e}")


# ──────────────────── Stage 2: Editable ────────────────────

def _render_stage2_editable(stage2_data: dict):
    """Stage2 결과를 공급처별로 표시 + 공급처 변경 & 학습."""
    suppliers = stage2_data.get("suppliers", [])
    if not suppliers:
        st.info("공급처 분류 결과가 없습니다.")
        return

    all_supplier_names = _get_all_suppliers()

    for sup_idx, sup in enumerate(suppliers):
        name = sup.get("supplier", "?")
        items_list = sup.get("items", [])
        preorder = sup.get("preorder_items", [])

        if "재고" in name or "창고" in name:
            icon = "📦"
        elif "시장" in name:
            icon = "🏪"
        else:
            icon = "🚚"

        st.markdown(f"#### {icon} {name}")

        for item_idx, entry in enumerate(items_list):
            item = entry.get("item", "")
            qty = entry.get("qty", 0)
            unit = entry.get("unit", "")
            source = entry.get("source", "")
            cgroup = entry.get("customer_group", "")

            qty_str = f"{qty:g}" if isinstance(qty, float) else str(qty)
            item_display = f"{item} {qty_str}{unit}"
            if source:
                item_display += f" ({source})"

            # 레이아웃: 품목+수량 | 공급처 셀렉트 | 학습 버튼
            c1, c2, c3 = st.columns([3, 2, 1])

            with c1:
                st.text(item_display)

            with c2:
                default_idx = 0
                if all_supplier_names and name in all_supplier_names:
                    try:
                        default_idx = all_supplier_names.index(name)
                    except ValueError:
                        pass

                selected = st.selectbox(
                    "공급처",
                    options=all_supplier_names,
                    index=default_idx,
                    key=f"s2_sup_{sup_idx}_{item_idx}",
                    label_visibility="collapsed",
                )

            with c3:
                is_changed = selected != name
                if st.button(
                    "학습", key=f"s2_learn_{sup_idx}_{item_idx}",
                    disabled=not is_changed,
                    type="primary" if is_changed else "secondary",
                ):
                    try:
                        _learn_stage2(item, selected, cgroup)
                        st.toast(f"학습 완료: {item} → {selected}")
                    except Exception as e:
                        st.error(f"학습 실패: {e}")

        if preorder:
            st.markdown("*선발주:*")
            for po in preorder:
                st.markdown(
                    f"- {po.get('item', '')} {po.get('qty', '')}{po.get('unit', '')}"
                )

        st.markdown("")
