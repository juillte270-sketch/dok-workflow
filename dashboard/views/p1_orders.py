"""Page 1: Order Processing (Stage 1) - Parse orders & generate delivery list."""
import sys
import json
import streamlit as st
from pathlib import Path
from dashboard.utils import (
    load_settings, run_script, mark_stage, get_stage_status,
    get_orders_path, get_processed_orders_path, INPUTS_DIR, OUTPUTS_DIR,
    append_log, format_elapsed, MAPPINGS_PATH, PROJECT_ROOT, EXECUTION_DIR,
)

# study04 for learning
STUDY04_ROOT = Path(r"C:\Users\DoKH_D\OneDrive\Desktop\study04")


# ──────────────────── Main Render ────────────────────

def render(get_date_str, get_date_compact):
    date_str = get_date_str()
    date_compact = get_date_compact()

    st.header("2. 발주 처리")

    settings = load_settings()
    status = get_stage_status("stage1", date_str)
    if status == "completed":
        st.success("오늘 발주 처리가 완료되었습니다.")

    st.divider()

    # --- 두 개의 탭: 자동 처리 / 최종 확정 입력 ---
    tab_auto, tab_final = st.tabs(["자동 처리", "최종 확정 입력"])

    with tab_auto:
        _render_auto_processing(date_str, date_compact, settings)

    with tab_final:
        _render_final_input(date_str, date_compact, settings)

    st.divider()

    # --- 처리 결과 & 배송리스트 (공통) ---
    _render_results(date_compact)


# ──────────────────── Tab 1: Auto Processing ────────────────────

def _render_auto_processing(date_str, date_compact, settings):
    """기존 자동 처리 탭."""
    st.subheader("발주 텍스트 입력")
    order_text = st.text_area(
        "카카오톡 발주 내용을 붙여넣으세요",
        height=300,
        placeholder="[봄날] 청경채 2\n케일 3\n깻잎 5...\n\n[파라이 역삼점]\n양배추 1망\n감자 1박스...",
        key="order_text_input",
    )

    uploaded = st.file_uploader("또는 파일 업로드 (.txt)", type=["txt"], key="order_file_upload")

    col1, col2, col3 = st.columns(3)

    # --- Process button ---
    with col1:
        can_run = bool(order_text) or uploaded is not None
        if st.button("발주 처리 시작", type="primary", disabled=not can_run, key="btn_process"):
            orders_path = INPUTS_DIR / f"orders_{date_compact}.txt"
            content = uploaded.read().decode("utf-8") if uploaded else order_text
            orders_path.parent.mkdir(parents=True, exist_ok=True)
            with open(orders_path, "w", encoding="utf-8") as f:
                f.write(content)
            append_log(f"발주 텍스트 저장: {orders_path.name} ({len(content)}자)")

            args = ["--input", str(orders_path), "--date", date_str]
            if settings.get("master_file"):
                args += ["--master-file", settings["master_file"]]
            with st.spinner("발주 처리 중... (재고 확인 포함)"):
                success, stdout, stderr, elapsed = run_script(
                    "process_orders.py", args,
                )
            if success:
                mark_stage("stage1", "completed", {"elapsed": elapsed}, date_str=date_str)
                st.toast(f"발주 처리 완료 ({format_elapsed(elapsed)})")
                st.rerun()
            else:
                mark_stage("stage1", "failed", {"error": stderr[-1000:]}, date_str=date_str)
                st.error("발주 처리 실패")
                st.code(stderr[-2000:], language="text")

    # --- Re-process existing ---
    with col2:
        existing = INPUTS_DIR / f"orders_{date_compact}.txt"
        if existing.exists():
            if st.button("기존 발주 재처리", key="btn_reprocess"):
                args = ["--input", str(existing), "--date", date_str]
                if settings.get("master_file"):
                    args += ["--master-file", settings["master_file"]]
                with st.spinner("재처리 중... (재고 확인 포함)"):
                    success, stdout, stderr, elapsed = run_script(
                        "process_orders.py", args,
                    )
                if success:
                    mark_stage("stage1", "completed", {"elapsed": elapsed}, date_str=date_str)
                    st.toast(f"재처리 완료 ({format_elapsed(elapsed)})")
                    st.rerun()
                else:
                    st.error("재처리 실패")
                    st.code(stderr[-2000:], language="text")
        else:
            st.info("기존 발주 파일 없음")

    # --- Load previous ---
    with col3:
        if existing.exists():
            with open(existing, "r", encoding="utf-8") as f:
                prev_content = f.read()
            st.download_button("원본 텍스트 다운로드", prev_content, file_name=existing.name)


# ──────────────────── Tab 2: Final Confirmed Input ────────────────────

def _render_final_input(date_str, date_compact, settings):
    """최종 확정 리스트 직접 입력 탭."""
    st.subheader("최종 확정 리스트 입력")
    st.caption("자동 처리 결과를 수정하거나, 최종 리스트를 직접 입력합니다.")

    # 기존 processed_orders에서 불러오기 버튼
    processed_path = INPUTS_DIR / f"processed_orders_{date_compact}.txt"
    if processed_path.exists():
        if st.button("기존 결과 불러오기", key="btn_load_existing"):
            with open(processed_path, "r", encoding="utf-8") as f:
                existing_text = f.read()
            customer_part, supplier_part = _split_customer_supplier(existing_text)
            st.session_state["final_customer_text"] = customer_part
            st.session_state["final_supplier_text"] = supplier_part
            st.rerun()

    # 두 텍스트 영역
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**고객 발주 리스트**")
        customer_text = st.text_area(
            "고객 발주 리스트",
            height=350,
            placeholder="-봄날\n청경채 2\n케일 3\n\n-파라이 역삼점\n양배추 1망\n감자 1박스",
            key="final_customer_text",
            label_visibility="collapsed",
        )

    with col_right:
        st.markdown("**공급처 리스트**")
        supplier_text = st.text_area(
            "공급처 리스트",
            height=350,
            placeholder=(
                "-영운농산\n깻잎 10개\n\n"
                "-재고, 창고소분\n양파 3\n\n"
                "-시장구매, 시장소분\n감자 5"
            ),
            key="final_supplier_text",
            label_visibility="collapsed",
        )

    has_content = bool(customer_text and customer_text.strip()) or bool(
        supplier_text and supplier_text.strip()
    )

    col_save, col_learn, col_workflow = st.columns(3)

    with col_save:
        if st.button(
            "확정 저장",
            type="primary",
            disabled=not has_content,
            key="btn_final_save",
        ):
            _save_final_and_generate(
                date_str, date_compact, customer_text or "", supplier_text or "", settings
            )

    with col_learn:
        orders_path = INPUTS_DIR / f"orders_{date_compact}.txt"
        has_original = orders_path.exists()
        if st.button(
            "학습 & 저장",
            disabled=not (has_content and has_original),
            key="btn_final_learn",
            help="원본 발주와 확정 공급처를 비교하여 품목→공급처 매핑을 학습합니다",
        ):
            _save_final_and_generate(
                date_str, date_compact, customer_text or "", supplier_text or "", settings
            )
            _learn_from_final(supplier_text or "")

    with col_workflow:
        if st.button(
            "워크플로우 ▶",
            disabled=not processed_path.exists(),
            key="btn_workflow_next",
            help="다음 단계(발주시트 입력)로 이동",
        ):
            st.toast("다음 단계로 이동합니다")


# ──────────────────── Save & Generate DOCX ────────────────────

def _save_final_and_generate(date_str, date_compact, customer_text, supplier_text, settings):
    """확정 리스트 저장 + DOCX 생성."""
    # 두 텍스트 합치기
    parts = []
    if customer_text.strip():
        parts.append(customer_text.strip())
    if supplier_text.strip():
        parts.append(supplier_text.strip())
    merged = "\n".join(parts) + "\n"

    # processed_orders 파일 저장
    processed_path = INPUTS_DIR / f"processed_orders_{date_compact}.txt"
    processed_path.parent.mkdir(parents=True, exist_ok=True)
    with open(processed_path, "w", encoding="utf-8") as f:
        f.write(merged)
    append_log(f"확정 리스트 저장: {processed_path.name} ({len(merged)}자)")

    # DOCX 생성 (simple_docx_converter 직접 호출)
    try:
        if str(EXECUTION_DIR) not in sys.path:
            sys.path.insert(0, str(EXECUTION_DIR))
        import simple_docx_converter

        mappings_file = str(MAPPINGS_PATH)
        docx_output = str(OUTPUTS_DIR / f"배송리스트_{date_compact}.docx")
        simple_docx_converter.create_docx(
            str(processed_path), mappings_file, docx_output, target_date=date_str
        )
        append_log(f"배송리스트 DOCX 생성: 배송리스트_{date_compact}.docx")
        mark_stage("stage1", "completed", {}, date_str=date_str)
        st.toast("확정 저장 + 배송리스트 생성 완료")
        st.rerun()
    except Exception as e:
        st.error(f"DOCX 생성 실패: {e}")
        append_log(f"DOCX 생성 실패: {e}", "ERROR")


# ──────────────────── Learning from Final List ────────────────────

def _learn_from_final(supplier_text: str):
    """확정 공급처 리스트에서 품목→공급처 매핑을 학습."""
    try:
        if str(STUDY04_ROOT) not in sys.path:
            sys.path.insert(0, str(STUDY04_ROOT))
        import app_optimized as engine
    except Exception as e:
        st.warning(f"학습 엔진 로드 실패: {e}")
        return

    # 확정 리스트 파싱: -공급처명 아래 품목들 추출
    final_mappings = _parse_supplier_items(supplier_text, engine)

    if not final_mappings:
        st.info("학습할 공급처 매핑이 없습니다.")
        return

    learned_count = 0
    try:
        with engine.atomic_json_update(engine.SUPPLIERS_FILE) as data:
            learned = data.setdefault("learned_suppliers", [])
            for supplier_name, item_names in final_mappings.items():
                # 재고/시장 섹션은 학습하지 않음
                if "재고" in supplier_name or "시장" in supplier_name:
                    continue
                for item_name in item_names:
                    # 기존 규칙 업데이트 or 추가
                    found = False
                    for entry in learned:
                        if entry.get("item") == item_name:
                            if entry.get("supplier") != supplier_name:
                                entry["supplier"] = supplier_name
                                learned_count += 1
                            found = True
                            break
                    if not found:
                        learned.append({
                            "item": item_name,
                            "supplier": supplier_name,
                        })
                        learned_count += 1
        engine._data_cache.invalidate()
    except Exception as e:
        st.warning(f"학습 저장 실패: {e}")
        return

    if learned_count > 0:
        st.toast(f"학습 완료: {learned_count}개 품목→공급처 매핑 저장")
        append_log(f"확정 리스트 학습: {learned_count}개 매핑 저장")
    else:
        st.toast("변경된 매핑이 없습니다")


def _parse_supplier_items(text: str, engine=None) -> dict:
    """공급처 리스트 텍스트에서 공급처별 품목 목록 추출.

    engine이 제공되면 parse_item_line()으로 정확한 품목명 추출.

    Returns: {supplier_name: [item_name, ...]}
    """
    result = {}
    current_supplier = None

    for line in text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("-"):
            current_supplier = line[1:].strip()
            if current_supplier not in result:
                result[current_supplier] = []
        elif current_supplier:
            # engine의 parse_item_line으로 품목명 추출
            if engine:
                try:
                    parsed = engine.parse_item_line(line)
                    if parsed:
                        result[current_supplier].append(parsed["item"])
                        continue
                except Exception:
                    pass
            # 폴백: 첫 번째 숫자 앞까지를 품목명으로
            parts = line.split()
            if parts:
                item = parts[0]
                i = 1
                while i < len(parts):
                    token = parts[i]
                    if token[0].isdigit():
                        break
                    item += " " + token
                    i += 1
                result[current_supplier].append(item)

    return result


# ──────────────────── Helper: Split Customer / Supplier ────────────────────

def _split_customer_supplier(text: str) -> tuple:
    """processed_orders 텍스트를 고객 부분과 공급처+재고/시장 부분으로 분리."""
    supplier_names = _load_supplier_names()

    lines = text.split("\n")
    customer_lines = []
    supplier_lines = []
    is_supplier_section = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("-"):
            name = stripped[1:].strip()
            if (
                name in supplier_names
                or name in ("재고, 창고소분", "시장구매, 시장소분")
            ):
                is_supplier_section = True

        if is_supplier_section:
            supplier_lines.append(line)
        else:
            customer_lines.append(line)

    return "\n".join(customer_lines).strip(), "\n".join(supplier_lines).strip()


# ──────────────────── Results Display ────────────────────

def _render_results(date_compact):
    """처리 결과 & 배송리스트 표시 (공통)."""
    processed_path = INPUTS_DIR / f"processed_orders_{date_compact}.txt"
    if processed_path.exists():
        st.subheader("처리 결과")
        with open(processed_path, "r", encoding="utf-8") as f:
            processed_text = f.read()

        # Count stores and items
        lines = processed_text.strip().split("\n")
        store_count = sum(1 for l in lines if l.strip().startswith("-"))
        item_count = sum(
            1
            for l in lines
            if l.strip()
            and not l.strip().startswith("-")
            and not l.strip().startswith("(")
        )

        st.info(f"구간: {store_count}개 | 품목 행: {item_count}개")

        sections = _parse_sections(processed_text)
        if sections:
            tabs = st.tabs(list(sections.keys()))
            for tab, (name, content) in zip(tabs, sections.items()):
                with tab:
                    st.code(content, language="text")
        else:
            with st.expander("전체 내용", expanded=True):
                st.code(processed_text[:5000], language="text")

    # --- Delivery list download ---
    docx_path = OUTPUTS_DIR / f"배송리스트_{date_compact}.docx"
    if docx_path.exists():
        st.subheader("배송리스트")
        size_kb = docx_path.stat().st_size / 1024
        st.success(f"배송리스트 생성됨 ({size_kb:.0f}KB)")
        with open(docx_path, "rb") as f:
            data = f.read()
        st.download_button(
            "배송리스트 다운로드 (.docx)",
            data=data,
            file_name=docx_path.name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


# ──────────────────── Section Parsing ────────────────────

def _load_supplier_names():
    """Load canonical supplier names from mappings.json."""
    try:
        with open(MAPPINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return set(data.get("supplier_names", []))
    except Exception:
        return set()


def _parse_sections(text):
    """Parse processed_orders text into named sections.

    Groups '-store' sections into: 고객매장 / 공급처 / 재고-시장.
    """
    suppliers = _load_supplier_names()

    customer_lines = []
    supplier_lines = []
    stock_market_lines = []

    current_bucket = customer_lines
    in_supplier = False

    for line in text.split("\n"):
        stripped = line.strip()
        if stripped.startswith("-"):
            name = stripped.lstrip("-").strip()
            if name in ("재고, 창고소분", "시장구매, 시장소분"):
                current_bucket = stock_market_lines
            elif name in suppliers:
                current_bucket = supplier_lines
                in_supplier = True
            elif in_supplier:
                # Once we've seen suppliers, anything new is still supplier
                current_bucket = supplier_lines
            else:
                current_bucket = customer_lines
        current_bucket.append(line)

    sections = {}
    if customer_lines:
        sections["고객매장"] = "\n".join(customer_lines).strip()
    if supplier_lines:
        sections["공급처"] = "\n".join(supplier_lines).strip()
    if stock_market_lines:
        sections["재고/시장"] = "\n".join(stock_market_lines).strip()

    return sections if len(sections) > 1 else None
