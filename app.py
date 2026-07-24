from __future__ import annotations

import os
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st
import streamlit.components.v1 as components

from excel_ai.models import Action, NormalizationOptions, OperationPlan
from excel_ai.operations import OperationError, execute_plan
from excel_ai.planner import PlannerError, create_plan
from excel_ai.schema import WorkbookValidationError, inspect_workbook

st.set_page_config(
    page_title="Excel Thuế Trợ Lý",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


APP_CSS = """
<style>
:root {
  --brand: #0f766e;
  --brand-dark: #115e59;
  --ink: #172033;
  --muted: #667085;
  --surface: #ffffff;
  --line: #e5e7eb;
  --soft: #f0fdfa;
}

.block-container {max-width: 1180px; padding-top: 1.4rem; padding-bottom: 4rem;}
[data-testid="stSidebar"] {border-right: 1px solid var(--line);}
[data-testid="stSidebar"] > div:first-child {padding-top: 1.3rem;}

.hero {
  padding: 1.35rem 1.5rem;
  border-radius: 22px;
  background: linear-gradient(135deg, #ecfdf5 0%, #f0fdfa 55%, #eff6ff 100%);
  border: 1px solid #ccfbf1;
  margin-bottom: 1.2rem;
}
.hero h1 {font-size: clamp(1.85rem, 4vw, 2.8rem); margin: 0; color: var(--ink); letter-spacing: -0.035em;}
.hero p {margin: .45rem 0 0; color: var(--muted); font-size: 1.02rem; max-width: 800px;}
.hero-badge {display: inline-block; font-size: .78rem; font-weight: 700; color: var(--brand-dark); background: #ccfbf1; padding: .3rem .65rem; border-radius: 999px; margin-bottom: .65rem;}

.step-card {
  height: 100%; padding: 1rem; border: 1px solid var(--line); border-radius: 16px;
  background: var(--surface); box-shadow: 0 8px 24px rgba(15, 23, 42, .04);
}
.step-number {display:inline-grid; place-items:center; width:30px; height:30px; border-radius:10px; background:var(--soft); color:var(--brand-dark); font-weight:800; margin-right:.45rem;}
.step-title {font-weight: 750; color: var(--ink);}
.step-copy {color: var(--muted); font-size: .9rem; margin-top: .45rem;}

.privacy-note {padding: .8rem 1rem; border-radius: 14px; background: #fffbeb; border: 1px solid #fde68a; color: #92400e; font-size: .92rem;}
.plan-box {padding: 1rem; border-radius: 16px; background: #f8fafc; border: 1px solid var(--line); margin: .65rem 0;}
.file-chip {display:inline-block; margin:.15rem .25rem .15rem 0; padding:.32rem .65rem; border-radius:999px; background:#f1f5f9; color:#334155; font-size:.82rem;}
.small-muted {color:var(--muted); font-size:.86rem;}

div[data-testid="stFileUploader"] section {border-radius: 18px; border: 1.5px dashed #5eead4; background: #f8fffd; padding: 1.2rem;}
div[data-testid="stFileUploader"] section:hover {border-color: var(--brand); background: #f0fdfa;}
.stButton > button, .stDownloadButton > button {border-radius: 12px; min-height: 2.65rem; font-weight: 700;}
.stButton > button[kind="primary"] {background: var(--brand); border-color: var(--brand);}
.stButton > button[kind="primary"]:hover {background: var(--brand-dark); border-color: var(--brand-dark);}
[data-testid="stMetric"] {background:#fff; border:1px solid var(--line); border-radius:14px; padding:.8rem;}

@media (max-width: 720px) {
  .block-container {padding-left: .9rem; padding-right: .9rem; padding-top: .75rem;}
  .hero {padding: 1rem; border-radius: 17px;}
  .hero h1 {font-size: 1.8rem;}
}
</style>
"""

st.markdown(APP_CSS, unsafe_allow_html=True)


def inject_pwa_metadata() -> None:
    components.html(
        """
        <script>
        (() => {
          const doc = window.parent.document;
          if (!doc.querySelector('link[data-excel-pwa="manifest"]')) {
            const manifest = doc.createElement('link');
            manifest.rel = 'manifest';
            manifest.href = '/app/static/manifest.json';
            manifest.dataset.excelPwa = 'manifest';
            doc.head.appendChild(manifest);
          }
          const metaValues = [
            ['theme-color', '#0f766e'],
            ['mobile-web-app-capable', 'yes'],
            ['apple-mobile-web-app-capable', 'yes'],
            ['apple-mobile-web-app-status-bar-style', 'default'],
            ['apple-mobile-web-app-title', 'Excel Thuế']
          ];
          metaValues.forEach(([name, content]) => {
            if (!doc.querySelector(`meta[name="${name}"]`)) {
              const meta = doc.createElement('meta');
              meta.name = name;
              meta.content = content;
              doc.head.appendChild(meta);
            }
          });
          if ('serviceWorker' in window.parent.navigator) {
            window.parent.navigator.serviceWorker.register('/app/static/sw.js').catch(() => {});
          }
        })();
        </script>
        """,
        height=0,
        width=0,
    )


inject_pwa_metadata()


@st.cache_data(show_spinner=False, max_entries=24)
def cached_inspect(data: bytes, filename: str, file_id: str):
    return inspect_workbook(data, filename, file_id)


def get_config_value(name: str, default: str = "") -> str:
    env_value = os.getenv(name)
    if env_value:
        return env_value
    try:
        value = st.secrets.get(name, default)
        return str(value) if value else default
    except Exception:
        return default


def column_catalog(schemas):
    labels: dict[str, str] = {}
    ids: list[str] = []
    for workbook in schemas:
        for sheet in workbook.sheets:
            for column in sheet.columns:
                ids.append(column.id)
                labels[column.id] = (
                    f"{workbook.display_name}  ›  {sheet.name}  ›  "
                    f"{column.letter} — {column.header} ({column.detected_type})"
                )
    return ids, labels


def render_schema_summary(schemas) -> None:
    with st.expander("Xem cấu trúc file mà hệ thống đã nhận diện", expanded=False):
        for workbook in schemas:
            st.markdown(f"**{workbook.display_name}**")
            rows = []
            for sheet in workbook.sheets:
                for column in sheet.columns:
                    rows.append(
                        {
                            "Sheet": sheet.name,
                            "Dòng tiêu đề": sheet.header_row,
                            "Cột": column.letter,
                            "Tên cột": column.header,
                            "Kiểu nhận diện": column.detected_type,
                            "Số dòng": sheet.max_row,
                        }
                    )
            if rows:
                st.dataframe(rows, use_container_width=True, hide_index=True)
            else:
                st.warning("Không tìm thấy cột có tiêu đề rõ ràng trong file này.")


def render_plan(plan: OperationPlan, labels: dict[str, str]) -> None:
    action_labels = {
        Action.COMPARE_COLUMNS: "Đối chiếu hai cột",
        Action.HIGHLIGHT_DUPLICATES: "Tìm dữ liệu trùng",
        Action.FIND_BLANKS: "Tìm ô trống",
    }
    source = labels.get(plan.source_column_id, plan.source_column_id)
    comparison = labels.get(plan.comparison_column_id or "", "")
    st.markdown(
        f"""
        <div class="plan-box">
          <strong>{action_labels[plan.action]}</strong><br>
          <span class="small-muted">{plan.explanation or 'Kế hoạch xử lý đã được tạo.'}</span><br><br>
          <strong>Cột nguồn:</strong> {source}<br>
          {f'<strong>Cột đối chiếu:</strong> {comparison}<br>' if comparison else ''}
          <strong>Điều kiện tô màu:</strong> {plan.highlight_condition}<br>
          <strong>Độ tin cậy:</strong> {plan.confidence:.0%}
        </div>
        """,
        unsafe_allow_html=True,
    )


def package_download(files: dict[str, bytes]) -> tuple[str, bytes, str]:
    if len(files) == 1:
        filename, data = next(iter(files.items()))
        suffix = Path(filename).suffix.lower()
        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if suffix == ".xlsm"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        return filename, data, mime

    buffer = BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for filename, data in files.items():
            archive.writestr(filename, data)
    return "ket_qua_excel.zip", buffer.getvalue(), "application/zip"


def run_plan(plan, schemas, uploaded_map) -> None:
    try:
        with st.spinner("Đang xử lý workbook trên máy chủ..."):
            result = execute_plan(plan, schemas, uploaded_map)
        st.session_state["last_result"] = result
        st.success(result.message)
    except (OperationError, ValueError) as exc:
        st.error(str(exc))
    except Exception as exc:
        st.error(f"Không thể xử lý file: {exc}")


def normalization_selector(prefix: str, default: str = "text_normalized") -> NormalizationOptions:
    mode_labels = {
        "text_normalized": "Văn bản chuẩn hóa",
        "digits_only": "Chỉ giữ chữ số",
        "number": "So sánh dạng số",
        "exact": "Khớp tuyệt đối",
    }
    modes = list(mode_labels)
    mode = st.selectbox(
        "Cách chuẩn hóa dữ liệu",
        modes,
        index=modes.index(default),
        format_func=lambda item: mode_labels[item],
        key=f"{prefix}_mode",
        help="Mã số thuế thường nên chọn ‘Chỉ giữ chữ số’. Số tiền nên chọn ‘So sánh dạng số’.",
    )
    col1, col2 = st.columns(2)
    ignore_case = col1.checkbox("Không phân biệt hoa/thường", value=True, key=f"{prefix}_case")
    remove_spaces = col2.checkbox("Bỏ khoảng trắng bên trong", value=False, key=f"{prefix}_spaces")
    return NormalizationOptions(
        mode=mode,
        ignore_case=ignore_case,
        trim_whitespace=True,
        remove_internal_spaces=remove_spaces,
    )


COLOR_OPTIONS = {
    "Vàng nhạt": "FFF59D",
    "Xanh lá nhạt": "BBF7D0",
    "Đỏ nhạt": "FECACA",
    "Xanh dương nhạt": "BFDBFE",
}


st.markdown(
    """
    <section class="hero">
      <span class="hero-badge">PWA · Xử lý tại backend · AI chỉ đọc metadata</span>
      <h1>Excel Thuế Trợ Lý</h1>
      <p>Tải file lên, mô tả yêu cầu bằng tiếng Việt hoặc chọn thao tác thủ công. Dữ liệu trong ô không được gửi cho Gemini.</p>
    </section>
    """,
    unsafe_allow_html=True,
)

step_cols = st.columns(3)
for col, number, title, copy in zip(
    step_cols,
    ["1", "2", "3"],
    ["Tải file", "Chọn cách xử lý", "Kiểm tra và tải xuống"],
    [
        "Hỗ trợ tối đa 2 file .xlsx hoặc .xlsm trong một lần xử lý.",
        "Dùng trợ lý AI hoặc các công cụ có sẵn, không cần viết công thức.",
        "Luôn xem lại kế hoạch trước khi hệ thống thay đổi màu trong workbook.",
    ],
):
    with col:
        st.markdown(
            f'<div class="step-card"><span class="step-number">{number}</span><span class="step-title">{title}</span><div class="step-copy">{copy}</div></div>',
            unsafe_allow_html=True,
        )

st.write("")

with st.sidebar:
    st.markdown("### ⚙️ Cài đặt")
    configured_key = get_config_value("GEMINI_API_KEY")
    session_key = st.text_input(
        "Gemini API key cho phiên hiện tại",
        type="password",
        placeholder="Để trống nếu đã cấu hình secret",
        help="Key chỉ được giữ trong session của ứng dụng và không ghi vào file Excel.",
    )
    api_key = session_key.strip() or configured_key
    model = st.text_input(
        "Gemini model",
        value=get_config_value("GEMINI_MODEL", "gemini-2.5-flash-lite"),
    ).strip()
    if api_key:
        st.success("Gemini đã sẵn sàng", icon="✅")
    else:
        st.info("Chưa có API key — vẫn dùng được toàn bộ công cụ thủ công.", icon="ℹ️")

    st.divider()
    st.markdown("### 🔐 Quyền riêng tư")
    st.caption(
        "Gemini chỉ nhận tên file, sheet, tên cột, kiểu dữ liệu và số dòng. "
        "Giá trị thực tế trong các ô được xử lý bằng openpyxl trên backend."
    )
    st.caption("Không tự động xoay nhiều API key để né hạn mức Free Tier.")

st.markdown("### 1. Tải workbook")
uploaded = st.file_uploader(
    "Kéo thả file vào đây hoặc bấm để chọn",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Tối đa 2 file. Mỗi file nên nhỏ hơn 50 MB.",
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown(
        """
        <div class="privacy-note">
          <strong>Mẹo:</strong> Với yêu cầu “so sánh mã số thuế giữa hai file, trùng thì tô vàng”, hãy tải cả hai file rồi mở tab <strong>Trợ lý AI</strong> hoặc <strong>Đối chiếu 2 cột</strong>.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

if len(uploaded) > 2:
    st.warning("Ứng dụng chỉ xử lý 2 file đầu tiên trong mỗi lượt.")
    uploaded = uploaded[:2]

schemas = []
uploaded_map: dict[str, tuple[str, bytes]] = {}
validation_failed = False
for index, file in enumerate(uploaded, start=1):
    file_id = f"F{index}"
    data = file.getvalue()
    try:
        schema = cached_inspect(data, file.name, file_id)
        schemas.append(schema)
        uploaded_map[file_id] = (file.name, data)
    except WorkbookValidationError as exc:
        st.error(f"{file.name}: {exc}")
        validation_failed = True
    except Exception as exc:
        st.error(f"Không đọc được {file.name}: {exc}")
        validation_failed = True

if validation_failed or not schemas:
    st.stop()

chips = "".join(
    f'<span class="file-chip">📄 {schema.display_name} · {len(schema.sheets)} sheet</span>'
    for schema in schemas
)
st.markdown(chips, unsafe_allow_html=True)
render_schema_summary(schemas)

column_ids, labels = column_catalog(schemas)
if not column_ids:
    st.error("Không tìm thấy cột nào để xử lý. Hãy kiểm tra lại dòng tiêu đề trong workbook.")
    st.stop()

st.markdown("### 2. Chọn cách xử lý")
ai_tab, compare_tab, duplicate_tab, blank_tab = st.tabs(
    ["✨ Trợ lý AI", "🔁 Đối chiếu 2 cột", "🧩 Tìm dữ liệu trùng", "⬜ Tìm ô trống"]
)

with ai_tab:
    st.caption("Gemini chỉ nhận bản đồ workbook, không nhận dữ liệu trong ô.")
    request = st.text_area(
        "Mô tả yêu cầu",
        placeholder="Ví dụ: So sánh cột Mã số thuế ở file bán ra với MST khách hàng ở file kế toán, nếu trùng thì tô vàng ở file bán ra.",
        height=120,
    )
    examples = st.columns(3)
    examples[0].caption("“Tìm số hóa đơn bị trùng và tô đỏ.”")
    examples[1].caption("“Đánh dấu ô trống trong cột Mã số thuế.”")
    examples[2].caption("“So sánh hai cột, không trùng thì tô xanh.”")

    if st.button("Phân tích yêu cầu", type="primary", use_container_width=True, disabled=not api_key):
        try:
            with st.spinner("Gemini đang đọc cấu trúc file và lập kế hoạch..."):
                st.session_state["ai_plan"] = create_plan(request, schemas, api_key, model)
        except PlannerError as exc:
            st.error(str(exc))

    plan = st.session_state.get("ai_plan")
    if plan:
        render_plan(plan, labels)
        if plan.needs_confirmation:
            st.warning(plan.question or "Có điểm chưa rõ. Hãy kiểm tra lại cột đã chọn.")

        source_index = column_ids.index(plan.source_column_id) if plan.source_column_id in column_ids else 0
        selected_source = st.selectbox(
            "Cột nguồn",
            column_ids,
            index=source_index,
            format_func=lambda item: labels[item],
            key="ai_source",
        )
        selected_comparison = plan.comparison_column_id
        if plan.action == Action.COMPARE_COLUMNS:
            comparison_index = (
                column_ids.index(plan.comparison_column_id)
                if plan.comparison_column_id in column_ids
                else min(1, len(column_ids) - 1)
            )
            selected_comparison = st.selectbox(
                "Cột đối chiếu",
                column_ids,
                index=comparison_index,
                format_func=lambda item: labels[item],
                key="ai_comparison",
            )

        checked = st.checkbox(
            "Tôi đã kiểm tra đúng file, sheet, cột và đồng ý chạy kế hoạch này.",
            key="ai_confirm",
        )
        if st.button("Chạy kế hoạch", type="primary", use_container_width=True, disabled=not checked):
            confirmed_plan = plan.model_copy(
                update={
                    "source_column_id": selected_source,
                    "comparison_column_id": selected_comparison,
                    "needs_confirmation": False,
                    "question": None,
                }
            )
            run_plan(confirmed_plan, schemas, uploaded_map)

with compare_tab:
    with st.form("compare_form"):
        source_column = st.selectbox(
            "Cột nguồn",
            column_ids,
            format_func=lambda item: labels[item],
            key="manual_compare_source",
        )
        default_compare_index = 1 if len(column_ids) > 1 else 0
        comparison_column = st.selectbox(
            "Cột dùng để đối chiếu",
            column_ids,
            index=default_compare_index,
            format_func=lambda item: labels[item],
            key="manual_compare_target",
        )
        col_a, col_b, col_c = st.columns(3)
        condition = col_a.selectbox(
            "Tô màu khi",
            ["matched", "unmatched"],
            format_func=lambda item: "Dữ liệu trùng" if item == "matched" else "Dữ liệu không trùng",
        )
        target = col_b.selectbox(
            "Tô màu ở",
            ["source", "comparison", "both"],
            format_func=lambda item: {"source": "Cột nguồn", "comparison": "Cột đối chiếu", "both": "Cả hai cột"}[item],
        )
        color_name = col_c.selectbox("Màu đánh dấu", list(COLOR_OPTIONS))
        normalization = normalization_selector("compare")
        submitted = st.form_submit_button("Đối chiếu và tạo file", type="primary", use_container_width=True)

    if submitted:
        plan = OperationPlan(
            action=Action.COMPARE_COLUMNS,
            source_column_id=source_column,
            comparison_column_id=comparison_column,
            highlight_target=target,
            highlight_condition=condition,
            color=COLOR_OPTIONS[color_name],
            normalization=normalization,
            explanation="Kế hoạch đối chiếu được tạo từ biểu mẫu thủ công.",
        )
        run_plan(plan, schemas, uploaded_map)

with duplicate_tab:
    with st.form("duplicates_form"):
        duplicate_column = st.selectbox(
            "Cột cần kiểm tra trùng",
            column_ids,
            format_func=lambda item: labels[item],
            key="duplicate_column",
        )
        color_name = st.selectbox("Màu đánh dấu", list(COLOR_OPTIONS), key="duplicate_color")
        normalization = normalization_selector("duplicate")
        submitted = st.form_submit_button("Tìm dữ liệu trùng", type="primary", use_container_width=True)

    if submitted:
        plan = OperationPlan(
            action=Action.HIGHLIGHT_DUPLICATES,
            source_column_id=duplicate_column,
            highlight_condition="duplicates",
            color=COLOR_OPTIONS[color_name],
            normalization=normalization,
            explanation="Tìm các giá trị xuất hiện từ hai lần trở lên trong cột đã chọn.",
        )
        run_plan(plan, schemas, uploaded_map)

with blank_tab:
    with st.form("blanks_form"):
        blank_column = st.selectbox(
            "Cột bắt buộc cần kiểm tra",
            column_ids,
            format_func=lambda item: labels[item],
            key="blank_column",
        )
        color_name = st.selectbox("Màu đánh dấu", list(COLOR_OPTIONS), index=2, key="blank_color")
        submitted = st.form_submit_button("Đánh dấu ô trống", type="primary", use_container_width=True)

    if submitted:
        plan = OperationPlan(
            action=Action.FIND_BLANKS,
            source_column_id=blank_column,
            highlight_condition="blank",
            color=COLOR_OPTIONS[color_name],
            explanation="Đánh dấu các ô không có giá trị trong cột đã chọn.",
        )
        run_plan(plan, schemas, uploaded_map)

result = st.session_state.get("last_result")
if result:
    st.divider()
    st.markdown("### 3. Kết quả")
    metric_cols = st.columns(4)
    metric_cols[0].metric("Dòng đã quét", f"{result.stats.scanned_rows:,}")
    metric_cols[1].metric("Ô đã tô màu", f"{result.stats.highlighted_cells:,}")
    metric_cols[2].metric("Giá trị trùng", f"{result.stats.matched_values or result.stats.duplicate_values:,}")
    metric_cols[3].metric("Không trùng / trống", f"{result.stats.unmatched_values or result.stats.blank_cells:,}")

    filename, data, mime = package_download(result.files)
    st.download_button(
        "⬇️ Tải file kết quả",
        data=data,
        file_name=filename,
        mime=mime,
        type="primary",
        use_container_width=True,
    )
    st.caption("File gốc không bị ghi đè. Hãy mở file kết quả và kiểm tra trước khi dùng cho nghiệp vụ thuế chính thức.")
