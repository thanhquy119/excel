from __future__ import annotations

import hashlib
import os
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st
import streamlit.components.v1 as components

from excel_ai.operations import OperationError, execute_plan
from excel_ai.planner import PlannerError, create_plan
from excel_ai.schema import WorkbookValidationError, inspect_workbook

st.set_page_config(
    page_title="Excel Thuế Trợ Lý",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_CSS = """
<style>
:root {
  --brand: #0f766e;
  --brand-dark: #115e59;
  --ink: #172033;
  --muted: #667085;
  --line: #e5e7eb;
}

.block-container {
  max-width: 1080px;
  padding-top: 1.35rem;
  padding-bottom: 4rem;
}

[data-testid="stSidebar"],
[data-testid="collapsedControl"] {
  display: none !important;
}

.hero {
  padding: 1.5rem 1.65rem;
  border-radius: 22px;
  background: linear-gradient(135deg, #ecfdf5 0%, #f0fdfa 55%, #eff6ff 100%);
  border: 1px solid #ccfbf1;
  margin-bottom: 1.35rem;
}

.hero h1 {
  font-size: clamp(1.9rem, 4vw, 2.75rem);
  margin: 0;
  color: var(--ink);
  letter-spacing: -0.035em;
}

.hero p {
  margin: .55rem 0 0;
  color: var(--muted);
  font-size: 1.02rem;
  max-width: 820px;
}

.file-chip {
  display: inline-block;
  margin: .2rem .35rem .35rem 0;
  padding: .35rem .72rem;
  border-radius: 999px;
  background: #f1f5f9;
  color: #334155;
  font-size: .85rem;
}

.examples {
  color: var(--muted);
  font-size: .9rem;
  line-height: 1.65;
  margin: .25rem 0 1rem;
}

div[data-testid="stFileUploader"] section {
  border-radius: 18px;
  border: 1.5px dashed #5eead4;
  background: #f8fffd;
  padding: 1.2rem;
}

div[data-testid="stFileUploader"] section:hover {
  border-color: var(--brand);
  background: #f0fdfa;
}

.stButton > button,
.stDownloadButton > button {
  border-radius: 12px;
  min-height: 2.75rem;
  font-weight: 700;
}

.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {
  background: var(--brand);
  border-color: var(--brand);
}

.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover {
  background: var(--brand-dark);
  border-color: var(--brand-dark);
}

@media (max-width: 720px) {
  .block-container {
    padding-left: .9rem;
    padding-right: .9rem;
    padding-top: .75rem;
  }
  .hero {
    padding: 1.1rem;
    border-radius: 17px;
  }
  .hero h1 {
    font-size: 1.85rem;
  }
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


def input_signature(uploaded_files, request: str) -> str:
    digest = hashlib.sha256(request.strip().encode("utf-8"))
    for file in uploaded_files:
        digest.update(file.name.encode("utf-8"))
        digest.update(file.getvalue())
    return digest.hexdigest()


st.markdown(
    """
    <section class="hero">
      <h1>Excel Thuế Trợ Lý</h1>
      <p>Tải workbook và mô tả yêu cầu bằng tiếng Việt. Gemini chỉ đọc cấu trúc file; dữ liệu trong các ô được xử lý trên backend.</p>
    </section>
    """,
    unsafe_allow_html=True,
)

st.subheader("Tải workbook")
uploaded = st.file_uploader(
    "Kéo thả file vào đây hoặc bấm để chọn",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Tối đa 2 file. Mỗi file nên nhỏ hơn 50 MB.",
    label_visibility="collapsed",
)

if len(uploaded) > 2:
    st.warning("Ứng dụng chỉ xử lý 2 file đầu tiên trong mỗi lượt.")
    uploaded = uploaded[:2]

schemas = []
uploaded_map: dict[str, tuple[str, bytes]] = {}
if uploaded:
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

    if validation_failed:
        st.stop()

    chips = "".join(
        f'<span class="file-chip">{schema.display_name} · {len(schema.sheets)} sheet</span>'
        for schema in schemas
    )
    st.markdown(chips, unsafe_allow_html=True)

st.subheader("Mô tả yêu cầu")
request = st.text_area(
    "Bạn muốn xử lý file như thế nào?",
    placeholder=(
        "Ví dụ: So sánh cột Mã số thuế ở file bán ra với cột MST ở file kế toán, "
        "nếu không trùng thì tô đỏ cả dòng trong file bán ra."
    ),
    height=140,
    label_visibility="collapsed",
)

st.markdown(
    """
    <div class="examples">
      Có thể yêu cầu đối chiếu hai cột; tìm dữ liệu trùng hoặc ô trống; kiểm tra định dạng mã số thuế;
      tô số tiền theo ngưỡng; lọc theo điều kiện ngày hoặc nội dung văn bản; tìm giá trị lớn nhất, nhỏ nhất,
      người già nhất hoặc trẻ nhất; và tô riêng ô hoặc toàn bộ dòng.
    </div>
    """,
    unsafe_allow_html=True,
)

api_key = get_config_value("GEMINI_API_KEY")
model = get_config_value("GEMINI_MODEL", "gemini-3.6-flash")
can_process = bool(uploaded and schemas and request.strip() and api_key)

if not api_key:
    st.error("Ứng dụng chưa được cấu hình GEMINI_API_KEY trong Streamlit Secrets.")

if st.button("Xử lý file", type="primary", use_container_width=True, disabled=not can_process):
    current_signature = input_signature(uploaded, request)
    st.session_state.pop("last_result", None)
    st.session_state.pop("last_signature", None)
    try:
        with st.spinner("Đang phân tích yêu cầu và xử lý workbook..."):
            plan = create_plan(request, schemas, api_key, model)
            if plan.needs_confirmation:
                st.warning(plan.question or "Yêu cầu chưa đủ rõ. Hãy bổ sung tên file, sheet, cột hoặc điều kiện.")
            else:
                result = execute_plan(plan, schemas, uploaded_map)
                st.session_state["last_result"] = result
                st.session_state["last_signature"] = current_signature
    except (PlannerError, OperationError, ValueError) as exc:
        st.error(str(exc))
    except Exception as exc:
        st.error(f"Không thể xử lý file: {exc}")

result = st.session_state.get("last_result")
last_signature = st.session_state.get("last_signature")
if result and uploaded and last_signature == input_signature(uploaded, request):
    filename, data, mime = package_download(result.files)
    st.download_button(
        "Tải file kết quả",
        data=data,
        file_name=filename,
        mime=mime,
        type="primary",
        use_container_width=True,
    )
