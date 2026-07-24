# Excel Thuế Trợ Lý

PWA tiếng Việt xây dựng bằng Streamlit để xử lý các nghiệp vụ Excel theo quy tắc: đối chiếu hai cột, tìm dữ liệu trùng và đánh dấu ô trống. Gemini là tính năng tùy chọn, chỉ chuyển yêu cầu tiếng Việt thành kế hoạch có cấu trúc; dữ liệu thật trong ô được xử lý bằng `openpyxl` trên backend.

## Điểm nổi bật

- Giao diện thân thiện trên máy tính và điện thoại.
- Tải tối đa 2 file `.xlsx` hoặc `.xlsm` trong một lượt.
- Tự nhận diện sheet, dòng tiêu đề, tên cột và kiểu dữ liệu.
- Đối chiếu hai cột giữa hai file hoặc hai sheet.
- Tô màu dữ liệu trùng, không trùng, bị lặp hoặc ô trống.
- Chuẩn hóa mã số thuế bằng chế độ “chỉ giữ chữ số”.
- Không gửi giá trị trong ô cho Gemini; chỉ gửi metadata workbook.
- Người dùng luôn phải xác nhận kế hoạch AI trước khi chạy.
- Không ghi đè file gốc; kết quả được tạo thành file mới.
- Có manifest và giao diện standalone để cài lên màn hình chính. Phần xử lý Excel vẫn cần kết nối tới server.

## Kiến trúc an toàn

```text
Người dùng tải workbook
        ↓
Streamlit/openpyxl đọc cấu trúc cục bộ trên backend
        ↓
Gemini chỉ nhận tên file, sheet, cột, kiểu dữ liệu và số dòng
        ↓
Gemini trả về OperationPlan dạng JSON
        ↓
Người dùng kiểm tra và xác nhận
        ↓
openpyxl xử lý giá trị thật và xuất file mới
```

Gemini không được phép tạo hoặc thực thi Python. Backend chỉ chấp nhận ba hành động đã định nghĩa:

- `compare_columns`
- `highlight_duplicates`
- `find_blanks`

## Chạy trên máy cá nhân

Yêu cầu Python 3.11 trở lên.

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

Sau đó mở địa chỉ Streamlit hiển thị trong terminal, thường là `http://localhost:8501`.

## Cấu hình Gemini

Ứng dụng vẫn dùng được các công cụ thủ công khi không có API key.

Cách 1 — biến môi trường:

```bash
export GEMINI_API_KEY="your_api_key"
export GEMINI_MODEL="gemini-2.5-flash-lite"
```

Cách 2 — Streamlit secrets:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Sau đó sửa `.streamlit/secrets.toml`:

```toml
GEMINI_API_KEY = "your_api_key"
GEMINI_MODEL = "gemini-2.5-flash-lite"
```

Không commit API key lên GitHub. Không dùng cơ chế tự động xoay nhiều key để né giới hạn Free Tier.

## Ví dụ yêu cầu

```text
So sánh cột Mã số thuế ở file bán ra với MST khách hàng ở file kế toán,
nếu trùng thì tô vàng bên file bán ra.
```

```text
Tìm số hóa đơn bị trùng trong sheet HoaDon và tô đỏ.
```

```text
Đánh dấu các ô trống trong cột Mã số thuế người mua.
```

## Chạy kiểm thử

```bash
pip install -r requirements-dev.txt
pytest -q
```

Các test kiểm tra rằng payload gửi AI không chứa giá trị ô và thao tác đối chiếu tô đúng ô trong workbook kết quả.

## Chạy bằng Docker

```bash
docker build -t excel-thue-tro-ly .
docker run --rm -p 8501:8501 \
  -e GEMINI_API_KEY="your_api_key" \
  excel-thue-tro-ly
```

## Triển khai

Streamlit cần tiến trình Python chạy liên tục và kết nối WebSocket, vì vậy không nên triển khai trực tiếp như một Vercel Serverless Function. Các lựa chọn phù hợp hơn gồm Streamlit Community Cloud hoặc một nền tảng chạy container/Docker.

Khi triển khai, cấu hình secret `GEMINI_API_KEY` ở phần quản trị của nền tảng. Không đặt key trong mã nguồn hoặc biến frontend công khai.

## Giới hạn hiện tại

- Chỉ hỗ trợ `.xlsx` và `.xlsm`; chưa hỗ trợ `.xls` cũ.
- Việc lưu `.xlsm` cố gắng giữ VBA, nhưng vẫn nên kiểm tra lại macro và tính năng Excel nâng cao.
- Header được tự nhận diện trong 20 dòng đầu; workbook có bố cục đặc biệt có thể cần cải tiến thêm.
- PWA có thể cài lên màn hình chính, nhưng không xử lý workbook hoàn toàn ngoại tuyến.
- Kết quả phục vụ hỗ trợ nghiệp vụ bảng tính, không thay thế việc kiểm tra chuyên môn thuế hoặc pháp lý.

## Cấu trúc dự án

```text
app.py                    # Giao diện Streamlit
excel_ai/models.py        # Schema kế hoạch và dữ liệu
excel_ai/schema.py        # Kiểm tra file và đọc metadata
excel_ai/planner.py       # Gemini planner, không gửi giá trị ô
excel_ai/operations.py    # Engine openpyxl
static/                   # Manifest, icon, service worker
tests/                    # Kiểm thử bảo mật và xử lý Excel
```
