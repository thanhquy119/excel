# Excel Thuế Trợ Lý

Ứng dụng PWA tiếng Việt xây dựng bằng Streamlit để xử lý workbook Excel theo yêu cầu tự nhiên. Gemini chỉ chuyển yêu cầu thành kế hoạch có cấu trúc dựa trên tên file, sheet, cột, kiểu dữ liệu và số dòng; giá trị thật trong các ô được xử lý bằng `openpyxl` trên backend.

## Trải nghiệm sử dụng

Giao diện chỉ gồm ba thao tác chính:

- Tải tối đa 2 file `.xlsx` hoặc `.xlsm`.
- Mô tả yêu cầu bằng tiếng Việt.
- Bấm **Xử lý file** rồi tải workbook kết quả.

Không có chế độ thao tác thủ công, không hiển thị kế hoạch trung gian và không ghi đè file gốc.

## Yêu cầu được hỗ trợ

Backend chỉ thực thi các hành động đã được định nghĩa và kiểm tra trước:

- Đối chiếu hai cột, tô giá trị trùng hoặc không trùng.
- Tìm dữ liệu trùng trong một cột.
- Đánh dấu ô trống.
- Kiểm tra định dạng mã số thuế 10 hoặc 13 chữ số.
- Tô số tiền hoặc số lượng theo ngưỡng và khoảng giá trị.
- Tô ngày trước, sau, đúng ngày hoặc trong một khoảng ngày.
- Tô văn bản chứa, không chứa, bắt đầu bằng, kết thúc bằng hoặc bằng một chuỗi.
- Tìm giá trị lớn nhất hoặc nhỏ nhất, gồm người già nhất và trẻ nhất theo cột ngày sinh.
- Tô riêng ô hoặc toàn bộ dòng theo yêu cầu.

Ví dụ:

```text
So sánh cột Mã số thuế ở file bán ra với cột MST ở file kế toán,
nếu không trùng thì tô đỏ cả dòng trong file bán ra.
```

```text
Tìm người già nhất dựa vào cột Ngày sinh và tô vàng toàn bộ dòng.
```

```text
Tô xanh các dòng có Tổng tiền lớn hơn 100 triệu.
```

```text
Đánh dấu mã số thuế sai định dạng trong sheet Danh sách khách hàng.
```

```text
Tô các ô trong cột Nội dung có chứa cụm từ "tiếp khách".
```

## Kiến trúc an toàn

```text
Người dùng tải workbook
        ↓
Streamlit đọc metadata workbook
        ↓
Gemini 3.6 Flash tạo OperationPlan dạng JSON
        ↓
Backend kiểm tra action, cột và tham số
        ↓
openpyxl xử lý dữ liệu thật
        ↓
Xuất workbook mới để tải xuống
```

Gemini không được phép tạo hoặc thực thi Python. Không dùng `exec()`, không gửi giá trị ô trong prompt và không tự động xoay nhiều API key để né hạn mức.

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

Tạo `.streamlit/secrets.toml`:

```toml
GEMINI_API_KEY = "your_api_key"
GEMINI_MODEL = "gemini-3.6-flash"
```

Không commit file `secrets.toml` thật lên GitHub. File này đã được khai báo trong `.gitignore`.

## Deploy lên Streamlit Community Cloud

1. Đăng nhập Streamlit Community Cloud và chọn **Create app**.
2. Chọn repository `thanhquy119/excel`, branch `main`, file chạy `app.py`.
3. Mở **Advanced settings**.
4. Trong ô **Secrets**, dán:

```toml
GEMINI_API_KEY = "API_KEY_THAT_CUA_BAN"
GEMINI_MODEL = "gemini-3.6-flash"
```

5. Chọn Python 3.12 và deploy.

Secret được lưu trong phần cài đặt của ứng dụng, không nằm trong repository. Khi cần đổi key, vào workspace, mở menu của app, chọn **Settings**, mở tab **Secrets**, sửa giá trị và lưu.

## Chạy kiểm thử

```bash
pip install -r requirements-dev.txt
pytest -q
```

Các kiểm thử bao gồm khởi động giao diện, xác nhận model mặc định, bảo đảm metadata gửi AI không chứa giá trị ô, đối chiếu cột, kiểm tra mã số thuế và tìm người già nhất theo ngày sinh.

## Chạy bằng Docker

```bash
docker build -t excel-thue-tro-ly .
docker run --rm -p 8501:8501 \
  -e GEMINI_API_KEY="your_api_key" \
  -e GEMINI_MODEL="gemini-3.6-flash" \
  excel-thue-tro-ly
```

## Giới hạn

- Chỉ hỗ trợ `.xlsx` và `.xlsm`, chưa hỗ trợ `.xls` cũ.
- Header được tự nhận diện trong 20 dòng đầu.
- Workbook có công thức phức tạp, macro hoặc bố cục đặc biệt vẫn cần kiểm tra lại sau khi xử lý.
- PWA có thể cài lên màn hình chính nhưng việc xử lý cần kết nối tới server.
- Kết quả hỗ trợ nghiệp vụ bảng tính, không thay thế việc rà soát chuyên môn thuế hoặc pháp lý.
