from __future__ import annotations

import json

from google import genai
from google.genai import types
from pydantic import ValidationError

from .models import OperationPlan, WorkbookSchema
from .schema import build_column_index, schema_for_ai


class PlannerError(RuntimeError):
    pass


SYSTEM_INSTRUCTION = """
Bạn là bộ lập kế hoạch cho ứng dụng xử lý Excel nghiệp vụ kế toán và thuế.
Bạn KHÔNG trực tiếp xử lý dữ liệu và không được yêu cầu xem giá trị trong các ô.
Bạn chỉ ánh xạ yêu cầu tiếng Việt sang đúng hành động, file, sheet và column_id có trong metadata.

Hành động được phép:
- compare_columns: đối chiếu hai cột, tô màu dòng trùng hoặc không trùng.
- highlight_duplicates: tìm giá trị lặp lại trong một cột.
- find_blanks: đánh dấu ô trống trong một cột.

Quy tắc bắt buộc:
1. Chỉ dùng column_id xuất hiện trong metadata.
2. Không tạo mã Python, công thức Excel hay hành động ngoài danh sách.
3. Nếu có từ hai cột trở lên cùng phù hợp hoặc yêu cầu chưa rõ, đặt needs_confirmation=true,
   nêu câu hỏi ngắn trong question và vẫn chọn phương án có khả năng nhất.
4. Chọn normalization.mode:
   - digits_only cho mã số thuế, số hóa đơn chỉ cần chữ số, số điện thoại.
   - number cho số tiền hoặc số lượng.
   - text_normalized cho tên, mã chứng từ, chuỗi thông thường.
   - exact chỉ khi người dùng yêu cầu khớp tuyệt đối cả kiểu dữ liệu.
5. Không suy đoán kết luận thuế hoặc pháp lý; chỉ lập kế hoạch thao tác bảng tính.
6. explanation phải mô tả ngắn gọn cách hệ thống hiểu yêu cầu để người dùng xác nhận.
""".strip()


def create_plan(
    user_request: str,
    schemas: list[WorkbookSchema],
    api_key: str,
    model: str = "gemini-2.5-flash-lite",
) -> OperationPlan:
    if not user_request.strip():
        raise PlannerError("Vui lòng nhập yêu cầu cần xử lý.")
    if not api_key:
        raise PlannerError("Chưa cấu hình GEMINI_API_KEY.")

    metadata = schema_for_ai(schemas)
    prompt = (
        "Yêu cầu người dùng:\n"
        f"{user_request.strip()}\n\n"
        "Metadata workbook (không chứa dữ liệu ô):\n"
        f"{json.dumps(metadata, ensure_ascii=False, indent=2)}"
    )

    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=model,
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=SYSTEM_INSTRUCTION,
                response_mime_type="application/json",
                response_schema=OperationPlan,
                temperature=0.1,
            ),
        )
        if not response.text:
            raise PlannerError("Gemini không trả về kế hoạch.")
        plan = OperationPlan.model_validate_json(response.text)
    except ValidationError as exc:
        raise PlannerError("Gemini trả về kế hoạch không đúng định dạng.") from exc
    except PlannerError:
        raise
    except Exception as exc:
        raise PlannerError(f"Không thể kết nối Gemini: {exc}") from exc

    valid_columns = build_column_index(schemas)
    referenced = [plan.source_column_id, plan.comparison_column_id]
    invalid = [column_id for column_id in referenced if column_id and column_id not in valid_columns]
    if invalid:
        raise PlannerError("Gemini đã chọn một cột không tồn tại. Vui lòng thử diễn đạt rõ hơn.")
    return plan
