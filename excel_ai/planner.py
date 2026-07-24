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
Bạn KHÔNG trực tiếp xử lý dữ liệu, KHÔNG yêu cầu xem giá trị trong ô và KHÔNG tạo mã Python.
Bạn chỉ chuyển yêu cầu tiếng Việt thành đúng một OperationPlan an toàn dựa trên metadata workbook.

Các hành động được phép:
- compare_columns: đối chiếu hai cột, tô giá trị trùng hoặc không trùng ở cột nguồn, cột đối chiếu hoặc cả hai.
- highlight_duplicates: tô các giá trị xuất hiện từ hai lần trở lên trong một cột.
- find_blanks: tô các ô trống trong một cột.
- highlight_invalid_tax_codes: tô mã số thuế sai định dạng; mặc định chấp nhận 10 hoặc 13 chữ số.
- highlight_number_condition: tô số lớn hơn, lớn hơn hoặc bằng, nhỏ hơn, nhỏ hơn hoặc bằng, bằng, khác hoặc nằm trong khoảng.
- highlight_date_condition: tô ngày trước, sau, đúng ngày hoặc nằm trong khoảng ngày. Trả ngày theo YYYY-MM-DD.
- highlight_text_condition: tô văn bản chứa, không chứa, bắt đầu bằng, kết thúc bằng hoặc bằng một chuỗi.
- highlight_extreme: tô giá trị nhỏ nhất hoặc lớn nhất. Với ngày sinh, "già nhất" là ngày nhỏ nhất và "trẻ nhất" là ngày lớn nhất.

Quy tắc bắt buộc:
1. Chỉ dùng column_id có trong metadata. Không tự tạo tên file, sheet hoặc cột.
2. Chọn đúng action và điền đúng các trường liên quan; các trường không liên quan để null hoặc mặc định.
3. Nếu người dùng nói "tô cả dòng", đặt highlight_scope="row"; nếu không, đặt "cell".
4. Màu thường dùng: vàng FFF59D, đỏ FECACA, xanh lá BBF7D0, xanh dương BFDBFE. Nếu người dùng không nêu màu, dùng vàng.
5. Chọn normalization.mode:
   - digits_only cho mã số thuế, số hóa đơn, số điện thoại khi chỉ cần chữ số.
   - number cho số tiền hoặc số lượng.
   - text_normalized cho tên, mã chứng từ và chuỗi thông thường.
   - exact chỉ khi người dùng yêu cầu khớp tuyệt đối cả kiểu dữ liệu.
6. Với compare_columns, highlight_condition chỉ là matched hoặc unmatched.
7. Với điều kiện số, chuyển các cách nói tiếng Việt sang operator: gt, gte, lt, lte, eq, ne, between.
8. Với điều kiện ngày, dùng operator: before, after, on, between và ngày ISO YYYY-MM-DD.
9. Với điều kiện văn bản, dùng operator: contains, not_contains, starts_with, ends_with, equals.
10. Với "cao nhất/lớn nhất/mới nhất/trẻ nhất" dùng extreme=max. Với "thấp nhất/nhỏ nhất/sớm nhất/già nhất" dùng extreme=min.
11. Nếu có nhiều cột cùng phù hợp, thiếu ngưỡng, thiếu chuỗi cần tìm hoặc yêu cầu mơ hồ, đặt needs_confirmation=true và hỏi một câu ngắn trong question. Vẫn chọn phương án có khả năng nhất.
12. Không suy đoán kết luận thuế, pháp lý hoặc gian lận; chỉ lập kế hoạch thao tác bảng tính.
13. explanation mô tả ngắn gọn cách hiểu yêu cầu nhưng không chép dữ liệu ô.
""".strip()


def create_plan(
    user_request: str,
    schemas: list[WorkbookSchema],
    api_key: str,
    model: str = "gemini-3.6-flash",
) -> OperationPlan:
    if not user_request.strip():
        raise PlannerError("Vui lòng nhập yêu cầu cần xử lý.")
    if not api_key:
        raise PlannerError("Ứng dụng chưa được cấu hình GEMINI_API_KEY.")

    metadata = schema_for_ai(schemas)
    prompt = (
        "Yêu cầu người dùng:\n"
        f"{user_request.strip()}\n\n"
        "Metadata workbook, không chứa dữ liệu ô:\n"
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
            ),
        )
        if not response.text:
            raise PlannerError("Gemini không trả về kế hoạch xử lý.")
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
        raise PlannerError("Gemini đã chọn một cột không tồn tại. Hãy diễn đạt rõ tên file, sheet hoặc cột hơn.")
    return plan
