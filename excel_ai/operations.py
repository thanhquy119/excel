from __future__ import annotations

import re
import unicodedata
from collections import Counter
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .models import Action, ExecutionResult, ExecutionStats, NormalizationOptions, OperationPlan, WorkbookSchema
from .schema import build_column_index


class OperationError(ValueError):
    pass


def normalize_value(value: Any, options: NormalizationOptions) -> Any:
    if value is None:
        return None

    if options.mode == "exact":
        if isinstance(value, str) and options.trim_whitespace:
            value = value.strip()
        return (type(value).__name__, value)

    text = unicodedata.normalize("NFKC", str(value))
    if options.trim_whitespace:
        text = text.strip()
    if options.remove_internal_spaces:
        text = re.sub(r"\s+", "", text)
    if options.ignore_case:
        text = text.casefold()

    if text == "":
        return None
    if options.mode == "digits_only":
        digits = re.sub(r"\D", "", text)
        return digits or None
    if options.mode == "number":
        cleaned = text.replace(" ", "").replace(",", "")
        try:
            return Decimal(cleaned)
        except InvalidOperation:
            return None
    return text


def _fill(color: str) -> PatternFill:
    rgb = color.upper()
    if len(rgb) == 6:
        rgb = f"FF{rgb}"
    return PatternFill(fill_type="solid", fgColor=rgb)


def _load_editable(filename: str, data: bytes):
    keep_vba = Path(filename).suffix.lower() == ".xlsm"
    return load_workbook(BytesIO(data), data_only=False, keep_vba=keep_vba)


def _output_name(filename: str) -> str:
    path = Path(filename)
    return f"{path.stem}_da_xu_ly{path.suffix.lower()}"


def _iter_cells(ws, column_letter: str, header_row: int):
    for row_idx in range(header_row + 1, ws.max_row + 1):
        yield ws[f"{column_letter}{row_idx}"]


def _save_changed(workbooks: dict[str, Any], filenames: dict[str, str], changed_ids: set[str]) -> dict[str, bytes]:
    files: dict[str, bytes] = {}
    for file_id in changed_ids:
        buffer = BytesIO()
        workbooks[file_id].save(buffer)
        files[_output_name(filenames[file_id])] = buffer.getvalue()
    return files


def _validate_plan_columns(plan: OperationPlan, schemas: list[WorkbookSchema]):
    index = build_column_index(schemas)
    if plan.source_column_id not in index:
        raise OperationError("Cột nguồn không tồn tại trong workbook hiện tại.")
    if plan.comparison_column_id and plan.comparison_column_id not in index:
        raise OperationError("Cột đối chiếu không tồn tại trong workbook hiện tại.")
    return index


def execute_plan(
    plan: OperationPlan,
    schemas: list[WorkbookSchema],
    uploaded_files: dict[str, tuple[str, bytes]],
) -> ExecutionResult:
    if plan.needs_confirmation:
        raise OperationError(plan.question or "Kế hoạch cần được xác nhận trước khi chạy.")

    column_index = _validate_plan_columns(plan, schemas)
    filenames = {file_id: item[0] for file_id, item in uploaded_files.items()}
    workbooks = {
        file_id: _load_editable(filename, data)
        for file_id, (filename, data) in uploaded_files.items()
    }

    try:
        if plan.action == Action.COMPARE_COLUMNS:
            result = _compare_columns(plan, column_index, workbooks, filenames)
        elif plan.action == Action.HIGHLIGHT_DUPLICATES:
            result = _highlight_duplicates(plan, column_index, workbooks, filenames)
        elif plan.action == Action.FIND_BLANKS:
            result = _find_blanks(plan, column_index, workbooks, filenames)
        else:
            raise OperationError("Hành động chưa được hỗ trợ.")
        return result
    finally:
        for workbook in workbooks.values():
            workbook.close()


def _compare_columns(plan, column_index, workbooks, filenames) -> ExecutionResult:
    source_book, source_sheet, source_column = column_index[plan.source_column_id]
    compare_book, compare_sheet, compare_column = column_index[plan.comparison_column_id]
    source_ws = workbooks[source_book.file_id][source_sheet.name]
    compare_ws = workbooks[compare_book.file_id][compare_sheet.name]

    source_cells = list(_iter_cells(source_ws, source_column.letter, source_sheet.header_row))
    compare_cells = list(_iter_cells(compare_ws, compare_column.letter, compare_sheet.header_row))
    source_normalized = [normalize_value(cell.value, plan.normalization) for cell in source_cells]
    compare_normalized = [normalize_value(cell.value, plan.normalization) for cell in compare_cells]
    source_set = {value for value in source_normalized if value is not None}
    compare_set = {value for value in compare_normalized if value is not None}

    matched = sum(value is not None and value in compare_set for value in source_normalized)
    unmatched = sum(value is not None and value not in compare_set for value in source_normalized)
    highlighted = 0
    changed_ids: set[str] = set()
    cell_fill = _fill(plan.color)

    if plan.highlight_target in {"source", "both"}:
        for cell, value in zip(source_cells, source_normalized):
            is_match = value is not None and value in compare_set
            should_highlight = is_match if plan.highlight_condition == "matched" else value is not None and not is_match
            if should_highlight:
                cell.fill = cell_fill
                highlighted += 1
        changed_ids.add(source_book.file_id)

    if plan.highlight_target in {"comparison", "both"}:
        for cell, value in zip(compare_cells, compare_normalized):
            is_match = value is not None and value in source_set
            should_highlight = is_match if plan.highlight_condition == "matched" else value is not None and not is_match
            if should_highlight:
                cell.fill = cell_fill
                highlighted += 1
        changed_ids.add(compare_book.file_id)

    files = _save_changed(workbooks, filenames, changed_ids)
    return ExecutionResult(
        files=files,
        stats=ExecutionStats(
            scanned_rows=len(source_cells) + len(compare_cells),
            highlighted_cells=highlighted,
            matched_values=matched,
            unmatched_values=unmatched,
        ),
        message=f"Đã đối chiếu {len(source_cells):,} dòng nguồn với {len(compare_cells):,} dòng tham chiếu.",
    )


def _highlight_duplicates(plan, column_index, workbooks, filenames) -> ExecutionResult:
    source_book, source_sheet, source_column = column_index[plan.source_column_id]
    ws = workbooks[source_book.file_id][source_sheet.name]
    cells = list(_iter_cells(ws, source_column.letter, source_sheet.header_row))
    normalized = [normalize_value(cell.value, plan.normalization) for cell in cells]
    counts = Counter(value for value in normalized if value is not None)
    duplicate_keys = {value for value, count in counts.items() if count > 1}
    cell_fill = _fill(plan.color)
    highlighted = 0

    for cell, value in zip(cells, normalized):
        if value in duplicate_keys:
            cell.fill = cell_fill
            highlighted += 1

    files = _save_changed(workbooks, filenames, {source_book.file_id})
    return ExecutionResult(
        files=files,
        stats=ExecutionStats(
            scanned_rows=len(cells),
            highlighted_cells=highlighted,
            duplicate_values=len(duplicate_keys),
        ),
        message=f"Đã tìm thấy {len(duplicate_keys):,} giá trị bị trùng.",
    )


def _find_blanks(plan, column_index, workbooks, filenames) -> ExecutionResult:
    source_book, source_sheet, source_column = column_index[plan.source_column_id]
    ws = workbooks[source_book.file_id][source_sheet.name]
    cells = list(_iter_cells(ws, source_column.letter, source_sheet.header_row))
    cell_fill = _fill(plan.color)
    blanks = 0

    for cell in cells:
        if cell.value is None or str(cell.value).strip() == "":
            cell.fill = cell_fill
            blanks += 1

    files = _save_changed(workbooks, filenames, {source_book.file_id})
    return ExecutionResult(
        files=files,
        stats=ExecutionStats(
            scanned_rows=len(cells),
            highlighted_cells=blanks,
            blank_cells=blanks,
        ),
        message=f"Đã đánh dấu {blanks:,} ô trống.",
    )
