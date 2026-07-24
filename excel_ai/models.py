from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field, model_validator


class Action(str, Enum):
    COMPARE_COLUMNS = "compare_columns"
    HIGHLIGHT_DUPLICATES = "highlight_duplicates"
    FIND_BLANKS = "find_blanks"
    HIGHLIGHT_INVALID_TAX_CODES = "highlight_invalid_tax_codes"
    HIGHLIGHT_NUMBER_CONDITION = "highlight_number_condition"
    HIGHLIGHT_DATE_CONDITION = "highlight_date_condition"
    HIGHLIGHT_TEXT_CONDITION = "highlight_text_condition"
    HIGHLIGHT_EXTREME = "highlight_extreme"


class ColumnSchema(BaseModel):
    id: str
    letter: str
    header: str
    detected_type: Literal["text", "number", "date", "mixed", "empty"]


class SheetSchema(BaseModel):
    id: str
    name: str
    header_row: int
    max_row: int
    max_column: int
    columns: list[ColumnSchema]


class WorkbookSchema(BaseModel):
    file_id: str
    display_name: str
    sheets: list[SheetSchema]


class NormalizationOptions(BaseModel):
    mode: Literal["exact", "text_normalized", "digits_only", "number"] = "text_normalized"
    ignore_case: bool = True
    trim_whitespace: bool = True
    remove_internal_spaces: bool = False


class OperationPlan(BaseModel):
    action: Action
    source_column_id: str
    comparison_column_id: str | None = None
    highlight_target: Literal["source", "comparison", "both"] = "source"
    highlight_scope: Literal["cell", "row"] = "cell"
    highlight_condition: Literal[
        "matched",
        "unmatched",
        "duplicates",
        "blank",
        "invalid",
        "condition",
        "extreme",
    ] = "condition"
    color: str = Field(default="FFF59D", pattern=r"^[0-9A-Fa-f]{6}$")
    normalization: NormalizationOptions = Field(default_factory=NormalizationOptions)

    number_operator: Literal["gt", "gte", "lt", "lte", "eq", "ne", "between"] | None = None
    number_value: float | None = None
    second_number_value: float | None = None

    date_operator: Literal["before", "after", "on", "between"] | None = None
    date_value: str | None = None
    second_date_value: str | None = None

    text_operator: Literal["contains", "not_contains", "starts_with", "ends_with", "equals"] | None = None
    text_value: str | None = None

    tax_code_lengths: list[int] = Field(default_factory=lambda: [10, 13])
    extreme: Literal["min", "max"] | None = None

    explanation: str = ""
    confidence: float = Field(default=1.0, ge=0.0, le=1.0)
    needs_confirmation: bool = False
    question: str | None = None

    @model_validator(mode="after")
    def validate_action_fields(self) -> "OperationPlan":
        if self.action == Action.COMPARE_COLUMNS:
            if not self.comparison_column_id:
                raise ValueError("compare_columns cần comparison_column_id")
            if self.highlight_condition not in {"matched", "unmatched"}:
                raise ValueError("compare_columns chỉ hỗ trợ matched hoặc unmatched")

        if self.action == Action.HIGHLIGHT_DUPLICATES:
            self.highlight_condition = "duplicates"

        if self.action == Action.FIND_BLANKS:
            self.highlight_condition = "blank"

        if self.action == Action.HIGHLIGHT_INVALID_TAX_CODES:
            self.highlight_condition = "invalid"
            if not self.tax_code_lengths or any(length <= 0 for length in self.tax_code_lengths):
                raise ValueError("tax_code_lengths không hợp lệ")

        if self.action == Action.HIGHLIGHT_NUMBER_CONDITION:
            self.highlight_condition = "condition"
            if not self.number_operator or self.number_value is None:
                raise ValueError("Điều kiện số cần number_operator và number_value")
            if self.number_operator == "between" and self.second_number_value is None:
                raise ValueError("Điều kiện between cần second_number_value")

        if self.action == Action.HIGHLIGHT_DATE_CONDITION:
            self.highlight_condition = "condition"
            if not self.date_operator or not self.date_value:
                raise ValueError("Điều kiện ngày cần date_operator và date_value")
            if self.date_operator == "between" and not self.second_date_value:
                raise ValueError("Điều kiện between cần second_date_value")

        if self.action == Action.HIGHLIGHT_TEXT_CONDITION:
            self.highlight_condition = "condition"
            if not self.text_operator or self.text_value is None:
                raise ValueError("Điều kiện văn bản cần text_operator và text_value")

        if self.action == Action.HIGHLIGHT_EXTREME:
            self.highlight_condition = "extreme"
            if not self.extreme:
                raise ValueError("highlight_extreme cần extreme=min hoặc max")

        return self


class ExecutionStats(BaseModel):
    scanned_rows: int = 0
    highlighted_cells: int = 0
    matched_values: int = 0
    unmatched_values: int = 0
    duplicate_values: int = 0
    blank_cells: int = 0
    invalid_values: int = 0
    conditional_values: int = 0
    extreme_values: int = 0


class ExecutionResult(BaseModel):
    files: dict[str, bytes]
    stats: ExecutionStats
    message: str

    model_config = {"arbitrary_types_allowed": True}
