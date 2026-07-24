from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field, model_validator


class Action(str, Enum):
    COMPARE_COLUMNS = "compare_columns"
    HIGHLIGHT_DUPLICATES = "highlight_duplicates"
    FIND_BLANKS = "find_blanks"


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
    highlight_condition: Literal["matched", "unmatched", "duplicates", "blank"]
    color: str = Field(default="FFF59D", pattern=r"^[0-9A-Fa-f]{6}$")
    normalization: NormalizationOptions = Field(default_factory=NormalizationOptions)
    explanation: str = ""
    confidence: float = Field(default=1.0, ge=0.0, le=1.0)
    needs_confirmation: bool = False
    question: str | None = None

    @model_validator(mode="after")
    def validate_action_fields(self) -> "OperationPlan":
        if self.action == Action.COMPARE_COLUMNS and not self.comparison_column_id:
            raise ValueError("compare_columns cần comparison_column_id")
        if self.action == Action.HIGHLIGHT_DUPLICATES:
            self.highlight_condition = "duplicates"
        if self.action == Action.FIND_BLANKS:
            self.highlight_condition = "blank"
        return self


class ExecutionStats(BaseModel):
    scanned_rows: int = 0
    highlighted_cells: int = 0
    matched_values: int = 0
    unmatched_values: int = 0
    duplicate_values: int = 0
    blank_cells: int = 0


class ExecutionResult(BaseModel):
    files: dict[str, bytes]
    stats: ExecutionStats
    message: str

    model_config = {"arbitrary_types_allowed": True}
