from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal


@dataclass(frozen=True)
class CellPosition:
    sheet_name: str
    row: int
    column: int

    def __post_init__(self) -> None:
        if self.row < 1:
            raise ValueError("row must be >= 1")
        if self.column < 1:
            raise ValueError("column must be >= 1")

    def shifted(self, row_offset: int, column_offset: int) -> "CellPosition":
        return CellPosition(
            sheet_name=self.sheet_name,
            row=self.row + row_offset,
            column=self.column + column_offset,
        )


@dataclass(frozen=True)
class MappingRule:
    field_name: str
    label: str
    row_offset: int = 0
    column_offset: int = 0
    sheet_name: str | None = None
    required: bool = True
    aliases: tuple[str, ...] = ()
    value_type: str | None = None
    target_strategy: str = "offset_or_infer"
    confidence_threshold: float | None = None
    write_mode: str = "value"
    mark_symbol: str = "X"

    def __post_init__(self) -> None:
        if not self.field_name.strip():
            raise ValueError("field_name cannot be empty")
        if not self.label.strip():
            raise ValueError("label cannot be empty")
        if self.value_type is not None and not self.value_type.strip():
            raise ValueError("value_type cannot be empty when provided")
        normalized_write_mode = self.write_mode.strip().lower()
        if self.target_strategy not in {
            "offset",
            "infer",
            "offset_or_infer",
        }:
            raise ValueError("target_strategy must be one of offset|infer|offset_or_infer")
        if self.confidence_threshold is not None and not (
            0 <= self.confidence_threshold <= 1
        ):
            raise ValueError("confidence_threshold must be between 0 and 1")
        if normalized_write_mode not in {"value", "mark"}:
            raise ValueError("write_mode must be one of value|mark")
        if not isinstance(self.mark_symbol, str) or not self.mark_symbol.strip():
            raise ValueError("mark_symbol cannot be empty")
        object.__setattr__(self, "write_mode", normalized_write_mode)
        object.__setattr__(self, "mark_symbol", self.mark_symbol.strip())

        cleaned_aliases: list[str] = []
        for alias in self.aliases:
            if not isinstance(alias, str):
                raise ValueError("aliases must contain text values")
            stripped = alias.strip()
            if not stripped:
                continue
            cleaned_aliases.append(stripped)
        object.__setattr__(self, "aliases", tuple(cleaned_aliases))


@dataclass(frozen=True)
class FieldWriteTrace:
    field_name: str
    label_position: CellPosition
    target_position: CellPosition
    value: Any


@dataclass(frozen=True)
class FieldWriteResult:
    field_name: str
    status: Literal["written", "skipped", "error"]
    detail: str | None = None


@dataclass(frozen=True)
class FillFormResult:
    output_path: Path
    written_fields: list[FieldWriteTrace] = field(default_factory=list)
    write_results: list[FieldWriteResult] = field(default_factory=list)
    skipped_optional_fields: list[str] = field(default_factory=list)
    unmapped_input_fields: list[str] = field(default_factory=list)
