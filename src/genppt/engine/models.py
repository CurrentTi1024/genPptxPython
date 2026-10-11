from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class ForeachDirective:
    item_path: str
    block_name: str
    body: str
    rows: int | None = None
    cols: int | None = None
    image_mode: str | None = None
    page: str | None = None


@dataclass
class ShapePrototype:
    shape: Any
    directive: ForeachDirective
    element: Any


@dataclass(frozen=True)
class ColumnSpec:
    key: str
    label: str
    format_spec: str | None = None
    null_value: str | None = None


@dataclass(frozen=True)
class TableOptions:
    column_width: str = "template"
    min_column_width_inches: float | None = None
    repeat_leading_columns: int = 0
    overflow: str = "wrap"
    max_row_height_inches: float | None = None
    min_font_size_points: float | None = None
    null_value: str = ""


@dataclass(frozen=True)
class NormalizedTable:
    columns: list[ColumnSpec]
    labels: list[str]
    keys: list[str]
    rows: list[list[str]]
    options: TableOptions
