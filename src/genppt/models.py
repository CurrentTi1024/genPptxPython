from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .errors import TemplateError


@dataclass(frozen=True)
class ForeachDirective:
    item_path: str
    block_name: str
    body: str
    rows: int | None = None
    cols: int | None = None
    image_mode: str | None = None


@dataclass
class ShapePrototype:
    shape: Any
    directive: ForeachDirective
    element: Any


@dataclass
class RepeatBlock:
    item_path: str
    block_name: str
    prototypes: list[ShapePrototype]
    left: int
    top: int
    right: int
    bottom: int

    @property
    def table_prototype(self) -> ShapePrototype:
        tables = [prototype for prototype in self.prototypes if prototype.shape.has_table]
        if len(tables) != 1:
            raise TemplateError(
                f"block {self.block_name!r} must contain exactly one table, found {len(tables)}"
            )
        return tables[0]

    @property
    def non_table_prototypes(self) -> list[ShapePrototype]:
        return [prototype for prototype in self.prototypes if not prototype.shape.has_table]


@dataclass
class ImageBlock:
    item_path: str
    block_name: str
    prototypes: list[ShapePrototype]
    left: int
    top: int
    right: int
    bottom: int
    rows: int | None
    cols: int | None

    @property
    def picture_prototypes(self) -> list[ShapePrototype]:
        return [
            prototype
            for prototype in self.prototypes
            if hasattr(prototype.shape, "_pic") and hasattr(prototype.shape, "image")
        ]


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
