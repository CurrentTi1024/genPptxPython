from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .errors import TemplateError


@dataclass(frozen=True)
class ForeachDirective:
    item_path: str
    block_name: str
    body: str


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


@dataclass(frozen=True)
class NormalizedTable:
    labels: list[str]
    keys: list[str]
    rows: list[list[str]]
