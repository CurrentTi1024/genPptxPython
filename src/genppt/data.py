from __future__ import annotations

import math
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Mapping, Sequence

from .config import GeneratorConfig
from .errors import TemplateError
from .models import NormalizedTable


def _display_value(value: Any, max_length: int) -> str:
    if value is None:
        result = ""
    elif isinstance(value, (str, int, bool, Decimal, date, datetime)):
        result = str(value)
    elif isinstance(value, float):
        if not math.isfinite(value):
            raise TemplateError("NaN and infinite cell values are not supported")
        result = str(value)
    else:
        raise TemplateError(f"unsupported cell value type: {type(value).__name__}")
    if len(result) > max_length:
        raise TemplateError(f"cell text exceeds {max_length} characters")
    return result


def normalize_table(item: Mapping[str, Any], config: GeneratorConfig) -> NormalizedTable:
    if "columns" not in item or "rows" not in item:
        raise TemplateError("each table item must contain columns and rows")
    columns = item["columns"]
    rows = item["rows"]
    if not isinstance(columns, Sequence) or isinstance(columns, (str, bytes)) or not columns:
        raise TemplateError("columns must be a non-empty array")
    if len(columns) > config.max_columns:
        raise TemplateError(f"column count exceeds configured limit {config.max_columns}")
    if not isinstance(rows, Sequence) or isinstance(rows, (str, bytes)):
        raise TemplateError("rows must be an array")
    if len(rows) > config.max_rows_per_table:
        raise TemplateError(f"row count exceeds configured limit {config.max_rows_per_table}")

    labels: list[str] = []
    keys: list[str] = []
    for index, column in enumerate(columns):
        if isinstance(column, Mapping):
            key = str(column.get("key", column.get("label", index)))
            label = str(column.get("label", key))
        else:
            key = str(column)
            label = str(column)
        keys.append(key)
        labels.append(_display_value(label, config.max_cell_text_length))
    if len(set(keys)) != len(keys):
        raise TemplateError("column keys must be unique")

    normalized_rows: list[list[str]] = []
    for row_index, row in enumerate(rows):
        if isinstance(row, Mapping):
            values = [row.get(key, "") for key in keys]
        elif isinstance(row, Sequence) and not isinstance(row, (str, bytes)):
            if len(row) != len(columns):
                raise TemplateError(
                    f"row {row_index} has {len(row)} values but columns has {len(columns)} entries"
                )
            values = list(row)
        else:
            raise TemplateError(f"row {row_index} must be an array or object")
        normalized_rows.append(
            [_display_value(value, config.max_cell_text_length) for value in values]
        )
    return NormalizedTable(labels, keys, normalized_rows)
