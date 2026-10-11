from __future__ import annotations

import math
import re
from collections.abc import Mapping, Sequence
from datetime import date, datetime
from decimal import Decimal
from typing import Any

from .config import GeneratorConfig
from .errors import TemplateError
from .models import ColumnSpec, NormalizedTable, TableOptions


def _display_value(
    value: Any,
    max_length: int,
    *,
    format_spec: str | None = None,
    null_value: str = "",
) -> str:
    if isinstance(value, Decimal) and not value.is_finite():
        raise TemplateError("NaN and infinite cell values are not supported")
    if format_spec and any(int(n) > max_length for n in re.findall(r"\d+", format_spec)):
        raise TemplateError("format width/precision exceeds cell text limit")
    if value is None:
        result = null_value
    elif isinstance(value, (str, int, bool, Decimal, date, datetime)):
        if format_spec and isinstance(value, str) and "%" in format_spec:
            try:
                parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
            except ValueError:
                try:
                    parsed = date.fromisoformat(value)
                except ValueError as exc:
                    raise TemplateError(
                        f"value {value!r} is not an ISO date for format {format_spec!r}"
                    ) from exc
            result = parsed.strftime(format_spec)
        elif format_spec and isinstance(value, (date, datetime)):
            result = value.strftime(format_spec)
        elif format_spec:
            try:
                result = format(value, format_spec)
            except (TypeError, ValueError) as exc:
                raise TemplateError(f"value {value!r} cannot use format {format_spec!r}") from exc
        else:
            result = str(value)
    elif isinstance(value, float):
        if not math.isfinite(value):
            raise TemplateError("NaN and infinite cell values are not supported")
        if format_spec:
            try:
                result = format(value, format_spec)
            except ValueError as exc:
                raise TemplateError(f"value {value!r} cannot use format {format_spec!r}") from exc
        else:
            result = str(value)
    else:
        raise TemplateError(f"unsupported cell value type: {type(value).__name__}")
    if len(result) > max_length:
        raise TemplateError(f"cell text exceeds {max_length} characters")
    return result


def _table_options(item: Mapping[str, Any], config: GeneratorConfig) -> TableOptions:
    raw = item.get("_table", {})
    if not isinstance(raw, Mapping):
        raise TemplateError("_table must be an object")
    allowed = {
        "columnWidth",
        "minColumnWidth",
        "repeatLeadingColumns",
        "overflow",
        "maxRowHeight",
        "minFontSize",
        "nullValue",
        "formats",
        "nullValues",
    }
    unknown = set(raw) - allowed
    if unknown:
        raise TemplateError(f"unsupported _table options: {', '.join(sorted(unknown))}")
    strategy = str(raw.get("columnWidth", "template"))
    if strategy not in {"template", "content", "equal"}:
        raise TemplateError("columnWidth must be template, content, or equal")
    overflow = str(raw.get("overflow", config.default_overflow))
    if overflow not in {"wrap", "shrink", "truncate", "error"}:
        raise TemplateError("overflow must be wrap, shrink, truncate, or error")
    try:
        min_width = float(raw.get("minColumnWidth", config.min_column_width_inches))
        max_height = float(raw.get("maxRowHeight", config.max_row_height_inches))
        min_font = float(raw.get("minFontSize", config.min_font_size_points))
        repeat = raw.get("repeatLeadingColumns", 0)
    except (TypeError, ValueError, OverflowError) as exc:
        raise TemplateError("table size options must be numeric") from exc
    if isinstance(repeat, bool) or not isinstance(repeat, int):
        raise TemplateError("repeatLeadingColumns must be a non-negative integer")
    if not all(math.isfinite(value) for value in (min_width, max_height, min_font)):
        raise TemplateError("table size options must be finite")
    if min_width <= 0 or max_height <= 0 or min_font <= 0 or repeat < 0:
        raise TemplateError("table size options must be positive and repeat count non-negative")
    null_value = raw.get("nullValue", "")
    if null_value is None:
        null_value = ""
    if not isinstance(null_value, str):
        raise TemplateError("nullValue must be a string")
    return TableOptions(strategy, min_width, repeat, overflow, max_height, min_font, null_value)


def normalize_table(item: Mapping[str, Any], config: GeneratorConfig) -> NormalizedTable:
    rows = item.get("rows")
    if rows is None:
        rows = []
    if not isinstance(rows, Sequence) or isinstance(rows, (str, bytes)):
        raise TemplateError("rows must be an array")
    if not rows:
        return NormalizedTable([], [], [], [], _table_options(item, config))
    columns = item.get("columns")
    if columns is None:
        if rows and isinstance(rows[0], Mapping):
            columns = list(rows[0])
        else:
            raise TemplateError("columns is required for matrix rows")
    options = _table_options(item, config)
    if not isinstance(columns, Sequence) or isinstance(columns, (str, bytes)) or not columns:
        raise TemplateError("columns must be a non-empty array")
    if len(columns) > config.max_columns:
        raise TemplateError(f"column count exceeds configured limit {config.max_columns}")
    if not isinstance(rows, Sequence) or isinstance(rows, (str, bytes)):
        raise TemplateError("rows must be an array")
    if len(rows) > config.max_rows_per_table:
        raise TemplateError(f"row count exceeds configured limit {config.max_rows_per_table}")

    columns_specs: list[ColumnSpec] = []
    formats = item.get("_table", {}).get("formats", {})
    null_values = item.get("_table", {}).get("nullValues", {})
    if not isinstance(formats, Mapping) or not isinstance(null_values, Mapping):
        raise TemplateError("formats and nullValues must be field mappings")
    for column in columns:
        if isinstance(column, Mapping):
            if len(column) != 1:
                raise TemplateError("each column mapping must contain one field and label")
            key, label = next(iter(column.items()))
            if not isinstance(key, str) or not isinstance(label, str):
                raise TemplateError("column fields and labels must be strings")
        else:
            if not isinstance(column, str):
                raise TemplateError("columns must contain strings or one-field mappings")
            key = label = column
        format_spec = formats.get(key)
        null_value = null_values.get(key)
        if format_spec is not None and not isinstance(format_spec, str):
            raise TemplateError("column format must be a string")
        if null_value is not None and not isinstance(null_value, str):
            raise TemplateError("column nullValue must be a string")
        columns_specs.append(ColumnSpec(key, label, format_spec, null_value))
    keys = [column.key for column in columns_specs]
    labels = [_display_value(column.label, config.max_cell_text_length) for column in columns_specs]
    if len(set(keys)) != len(keys):
        raise TemplateError("column keys must be unique")

    normalized_rows: list[list[str]] = []
    for row_index, row in enumerate(rows):
        if isinstance(row, Mapping):
            values = [row.get(key) for key in keys]
        elif isinstance(row, Sequence) and not isinstance(row, (str, bytes)):
            if len(row) != len(columns):
                raise TemplateError(
                    f"row {row_index} has {len(row)} values but columns has {len(columns)} entries"
                )
            values = list(row)
        else:
            raise TemplateError(f"row {row_index} must be an array or object")
        normalized_rows.append(
            [
                _display_value(
                    value,
                    config.max_cell_text_length,
                    format_spec=column.format_spec,
                    null_value=(
                        column.null_value if column.null_value is not None else options.null_value
                    ),
                )
                for column, value in zip(columns_specs, values, strict=True)
            ]
        )
    return NormalizedTable(columns_specs, labels, keys, normalized_rows, options)
