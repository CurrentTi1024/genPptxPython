from __future__ import annotations

from copy import deepcopy
from typing import Any, Mapping

from .directives import render_expression
from .models import NormalizedTable, ShapePrototype


def source_column_index(source_count: int, target_index: int, target_count: int) -> int:
    if source_count <= 1 or target_count <= 1:
        return 0
    if target_index == 0:
        return 0
    if target_index == target_count - 1:
        return source_count - 1
    if source_count == 2:
        return 0
    ratio = target_index / (target_count - 1)
    return min(source_count - 2, max(1, round(ratio * (source_count - 1))))


def source_row_index(source_count: int, data_index: int | None) -> int:
    if data_index is None or source_count <= 1:
        return 0
    return 1 + (data_index % (source_count - 1))


def copy_table_properties(source_table: Any, target_table: Any) -> None:
    target_table._tbl.replace(target_table._tbl.tblPr, deepcopy(source_table._tbl.tblPr))


def copy_cell_style(source_cell: Any, target_cell: Any) -> None:
    if source_cell._tc.tcPr is not None and target_cell._tc.tcPr is not None:
        target_cell._tc.replace(target_cell._tc.tcPr, deepcopy(source_cell._tc.tcPr))
    target_cell.margin_left = source_cell.margin_left
    target_cell.margin_right = source_cell.margin_right
    target_cell.margin_top = source_cell.margin_top
    target_cell.margin_bottom = source_cell.margin_bottom
    target_cell.vertical_anchor = source_cell.vertical_anchor


def set_cell_text(target_cell: Any, value: str, source_cell: Any) -> None:
    target_cell.text = value
    source_paragraph = source_cell.text_frame.paragraphs[0]
    target_paragraph = target_cell.text_frame.paragraphs[0]
    target_paragraph.alignment = source_paragraph.alignment
    target_paragraph.level = source_paragraph.level
    if source_paragraph.runs and target_paragraph.runs:
        source_r_pr = source_paragraph.runs[0]._r.get_or_add_rPr()
        target_r_pr = target_paragraph.runs[0]._r.get_or_add_rPr()
        target_paragraph.runs[0]._r.replace(target_r_pr, deepcopy(source_r_pr))


def column_widths(source_table: Any, target_count: int) -> list[int]:
    total_width = sum(column.width for column in source_table.columns)
    source_widths = [column.width for column in source_table.columns]
    weights = [
        source_widths[source_column_index(len(source_widths), index, target_count)]
        for index in range(target_count)
    ]
    weight_sum = sum(weights)
    widths = [max(1, round(total_width * weight / weight_sum)) for weight in weights]
    widths[-1] += total_width - sum(widths)
    return widths


def row_height(source_table: Any, global_data_index: int | None) -> int:
    return source_table.rows[source_row_index(len(source_table.rows), global_data_index)].height


def fragment_height(source_table: Any, row_offset: int, row_count: int) -> int:
    return row_height(source_table, None) + sum(
        row_height(source_table, row_offset + index) for index in range(row_count)
    )


def row_capacity(
    source_table: Any,
    available_height: int,
    table_relative_top: int,
    row_offset: int,
    remaining: int,
) -> int:
    consumed = table_relative_top + row_height(source_table, None)
    capacity = 0
    while capacity < remaining:
        next_height = row_height(source_table, row_offset + capacity)
        if consumed + next_height > available_height:
            break
        consumed += next_height
        capacity += 1
    return capacity


def add_table_fragment(
    slide: Any,
    prototype: ShapePrototype,
    normalized: NormalizedTable,
    context: Mapping[str, Any],
    left: int,
    top: int,
    row_offset: int,
    fragment_rows: list[list[str]],
    item_index: int,
) -> tuple[Any, int]:
    source_shape = prototype.shape
    source_table = source_shape.table
    row_heights = [row_height(source_table, None)] + [
        row_height(source_table, row_offset + index) for index in range(len(fragment_rows))
    ]
    table_height = sum(row_heights)
    table_shape = slide.shapes.add_table(
        1 + len(fragment_rows),
        len(normalized.labels),
        left,
        top,
        source_shape.width,
        table_height,
    )
    generated_name = render_expression(prototype.directive.body, context).strip()
    table_shape.name = generated_name or f"{prototype.directive.block_name}_table_{item_index + 1}"
    target_table = table_shape.table
    copy_table_properties(source_table, target_table)
    for index, width in enumerate(column_widths(source_table, len(normalized.labels))):
        target_table.columns[index].width = width
    for index, height in enumerate(row_heights):
        target_table.rows[index].height = height

    for target_row_index, row_values in enumerate([normalized.labels, *fragment_rows]):
        global_data_index = None if target_row_index == 0 else row_offset + target_row_index - 1
        source_row = source_row_index(len(source_table.rows), global_data_index)
        for target_col_index, value in enumerate(row_values):
            source_col = source_column_index(
                len(source_table.columns), target_col_index, len(normalized.labels)
            )
            source_cell = source_table.cell(source_row, source_col)
            target_cell = target_table.cell(target_row_index, target_col_index)
            copy_cell_style(source_cell, target_cell)
            set_cell_text(target_cell, value, source_cell)
    return table_shape, table_height
