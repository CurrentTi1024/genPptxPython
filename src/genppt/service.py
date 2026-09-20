from __future__ import annotations

import os
import tempfile
from copy import deepcopy
from pathlib import Path
from typing import Any, Mapping, Sequence

from pptx import Presentation

from .config import GeneratorConfig
from .data import normalize_table
from .directives import render_expression, resolve_path
from .errors import TemplateError
from .image_renderer import process_images
from .models import RepeatBlock
from .planner import analyze_template, validate_plan
from .pptx_utils import (
    append_element,
    clone_static_slide,
    compile_blocks,
    remove_shape,
    remove_slide,
    replace_shape_text,
)
from .table_renderer import (
    PreparedTable,
    add_table_fragment,
    prepare_table,
    split_column_indices,
)
from .text_renderer import process_text_expressions


def _row_capacity(prepared: PreparedTable, available: int, row_offset: int) -> int:
    consumed = prepared.row_heights[0]
    capacity = 0
    while row_offset + capacity < len(prepared.rows):
        next_height = prepared.row_heights[1 + row_offset + capacity]
        if consumed + next_height > available:
            break
        consumed += next_height
        capacity += 1
    return capacity


def _fragment_plan(
    prepared: PreparedTable,
    *,
    cursor: int,
    region_top: int,
    region_bottom: int,
    table_relative_top: int,
    non_table_extent: int,
    minimum_fragment_rows: int,
) -> tuple[list[tuple[bool, int, int]], int]:
    full_region_height = region_bottom - region_top
    minimum_rows = min(minimum_fragment_rows, len(prepared.rows))
    minimum_table_height = prepared.row_heights[0] + sum(
        prepared.row_heights[1 : 1 + minimum_rows]
    )
    minimum_block_height = max(non_table_extent, table_relative_top + minimum_table_height)
    if minimum_block_height > full_region_height:
        raise TemplateError(
            "table header and minimum data rows cannot fit in the block region; "
            "increase the region or use shrink/truncate overflow"
        )

    full_table_height = sum(prepared.row_heights)
    full_block_height = max(non_table_extent, table_relative_top + full_table_height)
    plan: list[tuple[bool, int, int]] = []
    row_offset = 0
    new_page_before = False
    if full_block_height > region_bottom - cursor and cursor != region_top:
        cursor = region_top
        new_page_before = True

    while row_offset < len(prepared.rows):
        available_for_table = region_bottom - cursor - table_relative_top
        remaining = len(prepared.rows) - row_offset
        capacity = _row_capacity(prepared, available_for_table, row_offset)
        if capacity < min(minimum_fragment_rows, remaining) and cursor != region_top:
            cursor = region_top
            new_page_before = True
            available_for_table = region_bottom - cursor - table_relative_top
            capacity = _row_capacity(prepared, available_for_table, row_offset)
        if capacity == 0:
            raise TemplateError(
                f"table row {row_offset + 1} cannot fit in the block region"
            )
        if remaining - capacity == 1 and capacity > minimum_fragment_rows:
            capacity -= 1
        plan.append((new_page_before, row_offset, capacity))
        table_height = prepared.row_heights[0] + sum(
            prepared.row_heights[1 + row_offset : 1 + row_offset + capacity]
        )
        block_height = max(non_table_extent, table_relative_top + table_height)
        cursor += block_height
        row_offset += capacity
        new_page_before = row_offset < len(prepared.rows)
        if new_page_before:
            cursor = region_top
    return plan, cursor


def _unique_shape_name(
    base: str,
    block: RepeatBlock,
    item_index: int,
    column_fragment: int,
    row_fragment: int,
    kind: str,
) -> str:
    stem = base.strip() or kind
    suffix = (
        f"__{block.block_name}_i{item_index + 1}"
        f"_c{column_fragment + 1}_r{row_fragment + 1}_{kind}"
    )
    return f"{stem[: max(1, 255 - len(suffix))]}{suffix}"


def _render_block_items(
    presentation: Any,
    source_slide: Any,
    block: RepeatBlock,
    items: Sequence[Mapping[str, Any]],
    static_elements: Sequence[Any],
    config: GeneratorConfig,
    region_bottom: int,
    insert_after: Any,
    total_cells_before: int,
) -> tuple[Any, int]:
    region_top = block.top
    if region_bottom <= region_top:
        raise TemplateError(f"block {block.block_name!r} has no usable layout region")
    cursor = region_top
    current_slide = source_slide
    last_inserted_slide = insert_after
    table_prototype = block.table_prototype
    source_table = table_prototype.shape.table
    if any(cell.is_merge_origin or cell.is_spanned for cell in source_table.iter_cells()):
        raise TemplateError("merged cells are not supported in a dynamic table prototype")
    table_relative_top = table_prototype.shape.top - block.top
    non_table_extent = max(
        [0]
        + [
            prototype.shape.top - block.top + prototype.shape.height
            for prototype in block.non_table_prototypes
        ]
    )
    total_cells = total_cells_before

    def new_page() -> None:
        nonlocal current_slide, cursor, last_inserted_slide
        if len(presentation.slides) >= config.max_output_slides:
            raise TemplateError(f"output slide count exceeds limit {config.max_output_slides}")
        current_slide = clone_static_slide(
            presentation, source_slide, static_elements, last_inserted_slide
        )
        last_inserted_slide = current_slide
        cursor = region_top

    for item_index, item in enumerate(items):
        if not isinstance(item, Mapping):
            raise TemplateError(f"{block.item_path}[{item_index}] must be an object")
        normalized = normalize_table(item, config)
        if not normalized.rows:
            continue
        total_cells += len(normalized.labels) * (1 + len(normalized.rows))
        if total_cells > config.max_total_cells:
            raise TemplateError(f"total cell count exceeds limit {config.max_total_cells}")

        column_groups = split_column_indices(
            len(normalized.labels),
            table_prototype.shape.width,
            normalized.options.min_column_width_inches
            or config.min_column_width_inches,
            normalized.options.repeat_leading_columns,
        )
        for column_index, column_group in enumerate(column_groups):
            prepared = prepare_table(source_table, normalized, column_group)
            plan, final_cursor = _fragment_plan(
                prepared,
                cursor=cursor,
                region_top=region_top,
                region_bottom=region_bottom,
                table_relative_top=table_relative_top,
                non_table_extent=non_table_extent,
                minimum_fragment_rows=config.minimum_fragment_rows,
            )
            for fragment_index, (page_break, row_offset, row_count) in enumerate(plan):
                if page_break:
                    new_page()
                is_continuation = column_index > 0 or fragment_index > 0
                context = dict(item)
                context["_page"] = {
                    "isContinuation": is_continuation,
                    "continuationSuffix": "（续）" if is_continuation else "",
                    "fragmentIndex": fragment_index + 1,
                    "fragmentCount": len(plan),
                    "columnFragmentIndex": column_index + 1,
                    "columnFragmentCount": len(column_groups),
                }
                block_bottom = cursor
                if fragment_index == 0 or config.repeat_non_table_shapes_on_continuation:
                    for prototype_index, prototype in enumerate(block.non_table_prototypes):
                        generated = append_element(
                            current_slide, prototype.element, source_slide
                        )
                        generated.left = prototype.shape.left
                        generated.top = cursor + (prototype.shape.top - block.top)
                        base_name = render_expression(prototype.directive.body, context)
                        generated.name = _unique_shape_name(
                            base_name,
                            block,
                            item_index,
                            column_index,
                            fragment_index,
                            f"shape{prototype_index + 1}",
                        )
                        replace_shape_text(generated, context)
                        block_bottom = max(block_bottom, generated.top + generated.height)

                table_top = cursor + table_relative_top
                base_name = render_expression(table_prototype.directive.body, context)
                table_name = _unique_shape_name(
                    base_name,
                    block,
                    item_index,
                    column_index,
                    fragment_index,
                    "table",
                )
                _, table_height = add_table_fragment(
                    current_slide,
                    table_prototype,
                    prepared,
                    table_prototype.shape.left,
                    table_top,
                    row_offset,
                    row_count,
                    table_name,
                )
                block_bottom = max(block_bottom, table_top + table_height)
                cursor = block_bottom + config.block_gap
            cursor = final_cursor + config.block_gap
    return last_inserted_slide, total_cells


def _overlaps_horizontally(first: RepeatBlock, second: RepeatBlock) -> bool:
    return first.left < second.right and second.left < first.right


def _blocks_overlap(first: RepeatBlock, second: RepeatBlock) -> bool:
    return (
        _overlaps_horizontally(first, second)
        and first.top < second.bottom
        and second.top < first.bottom
    )


def _region_bottom(
    block: RepeatBlock,
    blocks: Sequence[RepeatBlock],
    page_bottom: int,
    gap: int,
) -> int:
    boundaries = [
        other.top - gap
        for other in blocks
        if other is not block
        and _overlaps_horizontally(block, other)
        and other.top >= block.bottom
    ]
    return min(boundaries, default=page_bottom)


def _save_atomically(presentation: Any, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}.", suffix=".tmp.pptx", dir=output_path.parent
    )
    os.close(descriptor)
    temporary_path = Path(temporary_name)
    try:
        presentation.save(str(temporary_path))
        os.replace(temporary_path, output_path)
    finally:
        temporary_path.unlink(missing_ok=True)


def _process_table_blocks(
    presentation: Any,
    data: Mapping[str, Any],
    config: GeneratorConfig,
) -> bool:
    """Expand all table blocks in an already loaded presentation."""
    matching_slides: list[tuple[Any, list[RepeatBlock], set[int]]] = []
    for slide in presentation.slides:
        blocks, dynamic_ids = compile_blocks(slide)
        if blocks:
            matching_slides.append((slide, blocks, dynamic_ids))
    if not matching_slides:
        return False
    if len(matching_slides) > 1 and not config.allow_additional_template_slides:
        raise TemplateError(
            "template contains foreach blocks on multiple slides; keep one template slide "
            "or explicitly enable allow_additional_template_slides"
        )

    source_slide, blocks, dynamic_ids = matching_slides[0]
    for unused_slide, _, _ in matching_slides[1:]:
        remove_slide(presentation, unused_slide)

    block_items: list[tuple[RepeatBlock, Sequence[Mapping[str, Any]]]] = []
    table_count = 0
    for block in blocks:
        items = resolve_path(data, block.item_path)
        if not isinstance(items, Sequence) or isinstance(items, (str, bytes)):
            raise TemplateError(f"{block.item_path!r} must resolve to an array")
        table_count += len(items)
        block_items.append((block, items))
    if table_count > config.max_tables:
        raise TemplateError(f"table count exceeds configured limit {config.max_tables}")

    for index, block in enumerate(blocks):
        for other in blocks[index + 1 :]:
            if _blocks_overlap(block, other):
                raise TemplateError(
                    f"block {block.block_name!r} overlaps block {other.block_name!r}"
                )

    static_elements = [
        deepcopy(shape._element)
        for shape in source_slide.shapes
        if shape.shape_id not in dynamic_ids
    ]
    for shape in list(source_slide.shapes):
        if shape.shape_id in dynamic_ids:
            remove_shape(shape)

    insert_after = source_slide
    total_cells = 0
    page_bottom = presentation.slide_height - config.bottom_margin
    for block, items in block_items:
        region_bottom = _region_bottom(
            block,
            blocks,
            page_bottom,
            config.block_gap,
        )
        insert_after, total_cells = _render_block_items(
            presentation,
            source_slide,
            block,
            items,
            static_elements,
            config,
            region_bottom,
            insert_after,
            total_cells,
        )
    return True


def generate_presentation(
    template_path: str | Path,
    data: Mapping[str, Any],
    output_path: str | Path,
    *,
    config: GeneratorConfig | None = None,
) -> Path:
    """Generate table and image expansions from a PPTX template."""
    config = config or GeneratorConfig()
    template_path = Path(template_path)
    output_path = Path(output_path)
    if not template_path.is_file():
        raise TemplateError(f"template does not exist: {template_path}")
    if template_path.resolve() == output_path.resolve():
        raise TemplateError("output path must differ from template path")
    if not isinstance(data, Mapping):
        raise TemplateError("root payload must be an object")

    presentation = Presentation(str(template_path))
    if len(presentation.slides) > config.max_output_slides:
        raise TemplateError(
            f"template slide count exceeds limit {config.max_output_slides}"
        )
    plan = analyze_template(presentation)
    validate_plan(plan)
    # Resolve root-level text before generating dynamic shapes. Continuation
    # pages clone the resolved static shapes, while dynamic business values
    # containing literal ``{{...}}`` are never interpreted a second time.
    process_text_expressions(presentation, data)
    if plan.has_images:
        process_images(presentation, data, config)
    if len(presentation.slides) > config.max_output_slides:
        raise TemplateError(f"output slide count exceeds limit {config.max_output_slides}")
    if plan.has_tables:
        _process_table_blocks(presentation, data, config)
    _save_atomically(presentation, output_path)
    return output_path
