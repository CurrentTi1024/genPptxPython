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
from .models import RepeatBlock
from .pptx_utils import (
    append_element,
    clone_static_slide,
    compile_blocks,
    remove_shape,
    remove_slide,
    replace_shape_text,
)
from .table_renderer import add_table_fragment, fragment_height, row_capacity


def _render_block_items(
    presentation: Any,
    source_slide: Any,
    block: RepeatBlock,
    items: Sequence[Mapping[str, Any]],
    static_elements: Sequence[Any],
    config: GeneratorConfig,
) -> None:
    page_top = block.top
    page_bottom = presentation.slide_height - config.bottom_margin
    if page_bottom <= page_top:
        raise TemplateError("inferred dynamic region has no usable height")
    cursor = page_top
    current_slide = source_slide
    table_prototype = block.table_prototype
    source_table = table_prototype.shape.table
    if any(cell.is_merge_origin or cell.is_spanned for cell in source_table.iter_cells()):
        raise TemplateError("merged cells are not supported in a dynamic table prototype")
    table_relative_top = table_prototype.shape.top - block.top
    total_cells = 0

    def new_page() -> None:
        nonlocal current_slide, cursor
        if len(presentation.slides) >= config.max_output_slides:
            raise TemplateError(f"output slide count exceeds limit {config.max_output_slides}")
        current_slide = clone_static_slide(
            presentation, source_slide, static_elements, current_slide
        )
        cursor = page_top

    for item_index, item in enumerate(items):
        if not isinstance(item, Mapping):
            raise TemplateError(f"{block.item_path}[{item_index}] must be an object")
        normalized = normalize_table(item, config)
        total_cells += len(normalized.labels) * (1 + len(normalized.rows))
        if total_cells > config.max_total_cells:
            raise TemplateError(f"total cell count exceeds limit {config.max_total_cells}")
        row_offset = 0
        first_fragment = True

        full_table_height = fragment_height(source_table, 0, len(normalized.rows))
        full_block_height = max(
            [table_relative_top + full_table_height]
            + [
                prototype.shape.top - block.top + prototype.shape.height
                for prototype in block.non_table_prototypes
            ]
        )
        if full_block_height > page_bottom - cursor and cursor != page_top:
            new_page()

        while first_fragment or row_offset < len(normalized.rows):
            remaining = len(normalized.rows) - row_offset
            available = page_bottom - cursor
            if full_block_height <= available and first_fragment:
                capacity = remaining
            else:
                capacity = row_capacity(
                    source_table, available, table_relative_top, row_offset, remaining
                )
            if remaining == 0:
                capacity = 0
                required = table_relative_top + fragment_height(source_table, 0, 0)
                if required > available and cursor != page_top:
                    new_page()
                    available = page_bottom - cursor
                if required > available:
                    raise TemplateError("table header cannot fit in the inferred dynamic region")
            elif capacity < min(config.minimum_fragment_rows, remaining) and cursor != page_top:
                new_page()
                available = page_bottom - cursor
                capacity = row_capacity(
                    source_table, available, table_relative_top, row_offset, remaining
                )
            if remaining > 0 and capacity == 0:
                raise TemplateError("a table row cannot fit in the inferred dynamic region")
            if remaining - capacity == 1 and capacity > config.minimum_fragment_rows:
                capacity -= 1

            block_bottom = cursor
            if first_fragment or config.repeat_non_table_shapes_on_continuation:
                for prototype in block.non_table_prototypes:
                    generated = append_element(current_slide, prototype.element, source_slide)
                    generated.left = prototype.shape.left
                    generated.top = cursor + (prototype.shape.top - block.top)
                    generated_name = render_expression(prototype.directive.body, item).strip()
                    generated.name = generated_name or (
                        f"{block.block_name}_shape_{item_index + 1}_{generated.shape_id}"
                    )
                    replace_shape_text(generated, dict(item))
                    block_bottom = max(block_bottom, generated.top + generated.height)

            fragment_rows = normalized.rows[row_offset : row_offset + capacity]
            table_top = cursor + table_relative_top
            _, table_height = add_table_fragment(
                current_slide,
                table_prototype,
                normalized,
                item,
                table_prototype.shape.left,
                table_top,
                row_offset,
                fragment_rows,
                item_index,
            )
            block_bottom = max(block_bottom, table_top + table_height)
            cursor = block_bottom + config.block_gap
            row_offset += capacity
            first_fragment = False
            if row_offset < len(normalized.rows):
                new_page()


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


def generate_presentation(
    template_path: str | Path,
    data: Mapping[str, Any],
    output_path: str | Path,
    *,
    config: GeneratorConfig | None = None,
) -> Path:
    """Generate a PPTX from one template slide and a table-array payload."""
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
    matching_slides: list[tuple[Any, list[RepeatBlock], set[int]]] = []
    for slide in presentation.slides:
        blocks, dynamic_ids = compile_blocks(slide)
        if blocks:
            matching_slides.append((slide, blocks, dynamic_ids))
    if not matching_slides:
        raise TemplateError("template contains no foreach table blocks")
    if len(matching_slides) > 1 and not config.allow_additional_template_slides:
        raise TemplateError(
            "template contains foreach blocks on multiple slides; keep one template slide "
            "or explicitly enable allow_additional_template_slides"
        )

    source_slide, blocks, dynamic_ids = matching_slides[0]
    if len(blocks) != 1:
        raise TemplateError("one template slide must contain exactly one repeated table block")
    for unused_slide, _, _ in matching_slides[1:]:
        remove_slide(presentation, unused_slide)

    block = blocks[0]
    items = resolve_path(data, block.item_path)
    if not isinstance(items, Sequence) or isinstance(items, (str, bytes)):
        raise TemplateError(f"{block.item_path!r} must resolve to an array")
    if len(items) > config.max_tables:
        raise TemplateError(f"table count exceeds configured limit {config.max_tables}")

    static_elements = [
        deepcopy(shape._element)
        for shape in source_slide.shapes
        if shape.shape_id not in dynamic_ids
    ]
    for shape in list(source_slide.shapes):
        if shape.shape_id in dynamic_ids:
            remove_shape(shape)
    _render_block_items(presentation, source_slide, block, items, static_elements, config)
    _save_atomically(presentation, output_path)
    return output_path
