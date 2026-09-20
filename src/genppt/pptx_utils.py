from __future__ import annotations

from copy import deepcopy
from typing import Any, Sequence

from .directives import parse_foreach
from .errors import TemplateError
from .models import RepeatBlock, ShapePrototype
from .text_renderer import replace_text_frame_expressions


def compile_blocks(slide: Any) -> tuple[list[RepeatBlock], set[int]]:
    grouped: dict[tuple[str, str], list[ShapePrototype]] = {}
    for shape in slide.shapes:
        directive = parse_foreach(shape.name, shape.shape_id)
        if directive is None:
            continue
        grouped.setdefault((directive.item_path, directive.block_name), []).append(
            ShapePrototype(shape, directive, deepcopy(shape._element))
        )
    blocks: list[RepeatBlock] = []
    dynamic_ids: set[int] = set()
    for (item_path, block_name), prototypes in grouped.items():
        table_count = sum(1 for item in prototypes if item.shape.has_table)
        if table_count == 0:
            continue
        if table_count > 1:
            raise TemplateError(
                f"block {block_name!r} must contain one table, found {table_count}"
            )
        block = RepeatBlock(
            item_path,
            block_name,
            prototypes,
            min(item.shape.left for item in prototypes),
            min(item.shape.top for item in prototypes),
            max(item.shape.left + item.shape.width for item in prototypes),
            max(item.shape.top + item.shape.height for item in prototypes),
        )
        _ = block.table_prototype
        blocks.append(block)
        dynamic_ids.update(item.shape.shape_id for item in prototypes)
    blocks.sort(key=lambda block: (block.top, block.left))
    return blocks, dynamic_ids


def remove_shape(shape: Any) -> None:
    element = shape._element
    element.getparent().remove(element)


def _relationship_ids(element: Any) -> set[str]:
    return {
        value
        for node in element.iter()
        for value in node.attrib.values()
        if isinstance(value, str) and value.startswith("rId")
    }


def _remap_relationships(element: Any, source_slide: Any, target_slide: Any) -> None:
    mapping: dict[str, str] = {}
    for old_rid in _relationship_ids(element):
        if old_rid not in source_slide.part.rels:
            continue
        rel = source_slide.part.rels[old_rid]
        if rel.reltype.endswith("/slideLayout") or rel.reltype.endswith("/notesSlide"):
            continue
        target = rel.target_ref if rel.is_external else rel.target_part
        mapping[old_rid] = target_slide.part.relate_to(target, rel.reltype, rel.is_external)
    for node in element.iter():
        for attr, value in list(node.attrib.items()):
            if value in mapping:
                node.set(attr, mapping[value])


def append_element(target_slide: Any, element: Any, source_slide: Any) -> Any:
    clone = deepcopy(element)
    _remap_relationships(clone, source_slide, target_slide)
    c_nv_pr = clone.xpath(".//p:cNvPr")
    next_shape_id = target_slide.shapes._next_shape_id
    for offset, node in enumerate(c_nv_pr):
        node.set("id", str(next_shape_id + offset))
    target_slide.shapes._spTree.insert_element_before(clone, "p:extLst")
    return target_slide.shapes[-1]


def replace_shape_text(shape: Any, context: dict[str, Any]) -> None:
    if not shape.has_text_frame:
        return
    replace_text_frame_expressions(shape.text_frame, context)


def clone_static_slide(
    presentation: Any,
    source_slide: Any,
    static_elements: Sequence[Any],
    insert_after: Any,
) -> Any:
    target_slide = presentation.slides.add_slide(source_slide.slide_layout)
    for shape in list(target_slide.shapes):
        remove_shape(shape)
    for element in static_elements:
        append_element(target_slide, element, source_slide)
    slide_ids = presentation.slides._sldIdLst
    new_id = slide_ids[-1]
    slide_ids.remove(new_id)
    source_index = next(
        index for index, slide in enumerate(presentation.slides) if slide == insert_after
    )
    slide_ids.insert(source_index + 1, new_id)
    return target_slide


def remove_slide(presentation: Any, slide: Any) -> None:
    slide_id_list = presentation.slides._sldIdLst
    for slide_id in list(slide_id_list):
        if presentation.part.related_part(slide_id.rId) is slide.part:
            slide_id_list.remove(slide_id)
            presentation.part.drop_rel(slide_id.rId)
            return
    raise TemplateError("could not remove an unused template slide")
