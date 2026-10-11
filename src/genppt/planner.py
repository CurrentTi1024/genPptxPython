from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any

from .directives import EXPRESSION_RE, parse_foreach, parse_single_image
from .errors import TemplateError
from .models import ForeachDirective, ShapePrototype
from .text_renderer import validate_text_frame


@dataclass
class PageBlock:
    item_path: str
    block_name: str
    prototypes: list[ShapePrototype]
    root_text: bool = False

    @property
    def left(self) -> int:
        return min(p.shape.left for p in self.prototypes)

    @property
    def right(self) -> int:
        return max(p.shape.left + p.shape.width for p in self.prototypes)

    @property
    def top(self) -> int:
        return min(p.shape.top for p in self.prototypes)

    @property
    def bottom(self) -> int:
        return max(p.shape.top + p.shape.height for p in self.prototypes)

    @property
    def page(self) -> str:
        return next((p.directive.page for p in self.prototypes if p.directive.page), "flow")


@dataclass
class PagePlan:
    slide: Any
    source_index: int
    blocks: list[PageBlock]
    dynamic_ids: set[int]


def _validate_shape(shape: Any) -> None:
    if shape.has_text_frame:
        validate_text_frame(shape.text_frame)
    if shape.has_table:
        for cell in shape.table.iter_cells():
            validate_text_frame(cell.text_frame)
    if hasattr(shape, "shapes"):
        for child in shape.shapes:
            if "<foreach" in child.name:
                raise TemplateError(
                    "loop directives must be on top-level Shapes, not group children"
                )
            _validate_shape(child)


def analyze_template(presentation: Any) -> list[PagePlan]:
    plans = []
    for index, slide in enumerate(list(presentation.slides)):
        grouped: dict[tuple[str, str], PageBlock] = {}
        dynamic_ids = set()
        for shape in slide.shapes:
            _validate_shape(shape)
            directive = parse_foreach(shape.name, shape.shape_id)
            if "<foreach" in shape.name and directive is None:
                raise TemplateError(f"malformed foreach on source slide {index + 1}: {shape.name}")
            if (
                directive is None
                and hasattr(shape, "_pic")
                and EXPRESSION_RE.search(shape.name)
                and parse_single_image(shape.name) is None
            ):
                raise TemplateError(f"malformed image directive: {shape.name}")
            root_text = False
            if (
                directive is None
                and shape.has_text_frame
                and (
                    EXPRESSION_RE.search(shape.text)
                    or "<if" in shape.text
                    or "<foreach" in shape.text
                )
            ):
                root_text = True
                directive = ForeachDirective("", f"__text_{shape.shape_id}", shape.name)
            if directive is None:
                continue
            if (shape.rotation or 0) % 360:
                raise TemplateError("loop layout does not support rotated prototypes")
            if not hasattr(shape, "_pic") and (
                directive.rows is not None or directive.cols is not None or directive.image_mode
            ):
                raise TemplateError("rows/cols and size suffix are only supported on pictures")
            if (
                hasattr(shape, "_pic")
                and directive.body
                and not EXPRESSION_RE.fullmatch(directive.body)
            ):
                raise TemplateError("image selector must be a single {{field}} expression")
            if (
                shape.has_table
                and ("{{" in directive.body or "}}" in directive.body)
                and not EXPRESSION_RE.fullmatch(directive.body)
            ):
                raise TemplateError("table selector must be a single {{field}} expression")
            key = directive.item_path, directive.block_name
            block = grouped.setdefault(key, PageBlock(*key, [], root_text))
            block.prototypes.append(ShapePrototype(shape, directive, deepcopy(shape._element)))
            dynamic_ids.add(shape.shape_id)
        blocks = sorted(grouped.values(), key=lambda b: (b.top, b.left))
        for block in blocks:
            pages = {p.directive.page for p in block.prototypes if p.directive.page}
            if len(pages) > 1:
                raise TemplateError(f"block {block.block_name!r} has inconsistent page attributes")
            for other in blocks:
                if other is block:
                    continue
                if (
                    block.left < other.right
                    and other.left < block.right
                    and block.top < other.bottom
                    and other.top < block.bottom
                ):
                    raise TemplateError(f"block {block.block_name!r} overlaps {other.block_name!r}")
        plans.append(PagePlan(slide, index, blocks, dynamic_ids))
    return plans
