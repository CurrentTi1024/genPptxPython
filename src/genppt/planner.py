from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .directives import EXPRESSION_RE, parse_foreach, parse_single_image
from .errors import TemplateError
from .image_renderer import compile_image_blocks
from .pptx_utils import compile_blocks


@dataclass(frozen=True)
class TemplatePlan:
    table_blocks: int
    image_blocks: int
    single_images: int
    text_expressions: int
    unsupported_foreach: tuple[str, ...]

    @property
    def has_tables(self) -> bool:
        return self.table_blocks > 0

    @property
    def has_images(self) -> bool:
        return self.image_blocks > 0 or self.single_images > 0


def _count_text_expressions(shape: Any) -> int:
    count = 0
    if getattr(shape, "has_text_frame", False):
        count += sum(
            len(EXPRESSION_RE.findall(paragraph.text))
            for paragraph in shape.text_frame.paragraphs
        )
    if getattr(shape, "has_table", False):
        count += sum(
            len(EXPRESSION_RE.findall(cell.text)) for cell in shape.table.iter_cells()
        )
    if hasattr(shape, "shapes"):
        count += sum(_count_text_expressions(child) for child in shape.shapes)
    return count


def analyze_template(presentation: Any) -> TemplatePlan:
    table_blocks = 0
    image_blocks = 0
    single_images = 0
    text_expressions = 0
    unsupported: list[str] = []
    for slide_index, slide in enumerate(presentation.slides, start=1):
        tables, table_ids = compile_blocks(slide)
        images, image_ids = compile_image_blocks(slide)
        if tables and images:
            raise TemplateError(
                f"slide {slide_index} cannot mix image-loop and table-loop blocks; "
                "use separate template slides so each layout region has one paginator"
            )
        if len(images) > 1:
            raise TemplateError(
                f"slide {slide_index} contains multiple image loop blocks"
            )
        for block in tables:
            if any(hasattr(prototype.shape, "_pic") for prototype in block.prototypes):
                raise TemplateError(
                    f"table block {block.block_name!r} cannot also contain a picture"
                )
        table_blocks += len(tables)
        image_blocks += len(images)
        claimed = table_ids | image_ids
        for shape in slide.shapes:
            single_image = parse_single_image(shape.name)
            if single_image is not None and hasattr(shape, "_pic"):
                single_images += 1
            directive = parse_foreach(shape.name, shape.shape_id)
            if "<foreach" in (shape.name or "") and directive is None:
                unsupported.append(
                    f"slide {slide_index} malformed foreach shape {shape.name!r}"
                )
            if (
                hasattr(shape, "_pic")
                and EXPRESSION_RE.search(shape.name or "")
                and single_image is None
                and directive is None
            ):
                unsupported.append(
                    f"slide {slide_index} malformed image shape {shape.name!r}"
                )
            if directive is not None and shape.shape_id not in claimed:
                unsupported.append(f"slide {slide_index} shape {shape.name!r}")
            if directive is None:
                text_expressions += _count_text_expressions(shape)
    return TemplatePlan(
        table_blocks,
        image_blocks,
        single_images,
        text_expressions,
        tuple(unsupported),
    )


def validate_plan(plan: TemplatePlan) -> None:
    if plan.unsupported_foreach:
        details = "; ".join(plan.unsupported_foreach[:5])
        raise TemplateError(
            "unsupported foreach block without a table or picture: " + details
        )
