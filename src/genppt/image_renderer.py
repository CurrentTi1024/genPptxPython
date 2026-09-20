from __future__ import annotations

import warnings
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

from PIL import Image, ImageOps, UnidentifiedImageError

from .config import GeneratorConfig
from .directives import parse_foreach, parse_single_image, render_expression, resolve_path
from .errors import TemplateError
from .models import ImageBlock, ShapePrototype
from .pptx_utils import append_element, clone_static_slide, remove_shape, replace_shape_text


@dataclass(frozen=True)
class ImageInfo:
    path: Path
    width: int
    height: int
    file_size: int


@dataclass
class ImageCache:
    values: dict[str, ImageInfo | None]
    total_bytes: int = 0


@dataclass(frozen=True)
class ImageItem:
    context: Mapping[str, Any]
    images: tuple[ImageInfo, ...]
    width: int
    height: int
    source_index: int


def _is_picture(shape: Any) -> bool:
    return hasattr(shape, "_pic") and hasattr(shape, "image")


def _validate_image(
    raw_path: Any,
    config: GeneratorConfig,
    cache: ImageCache,
) -> ImageInfo | None:
    if raw_path is None:
        return None
    try:
        value = str(raw_path).strip()
    except Exception:
        return None
    if not value:
        return None
    path = Path(value).expanduser()
    try:
        resolved = path.resolve()
        cache_key = str(resolved)
        if cache_key in cache.values:
            return cache.values[cache_key]
        if config.allowed_image_roots and not any(
            resolved.is_relative_to(Path(root).expanduser().resolve())
            for root in config.allowed_image_roots
        ):
            cache.values[cache_key] = None
            return None
        file_size = resolved.stat().st_size
        if not resolved.is_file() or file_size > config.max_image_bytes:
            cache.values[cache_key] = None
            return None
        if cache.total_bytes + file_size > config.max_total_image_bytes:
            raise TemplateError(
                f"total image bytes exceed limit {config.max_total_image_bytes}"
            )
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(resolved) as image:
                image.verify()
            with Image.open(resolved) as image:
                width, height = ImageOps.exif_transpose(image).size
        if width < 1 or height < 1 or width * height > config.max_image_pixels:
            cache.values[cache_key] = None
            return None
    except TemplateError:
        raise
    except (
        OSError,
        ValueError,
        UnidentifiedImageError,
        Image.DecompressionBombError,
        Image.DecompressionBombWarning,
    ):
        cache.values[str(path.resolve())] = None
        return None
    info = ImageInfo(resolved, width, height, file_size)
    cache.values[cache_key] = info
    cache.total_bytes += file_size
    return info


def _picture_path(prototype: ShapePrototype, item: Any) -> Any:
    body = prototype.directive.body
    if not body:
        return item if not isinstance(item, Mapping) else None
    if not isinstance(item, Mapping):
        return None
    return render_expression(body, item).strip()


def _picture_size(shape: Any, info: ImageInfo, mode: str | None) -> tuple[int, int]:
    mode = mode or "x"
    if mode == "x":
        return shape.width, max(1, round(shape.width * info.height / info.width))
    if mode == "y":
        return max(1, round(shape.height * info.width / info.height)), shape.height
    return shape.width, shape.height


def compile_image_blocks(slide: Any) -> tuple[list[ImageBlock], set[int]]:
    grouped: dict[tuple[str, str], list[ShapePrototype]] = {}
    for shape in slide.shapes:
        directive = parse_foreach(shape.name, shape.shape_id)
        if directive is None:
            continue
        grouped.setdefault((directive.item_path, directive.block_name), []).append(
            ShapePrototype(shape, directive, deepcopy(shape._element))
        )
    blocks: list[ImageBlock] = []
    dynamic_ids: set[int] = set()
    for (item_path, block_name), prototypes in grouped.items():
        pictures = [prototype for prototype in prototypes if _is_picture(prototype.shape)]
        if not pictures:
            continue
        if any(prototype.shape.has_table for prototype in prototypes):
            continue
        if any((prototype.shape.rotation or 0) % 360 for prototype in prototypes):
            raise TemplateError(
                f"image block {block_name!r} does not support rotated prototypes"
            )
        rows_values = {p.directive.rows for p in prototypes if p.directive.rows is not None}
        cols_values = {p.directive.cols for p in prototypes if p.directive.cols is not None}
        if len(rows_values) > 1 or len(cols_values) > 1:
            raise TemplateError(
                f"image block {block_name!r} has inconsistent rows or cols attributes"
            )
        block = ImageBlock(
            item_path,
            block_name,
            prototypes,
            min(p.shape.left for p in prototypes),
            min(p.shape.top for p in prototypes),
            max(p.shape.left + p.shape.width for p in prototypes),
            max(p.shape.top + p.shape.height for p in prototypes),
            next(iter(rows_values), None),
            next(iter(cols_values), None),
        )
        blocks.append(block)
        dynamic_ids.update(p.shape.shape_id for p in prototypes)
    blocks.sort(key=lambda block: (block.top, block.left))
    return blocks, dynamic_ids


def _replace_picture(slide: Any, shape: Any, info: ImageInfo, mode: str | None) -> None:
    _, relationship_id = slide.part.get_or_add_image_part(str(info.path))
    shape._pic.blipFill.blip.rEmbed = relationship_id
    shape.crop_left = 0
    shape.crop_right = 0
    shape.crop_top = 0
    shape.crop_bottom = 0
    shape.width, shape.height = _picture_size(shape, info, mode)


def replace_single_images(
    presentation: Any,
    data: Mapping[str, Any],
    config: GeneratorConfig,
    cache: ImageCache,
) -> tuple[bool, int]:
    found = False
    count = 0
    for slide in presentation.slides:
        for shape in list(slide.shapes):
            if not _is_picture(shape):
                continue
            parsed = parse_single_image(shape.name)
            if parsed is None:
                continue
            found = True
            count += 1
            if count > config.max_images:
                raise TemplateError(f"image count exceeds configured limit {config.max_images}")
            path, mode = parsed
            try:
                raw_value = resolve_path(data, path)
            except TemplateError:
                raw_value = None
            info = _validate_image(raw_value, config, cache)
            if info is None:
                remove_shape(shape)
            else:
                _replace_picture(slide, shape, info, mode)
                shape.name = f"{path}__image"
    return found, count


def _prepare_items(
    block: ImageBlock,
    raw_items: Sequence[Any],
    config: GeneratorConfig,
    cache: ImageCache,
) -> list[ImageItem]:
    prepared: list[ImageItem] = []
    pictures = block.picture_prototypes
    for source_index, raw_item in enumerate(raw_items):
        context: Mapping[str, Any]
        if isinstance(raw_item, Mapping):
            context = raw_item
        elif len(pictures) == 1 and not pictures[0].directive.body:
            context = {"value": raw_item}
        else:
            continue
        infos: list[ImageInfo] = []
        valid = True
        item_right = block.right - block.left
        item_bottom = block.bottom - block.top
        for prototype in pictures:
            info = _validate_image(_picture_path(prototype, raw_item), config, cache)
            if info is None:
                valid = False
                break
            infos.append(info)
            width, height = _picture_size(
                prototype.shape, info, prototype.directive.image_mode
            )
            item_right = max(item_right, prototype.shape.left - block.left + width)
            item_bottom = max(item_bottom, prototype.shape.top - block.top + height)
        if valid:
            prepared.append(
                ImageItem(context, tuple(infos), item_right, item_bottom, source_index)
            )
    return prepared


def _grid(
    presentation: Any,
    block: ImageBlock,
    items: Sequence[ImageItem],
    config: GeneratorConfig,
) -> tuple[int, int, float]:
    gap = config.block_gap
    region_width = presentation.slide_width - config.bottom_margin - block.left
    region_height = presentation.slide_height - config.bottom_margin - block.top
    if region_width <= 0 or region_height <= 0:
        raise TemplateError(f"image block {block.block_name!r} has no usable region")
    item_width = max(item.width for item in items)
    item_height = max(item.height for item in items)
    cols = block.cols or 1
    rows = block.rows or 1
    width_scale = (region_width - gap * (cols - 1)) / (cols * item_width)
    height_scale = (region_height - gap * (rows - 1)) / (rows * item_height)
    scale = min(1.0, width_scale, height_scale)
    if scale <= 0:
        raise TemplateError(f"image block {block.block_name!r} grid has no usable cell size")
    scaled_width = item_width * scale
    scaled_height = item_height * scale
    if block.cols is None:
        cols = max(1, int((region_width + gap) // (scaled_width + gap)))
    if block.rows is None:
        rows = max(1, int((region_height + gap) // (scaled_height + gap)))
    return rows, cols, scale


def _scale_text(shape: Any, scale: float) -> None:
    if scale >= 1 or not getattr(shape, "has_text_frame", False):
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size:
                run.font.size = max(1, round(run.font.size * scale))


def _render_item(
    slide: Any,
    source_slide: Any,
    block: ImageBlock,
    item: ImageItem,
    left: int,
    top: int,
    scale: float,
    page_index: int,
) -> None:
    picture_index = 0
    for prototype_index, prototype in enumerate(block.prototypes):
        generated = append_element(slide, prototype.element, source_slide)
        generated.left = left + round((prototype.shape.left - block.left) * scale)
        generated.top = top + round((prototype.shape.top - block.top) * scale)
        generated.width = max(1, round(prototype.shape.width * scale))
        generated.height = max(1, round(prototype.shape.height * scale))
        context = dict(item.context)
        context["_page"] = {"index": page_index + 1}
        base_name = render_expression(prototype.directive.body, context).strip()
        generated.name = (
            f"{(base_name or 'image')[:180]}__{block.block_name}"
            f"_i{item.source_index + 1}_s{prototype_index + 1}"
        )
        if _is_picture(generated):
            info = item.images[picture_index]
            picture_index += 1
            _replace_picture(slide, generated, info, prototype.directive.image_mode)
        else:
            replace_shape_text(generated, context)
            _scale_text(generated, scale)


def expand_image_loops(
    presentation: Any,
    data: Mapping[str, Any],
    config: GeneratorConfig,
    cache: ImageCache,
    image_count_before: int,
) -> tuple[bool, int]:
    slide_specs: list[tuple[Any, ImageBlock, set[int]]] = []
    for slide in presentation.slides:
        blocks, dynamic_ids = compile_image_blocks(slide)
        if len(blocks) > 1:
            raise TemplateError("one template slide can contain only one image loop block")
        if blocks:
            slide_specs.append((slide, blocks[0], dynamic_ids))
    if not slide_specs:
        return False, image_count_before

    image_count = image_count_before
    for source_slide, block, dynamic_ids in reversed(slide_specs):
        raw_items = resolve_path(data, block.item_path)
        if not isinstance(raw_items, Sequence) or isinstance(raw_items, (str, bytes)):
            raise TemplateError(f"{block.item_path!r} must resolve to an array")
        image_count += len(raw_items) * len(block.picture_prototypes)
        if image_count > config.max_images:
            raise TemplateError(f"image count exceeds configured limit {config.max_images}")
        items = _prepare_items(block, raw_items, config, cache)
        static_elements = [
            deepcopy(shape._element)
            for shape in source_slide.shapes
            if shape.shape_id not in dynamic_ids
        ]
        for shape in list(source_slide.shapes):
            if shape.shape_id in dynamic_ids:
                remove_shape(shape)
        if not items:
            continue
        rows, cols, scale = _grid(presentation, block, items, config)
        page_capacity = rows * cols
        item_width = max(item.width for item in items) * scale
        item_height = max(item.height for item in items) * scale
        last_slide = source_slide
        for index, item in enumerate(items):
            page_index = index // page_capacity
            page_position = index % page_capacity
            if page_position == 0 and index > 0:
                if len(presentation.slides) >= config.max_output_slides:
                    raise TemplateError(
                        f"output slide count exceeds limit {config.max_output_slides}"
                    )
                last_slide = clone_static_slide(
                    presentation, source_slide, static_elements, last_slide
                )
            row = page_position // cols
            col = page_position % cols
            left = block.left + round(col * (item_width + config.block_gap))
            top = block.top + round(row * (item_height + config.block_gap))
            _render_item(
                last_slide,
                source_slide,
                block,
                item,
                left,
                top,
                scale,
                page_index,
            )
    return True, image_count


def process_images(
    presentation: Any,
    data: Mapping[str, Any],
    config: GeneratorConfig,
) -> bool:
    cache = ImageCache({})
    singles, image_count = replace_single_images(presentation, data, config, cache)
    loops, _ = expand_image_loops(
        presentation, data, config, cache, image_count
    )
    return singles or loops
