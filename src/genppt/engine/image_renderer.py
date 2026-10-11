from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any

from PIL import Image, ImageOps, UnidentifiedImageError

from ..core.config import GeneratorConfig
from ..core.errors import TemplateError
from .directives import optional_path, parse_single_image
from .pptx_utils import remove_shape


@dataclass(frozen=True)
class ImageInfo:
    path: Path
    width: int
    height: int
    file_size: int
    data: bytes


@dataclass
class ImageCache:
    values: dict[str, ImageInfo | None]
    total_bytes: int = 0


def _validate_image(raw_path: Any, config: GeneratorConfig, cache: ImageCache) -> ImageInfo | None:
    if not isinstance(raw_path, (str, Path)) or not str(raw_path).strip():
        return None
    cache_key = str(raw_path)
    try:
        resolved = Path(raw_path).expanduser().resolve()
        cache_key = str(resolved)
        if cache_key in cache.values:
            return cache.values[cache_key]
        if config.allowed_image_roots and not any(
            resolved.is_relative_to(Path(root).expanduser().resolve())
            for root in config.allowed_image_roots
        ):
            cache.values[cache_key] = None
            return None
        if not resolved.is_file() or resolved.stat().st_size > config.max_image_bytes:
            cache.values[cache_key] = None
            return None
        with resolved.open("rb") as stream:
            data = stream.read(config.max_image_bytes + 1)
        if len(data) > config.max_image_bytes:
            cache.values[cache_key] = None
            return None
        if cache.total_bytes + len(data) > config.max_total_image_bytes:
            raise TemplateError("total image bytes exceed limit")
        with Image.open(BytesIO(data)) as image:
            width, height = image.size
            if width < 1 or height < 1 or width * height > config.max_image_pixels:
                cache.values[cache_key] = None
                return None
            image.verify()
        with Image.open(BytesIO(data)) as image:
            # Freeze decoded/verified data for the request. This also makes WEBP
            # and EXIF-rotated JPEG safe for python-pptx and later pagination.
            if (
                image.format not in {"BMP", "GIF", "JPEG", "PNG", "TIFF"}
                or image.getexif().get(274, 1) != 1
            ):
                normalized = ImageOps.exif_transpose(image)
                width, height = normalized.size
                buffer = BytesIO()
                normalized.convert("RGBA").save(buffer, format="PNG")
                data = buffer.getvalue()
        if len(data) > config.max_image_bytes:
            cache.values[cache_key] = None
            return None
        if cache.total_bytes + len(data) > config.max_total_image_bytes:
            raise TemplateError("total image bytes exceed limit")
    except TemplateError:
        raise
    except (
        OSError,
        ValueError,
        RuntimeError,
        UnidentifiedImageError,
        Image.DecompressionBombError,
    ):
        cache.values[cache_key] = None
        return None
    info = ImageInfo(resolved, width, height, len(data), data)
    cache.values[cache_key] = info
    cache.total_bytes += len(data)
    return info


def _picture_size(shape: Any, info: ImageInfo, mode: str | None) -> tuple[int, int]:
    if mode is None or mode == "x":
        return shape.width, max(1, round(shape.width * info.height / info.width))
    if mode == "y":
        return max(1, round(shape.height * info.width / info.height)), shape.height
    return shape.width, shape.height


def _replace_picture(slide: Any, shape: Any, info: ImageInfo, mode: str | None) -> None:
    _, rid = slide.part.get_or_add_image_part(BytesIO(info.data))
    shape._pic.blipFill.blip.rEmbed = rid
    shape.crop_left = shape.crop_right = shape.crop_top = shape.crop_bottom = 0
    shape.width, shape.height = _picture_size(shape, info, mode)


def replace_single_images(
    presentation: Any,
    data: Mapping[str, Any],
    config: GeneratorConfig,
    cache: ImageCache,
) -> tuple[bool, int]:
    count = 0
    for slide in presentation.slides:
        for shape in list(slide.shapes):
            if not hasattr(shape, "_pic"):
                continue
            directive = parse_single_image(shape.name)
            if directive is None:
                continue
            count += 1
            if count > config.max_images:
                raise TemplateError("image count exceeds configured limit")
            path, mode = directive
            info = _validate_image(optional_path(data, path), config, cache)
            if info is None:
                remove_shape(shape)
            else:
                _replace_picture(slide, shape, info, mode)
                shape.name = (
                    f"{path}__image_p{presentation.slides.index(slide) + 1}_s{shape.shape_id}"
                )
    return bool(count), count
