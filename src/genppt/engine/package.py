"""Bounded template loading shared by paths, byte buffers and uploaded streams."""

from __future__ import annotations

import inspect
from io import BytesIO
from pathlib import Path
from typing import BinaryIO
from zipfile import BadZipFile, ZipFile

from lxml.etree import XMLSyntaxError
from pptx import Presentation
from pptx.exc import InvalidXmlError, PackageNotFoundError

from ..core.config import GeneratorConfig
from ..core.errors import TemplateError

TemplateSource = str | Path | bytes | BinaryIO


def _read_stream(stream: BinaryIO, limit: int) -> bytes:
    chunks, size = [], 0
    while True:
        chunk = stream.read(min(1024 * 1024, limit + 1 - size))
        if not isinstance(chunk, bytes):
            raise TemplateError("template must be a synchronous binary stream")
        if not chunk:
            return b"".join(chunks)
        size += len(chunk)
        if size > limit:
            raise TemplateError("template file size exceeds limit")
        chunks.append(chunk)


def read_template(source: TemplateSource, config: GeneratorConfig) -> bytes:
    limit = config.max_template_bytes
    if isinstance(source, (str, Path)):
        path = Path(source)
        if not path.is_file():
            raise TemplateError("template does not exist or is not a regular file")
        if path.stat().st_size > limit:
            raise TemplateError("template file size exceeds limit")
        with path.open("rb") as stream:
            return _read_stream(stream, limit)
    if isinstance(source, bytes):
        if len(source) > limit:
            raise TemplateError("template file size exceeds limit")
        return source
    if not callable(getattr(source, "read", None)) or inspect.iscoroutinefunction(source.read):
        raise TemplateError("template must be a path, bytes or binary stream")
    # Uploaded streams may already have been inspected. Rewind if possible,
    # then restore the caller's position; never close a caller-owned stream.
    position = None
    try:
        position = source.tell()
        source.seek(0)
    except (AttributeError, OSError, ValueError):
        position = None
    try:
        return _read_stream(source, limit)
    except (OSError, ValueError) as exc:
        if isinstance(exc, TemplateError):
            raise
        raise TemplateError("template stream cannot be read") from exc
    finally:
        if position is not None:
            source.seek(position)


def load_template(source: TemplateSource, config: GeneratorConfig):
    data = read_template(source, config)
    try:
        with ZipFile(BytesIO(data)) as archive:
            parts = archive.infolist()
            if (
                len(parts) > config.max_template_parts
                or sum(part.file_size for part in parts) > config.max_template_uncompressed_bytes
            ):
                raise TemplateError("uncompressed template size or part count exceeds limit")
            if len({part.filename for part in parts}) != len(parts):
                raise TemplateError("duplicate template package parts are not supported")
        return Presentation(BytesIO(data))
    except TemplateError:
        raise
    except (
        BadZipFile,
        XMLSyntaxError,
        PackageNotFoundError,
        InvalidXmlError,
        KeyError,
        ValueError,
        RuntimeError,
        NotImplementedError,
    ) as exc:
        raise TemplateError("template is not a valid PowerPoint package") from exc
