"""Two-argument service used by HTTP, workers and Python callers."""

from __future__ import annotations

import json
import math
import re
from dataclasses import dataclass
from itertools import chain
from typing import Any
from uuid import uuid4

from ..core.errors import TemplateError
from ..core.settings import Settings
from ..engine.generation import generate_presentation
from ..engine.package import TemplateSource

INVALID_XML_TEXT = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


def reject_duplicate_fields(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result = {}
    for key, value in pairs:
        if key in result:
            raise TemplateError(f"duplicate JSON field: {key}")
        result[key] = value
    return result


def parse_data(data: str, settings: Settings) -> dict[str, Any]:
    if not isinstance(data, str):
        raise TemplateError("data must be a JSON string")
    try:
        if (
            len(data) > settings.max_data_bytes
            or len(data.encode("utf-8")) > settings.max_data_bytes
        ):
            raise TemplateError("data JSON size exceeds limit")
        payload = json.loads(data, object_pairs_hook=reject_duplicate_fields)
    except (ValueError, RecursionError, UnicodeError) as exc:
        if isinstance(exc, TemplateError):
            raise
        raise TemplateError("data is not valid JSON") from exc
    if not isinstance(payload, dict):
        raise TemplateError("data JSON root must be an object")
    # Iterator frames grow with depth, not with the number of array members.
    # A wide unreferenced JSON array must not create millions of traversal tuples.
    stack = [(iter((payload,)), 1)]
    nodes = 0
    while stack:
        values, depth = stack[-1]
        try:
            value = next(values)
        except StopIteration:
            stack.pop()
            continue
        nodes += 1
        if nodes > settings.max_json_nodes:
            raise TemplateError("data JSON node count exceeds limit")
        if depth > settings.max_json_depth:
            raise TemplateError("data JSON nesting exceeds limit")
        if isinstance(value, dict):
            stack.append((chain.from_iterable(value.items()), depth + 1))
        elif isinstance(value, list):
            stack.append((iter(value), depth + 1))
        elif isinstance(value, float) and not math.isfinite(value):
            raise TemplateError("non-finite JSON numbers are not supported")
        elif isinstance(value, str) and INVALID_XML_TEXT.search(value):
            raise TemplateError("JSON text contains characters not supported by PowerPoint XML")
    return payload


@dataclass(frozen=True)
class PresentationService:
    settings: Settings

    def generate(self, template: TemplateSource, data: str) -> str:
        """Generate one uniquely named PPTX and return its absolute local path."""
        payload = parse_data(data, self.settings)
        # The caller never chooses a destination, so concurrent requests cannot
        # overwrite each other or the template. Nothing is created on validation failure.
        output = self.settings.output_dir / f"{uuid4().hex}.pptx"
        result = generate_presentation(template, payload, output, config=self.settings.generator)
        return str(result.resolve())


def generate_ppt(template: TemplateSource, data: str) -> str:
    """Convenience entry: exactly two business arguments, settings from environment."""
    return PresentationService(Settings.from_env()).generate(template, data)
