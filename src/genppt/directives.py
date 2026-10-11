from __future__ import annotations

import re
import shlex
from collections.abc import Mapping
from typing import Any

from .errors import TemplateError
from .models import ForeachDirective

FOREACH_RE = re.compile(
    r"^\s*<foreach\s+(?P<attrs>[^>]*)>(?P<body>.*?)</foreach>"
    r"\s*(?:\((?P<mode>x|y|x\s*,\s*y)\))?\s*$",
    re.DOTALL,
)
EXPRESSION_RE = re.compile(r"{{\s*([A-Za-z_][\w.]*)\s*}}")
SINGLE_IMAGE_RE = re.compile(
    r"^\s*{{\s*(?P<path>[A-Za-z_][\w.]*)\s*}}"
    r"\s*(?:\((?P<mode>x|y|x\s*,\s*y)\))?\s*$"
)
ALLOWED_FOREACH_ATTRIBUTES = frozenset({"item", "block", "rows", "cols", "page"})


def normalize_image_suffix(name: str) -> str:
    # Normalize only the terminal size suffix, never a user's literal Shape name.
    return re.sub(
        r"[（(]\s*([xy]\s*(?:[,，]\s*y)?)\s*[）)]\s*$",
        lambda match: "(" + match.group(1).replace("，", ",") + ")",
        name or "",
    )


def parse_foreach(name: str, shape_id: int) -> ForeachDirective | None:
    match = FOREACH_RE.match(normalize_image_suffix(name))
    if not match:
        return None
    attrs: dict[str, str] = {}
    try:
        tokens = shlex.split(match.group("attrs"))
    except ValueError as exc:
        raise TemplateError(f"invalid foreach directive on shape {name!r}: {exc}") from exc
    for token in tokens:
        if "=" not in token:
            raise TemplateError(f"invalid foreach attribute {token!r} on shape {name!r}")
        key, value = token.split("=", 1)
        key = key.strip()
        if key not in ALLOWED_FOREACH_ATTRIBUTES:
            raise TemplateError(f"unsupported foreach attribute {key!r} on shape {name!r}")
        if key in attrs:
            raise TemplateError(f"duplicate foreach attribute {key!r} on shape {name!r}")
        attrs[key] = value.strip()
    item_path = attrs.get("item")
    if not item_path:
        raise TemplateError(f"foreach directive is missing item: {name!r}")
    block_name = attrs.get("block") or f"__shape_{shape_id}"
    dimensions: dict[str, int | None] = {"rows": None, "cols": None}
    for key in dimensions:
        if key not in attrs:
            continue
        try:
            value = int(attrs[key])
        except ValueError as exc:
            raise TemplateError(f"foreach {key} must be a positive integer") from exc
        if value < 1:
            raise TemplateError(f"foreach {key} must be a positive integer")
        dimensions[key] = value
    mode = match.group("mode")
    if mode:
        mode = mode.replace(" ", "")
    page = attrs.get("page")
    if page is not None and page not in {"flow", "group"}:
        raise TemplateError("foreach page must be flow or group")
    if not re.fullmatch(r"[A-Za-z_][\w.]*", item_path):
        raise TemplateError("foreach item must be a data path")
    return ForeachDirective(
        item_path,
        block_name,
        match.group("body").strip(),
        dimensions["rows"],
        dimensions["cols"],
        mode,
        page,
    )


def parse_single_image(name: str) -> tuple[str, str] | None:
    match = SINGLE_IMAGE_RE.match(normalize_image_suffix(name))
    if not match:
        return None
    mode = (match.group("mode") or "x").replace(" ", "")
    return match.group("path"), mode


def optional_path(value: Any, path: str) -> Any:
    try:
        return resolve_path(value, path)
    except TemplateError:
        return None


def select_value(body: str, context: Any) -> Any:
    match = EXPRESSION_RE.fullmatch(body)
    return optional_path(context, match.group(1)) if match else context


def resolve_path(value: Any, path: str) -> Any:
    current = value
    for part in path.split("."):
        if not isinstance(current, Mapping) or part not in current:
            raise TemplateError(f"data path {path!r} is missing component {part!r}")
        current = current[part]
    return current


def render_expression(text: str, context: Mapping[str, Any]) -> str:
    def replace(match: re.Match[str]) -> str:
        try:
            value = resolve_path(context, match.group(1))
        except TemplateError:
            return ""
        return "" if value is None else str(value)

    return EXPRESSION_RE.sub(replace, text)
