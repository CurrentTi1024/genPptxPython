from __future__ import annotations

from typing import Any, Mapping

from .directives import EXPRESSION_RE, parse_foreach, resolve_path
from .errors import TemplateError


def _replacement(context: Mapping[str, Any], path: str) -> str:
    try:
        value = resolve_path(context, path)
    except TemplateError:
        return ""
    return "" if value is None else str(value)


def replace_paragraph_expressions(paragraph: Any, context: Mapping[str, Any]) -> int:
    runs = paragraph.runs
    if not runs:
        original = paragraph.text
        matches = list(EXPRESSION_RE.finditer(original))
        if matches:
            paragraph.text = EXPRESSION_RE.sub(
                lambda match: _replacement(context, match.group(1)), original
            )
        return len(matches)

    original = "".join(run.text for run in runs)
    matches = list(EXPRESSION_RE.finditer(original))
    if not matches:
        return 0
    starts: list[int] = []
    cursor = 0
    for run in runs:
        starts.append(cursor)
        cursor += len(run.text)

    def locate(position: int, *, end: bool = False) -> tuple[int, int]:
        if position == len(original):
            return len(runs) - 1, len(runs[-1].text)
        for index, start in enumerate(starts):
            run_end = start + len(runs[index].text)
            if start <= position < run_end or (end and position == run_end):
                return index, position - start
        return len(runs) - 1, len(runs[-1].text)

    for match in reversed(matches):
        start_run, start_offset = locate(match.start())
        end_run, end_offset = locate(match.end(), end=True)
        value = _replacement(context, match.group(1))
        if start_run == end_run:
            text = runs[start_run].text
            runs[start_run].text = text[:start_offset] + value + text[end_offset:]
            continue
        prefix = runs[start_run].text[:start_offset]
        suffix = runs[end_run].text[end_offset:]
        runs[start_run].text = prefix + value
        for index in range(start_run + 1, end_run):
            runs[index].text = ""
        runs[end_run].text = suffix
    return len(matches)


def replace_text_frame_expressions(text_frame: Any, context: Mapping[str, Any]) -> int:
    return sum(
        replace_paragraph_expressions(paragraph, context)
        for paragraph in text_frame.paragraphs
    )


def _process_shape(shape: Any, context: Mapping[str, Any]) -> int:
    if parse_foreach(shape.name, shape.shape_id) is not None:
        return 0
    count = 0
    if getattr(shape, "has_text_frame", False):
        count += replace_text_frame_expressions(shape.text_frame, context)
    if getattr(shape, "has_table", False):
        seen_cells: set[int] = set()
        for cell in shape.table.iter_cells():
            identity = id(cell._tc)
            if identity in seen_cells:
                continue
            seen_cells.add(identity)
            count += replace_text_frame_expressions(cell.text_frame, context)
    if hasattr(shape, "shapes"):
        for child in shape.shapes:
            count += _process_shape(child, context)
    return count


def process_text_expressions(presentation: Any, data: Mapping[str, Any]) -> int:
    return sum(
        _process_shape(shape, data)
        for slide in presentation.slides
        for shape in slide.shapes
    )
