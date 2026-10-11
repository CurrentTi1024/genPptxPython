from __future__ import annotations

import json
from pathlib import Path

from fastapi import APIRouter, HTTPException, Request
from pydantic import ValidationError
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import UploadFile

from ..core.settings import Settings
from ..services.presentation import reject_duplicate_fields
from .schemas import GenerationRequest

router = APIRouter(prefix="/api/v1")


def _allowed_template(path: str, settings: Settings) -> Path:
    try:
        resolved = Path(path).resolve()
    except (ValueError, OSError, RuntimeError) as exc:
        raise HTTPException(422, "invalid template path") from exc
    if not any(resolved.is_relative_to(root) for root in settings.template_roots):
        raise HTTPException(403, "template path is outside configured roots")
    return resolved


@router.post(
    "/presentations",
    response_model=str,
    summary="Generate PPTX from a local template path or uploaded template",
    openapi_extra={
        "requestBody": {
            "required": True,
            "content": {
                "application/json": {"schema": GenerationRequest.model_json_schema()},
                "multipart/form-data": {
                    "schema": {
                        "type": "object",
                        "required": ["template", "data"],
                        "additionalProperties": False,
                        "properties": {
                            "template": {"type": "string", "format": "binary"},
                            "data": {"type": "string", "description": "JSON object string"},
                        },
                    }
                },
            },
        },
    },
)
async def generate(request: Request) -> str:
    service = request.app.state.presentation_service
    settings = service.settings
    content_type = request.headers.get("content-type", "").split(";", 1)[0].strip().lower()
    if content_type == "application/json":
        try:
            body = json.loads(await request.body(), object_pairs_hook=reject_duplicate_fields)
            payload = GenerationRequest.model_validate(body)
        except (ValueError, ValidationError, RecursionError, UnicodeError) as exc:
            raise HTTPException(422, "expected exactly template path and data JSON string") from exc
        # Local-path access is an HTTP security policy; the library itself can
        # still receive arbitrary caller-owned files and paths.
        template = await run_in_threadpool(_allowed_template, payload.template, settings)
        return await run_in_threadpool(service.generate, template, payload.data)
    if content_type == "multipart/form-data":
        async with request.form(
            max_files=1,
            max_fields=1,
            max_part_size=settings.max_data_bytes,
        ) as form:
            if len(form.multi_items()) != 2 or set(form) != {"template", "data"}:
                raise HTTPException(422, "expected exactly template file and data JSON string")
            template, data = form["template"], form["data"]
            if not isinstance(template, UploadFile) or not isinstance(data, str):
                raise HTTPException(422, "template must be a file and data must be a JSON string")
            if template.size is not None and template.size > settings.generator.max_template_bytes:
                raise HTTPException(413, "template file size exceeds limit")
            # Generation is blocking. Keep the upload context open until the
            # worker has finished reading it, including failure/cancellation.
            return await run_in_threadpool(service.generate, template.file, data)
    raise HTTPException(415, "use application/json or multipart/form-data")
