"""ASGI application factory: uvicorn genppt.main:create_app --factory."""

from __future__ import annotations

import logging

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from starlette.middleware.body_limit import RequestBodyLimitMiddleware

from .api.middleware import GenerationAdmissionMiddleware
from .api.routes import router
from .core.errors import TemplateError
from .core.settings import Settings
from .services.presentation import PresentationService

logger = logging.getLogger(__name__)


def create_app(settings: Settings | None = None) -> FastAPI:
    settings = settings or Settings.from_env()
    app = FastAPI(title="Template PowerPoint API", version="0.5.0")
    app.state.presentation_service = PresentationService(settings)
    app.add_middleware(RequestBodyLimitMiddleware, max_body_size=settings.max_request_bytes)
    app.add_middleware(GenerationAdmissionMiddleware, max_jobs=settings.max_concurrent_jobs)
    app.include_router(router)

    @app.get("/health", include_in_schema=False)
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    @app.exception_handler(TemplateError)
    async def invalid_input(request: Request, exc: TemplateError) -> JSONResponse:
        return JSONResponse({"detail": str(exc)}, status_code=422)

    @app.exception_handler(Exception)
    async def unexpected_error(request: Request, exc: Exception) -> JSONResponse:
        # Log the traceback server-side, never expose paths, payloads or internal
        # stack traces in the HTTP error response.
        logger.error("PPT generation failed", exc_info=exc)
        return JSONResponse({"detail": "generation failed"}, status_code=500)

    return app
