from threading import BoundedSemaphore

from starlette.responses import JSONResponse
from starlette.types import ASGIApp, Receive, Scope, Send


class GenerationAdmissionMiddleware:
    """Reject excess jobs before reading uploads rather than queueing them in memory."""

    def __init__(self, app: ASGIApp, max_jobs: int) -> None:
        self.app = app
        self.slots = BoundedSemaphore(max_jobs)

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        path, prefix = scope["path"], scope.get("root_path", "")
        if prefix and path.startswith(prefix + "/"):
            path = path[len(prefix) :]
        if path.rstrip("/") != "/api/v1/presentations":
            await self.app(scope, receive, send)
            return
        if not self.slots.acquire(blocking=False):
            response = JSONResponse(
                {"detail": "generation capacity is busy; retry later"},
                status_code=503,
                headers={"Retry-After": "1"},
            )
            await response(scope, receive, send)
            return
        try:
            await self.app(scope, receive, send)
        finally:
            self.slots.release()
