"""Deployment settings; business calls need only template and data."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from .config import GeneratorConfig


@dataclass(frozen=True)
class Settings:
    output_dir: Path = field(default_factory=lambda: Path.cwd() / "var/generated")
    template_roots: tuple[Path, ...] = field(default_factory=lambda: (Path.cwd() / "templates",))
    generator: GeneratorConfig = field(
        default_factory=lambda: GeneratorConfig(allowed_image_roots=(str(Path.cwd() / "assets"),))
    )
    max_data_bytes: int = 20 * 1024 * 1024
    max_json_depth: int = 64
    max_json_nodes: int = 500_000
    max_concurrent_jobs: int = 2

    def __post_init__(self) -> None:
        for name in ("max_data_bytes", "max_json_depth", "max_json_nodes", "max_concurrent_jobs"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, int) or value < 1:
                raise ValueError(f"{name} must be a positive integer")
        object.__setattr__(self, "output_dir", Path(self.output_dir).resolve())
        object.__setattr__(
            self, "template_roots", tuple(Path(p).resolve() for p in self.template_roots)
        )

    @property
    def max_request_bytes(self) -> int:
        # JSON escaping and multipart framing need room beyond business data.
        return (
            max(
                self.generator.max_template_bytes + self.max_data_bytes,
                self.max_data_bytes * 6,
            )
            + 1024 * 1024
        )

    @classmethod
    def from_env(cls) -> Settings:
        def roots(name: str, default: str) -> tuple[str, ...]:
            return tuple(p for p in os.getenv(name, default).split(os.pathsep) if p)

        image_roots = roots("GENPPT_IMAGE_ROOTS", "assets")
        if not image_roots:
            # In the low-level library an empty tuple means unrestricted access.
            # Never silently turn an empty deployment variable into that policy.
            raise ValueError("GENPPT_IMAGE_ROOTS must specify at least one allowed directory")
        return cls(
            output_dir=Path(os.getenv("GENPPT_OUTPUT_DIR", "var/generated")),
            template_roots=tuple(Path(p) for p in roots("GENPPT_TEMPLATE_ROOTS", "templates")),
            generator=GeneratorConfig(allowed_image_roots=image_roots),
            max_data_bytes=int(os.getenv("GENPPT_MAX_DATA_BYTES", str(20 * 1024 * 1024))),
            max_concurrent_jobs=int(os.getenv("GENPPT_MAX_CONCURRENT_JOBS", "2")),
        )
