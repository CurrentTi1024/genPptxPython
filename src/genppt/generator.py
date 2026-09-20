"""Backward-compatible imports for the original prototype module."""

from .config import GeneratorConfig
from .data import normalize_table
from .errors import GenPptError, TemplateError
from .service import generate_presentation

__all__ = [
    "GenPptError",
    "GeneratorConfig",
    "TemplateError",
    "generate_presentation",
    "normalize_table",
]
