"""Template-driven PowerPoint generation."""

from .config import GeneratorConfig
from .errors import GenPptError, TemplateError
from .service import generate_presentation

__all__ = ["GenPptError", "GeneratorConfig", "TemplateError", "generate_presentation"]
