"""Template-driven PowerPoint generation."""

from .config import GeneratorConfig
from .errors import GenPptError, TemplateError
from .service import generate_presentation
from .services.presentation import PresentationService, generate_ppt

__all__ = [
    "GenPptError",
    "GeneratorConfig",
    "TemplateError",
    "PresentationService",
    "generate_ppt",
    "generate_presentation",
]
