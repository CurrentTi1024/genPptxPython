class GenPptError(Exception):
    """Base exception for predictable generation failures."""


class TemplateError(GenPptError, ValueError):
    """Raised when a template or payload violates the supported contract."""
