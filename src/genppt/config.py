from __future__ import annotations

from dataclasses import dataclass

from pptx.util import Inches, Pt


@dataclass(frozen=True)
class GeneratorConfig:
    """Generation policy and resource limits suitable for an HTTP service."""

    bottom_margin_inches: float = 0.55
    block_gap_points: float = 12.0
    minimum_fragment_rows: int = 2
    repeat_non_table_shapes_on_continuation: bool = True
    allow_additional_template_slides: bool = False
    max_tables: int = 200
    max_columns: int = 50
    max_rows_per_table: int = 10_000
    max_total_cells: int = 200_000
    max_output_slides: int = 500
    max_cell_text_length: int = 10_000
    min_column_width_inches: float = 0.65
    max_row_height_inches: float = 1.4
    min_font_size_points: float = 8.0
    default_overflow: str = "wrap"
    max_images: int = 1_000
    max_image_bytes: int = 25 * 1024 * 1024
    max_total_image_bytes: int = 250 * 1024 * 1024
    max_image_pixels: int = 40_000_000
    allowed_image_roots: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        positive_fields = (
            "max_tables",
            "max_columns",
            "max_rows_per_table",
            "max_total_cells",
            "max_output_slides",
            "max_cell_text_length",
            "max_images",
            "max_image_bytes",
            "max_total_image_bytes",
            "max_image_pixels",
        )
        if self.bottom_margin_inches < 0 or self.block_gap_points < 0:
            raise ValueError("margins and gaps cannot be negative")
        if self.min_column_width_inches <= 0 or self.max_row_height_inches <= 0:
            raise ValueError("column and row size limits must be positive")
        if self.min_font_size_points <= 0:
            raise ValueError("minimum font size must be positive")
        if self.default_overflow not in {"wrap", "shrink", "truncate", "error"}:
            raise ValueError("unsupported default overflow policy")
        if self.minimum_fragment_rows < 1:
            raise ValueError("minimum_fragment_rows must be positive")
        for field in positive_fields:
            if getattr(self, field) < 1:
                raise ValueError(f"{field} must be positive")

    @property
    def bottom_margin(self) -> int:
        return Inches(self.bottom_margin_inches)

    @property
    def block_gap(self) -> int:
        return Pt(self.block_gap_points)
