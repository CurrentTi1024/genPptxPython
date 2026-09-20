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

    def __post_init__(self) -> None:
        positive_fields = (
            "max_tables",
            "max_columns",
            "max_rows_per_table",
            "max_total_cells",
            "max_output_slides",
            "max_cell_text_length",
        )
        if self.bottom_margin_inches < 0 or self.block_gap_points < 0:
            raise ValueError("margins and gaps cannot be negative")
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
