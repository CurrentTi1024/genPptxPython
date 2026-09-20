from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path

from .config import GeneratorConfig
from .errors import GenPptError
from .service import generate_presentation


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate dynamic tables from a PPTX template")
    parser.add_argument("template", type=Path)
    parser.add_argument("data", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--bottom-margin-in", type=float, default=0.55)
    parser.add_argument("--block-gap-pt", type=float, default=12.0)
    parser.add_argument("--allow-additional-template-slides", action="store_true")
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
    )
    args = parser.parse_args()

    logging.basicConfig(level=getattr(logging, args.log_level), format="%(levelname)s %(message)s")
    config = GeneratorConfig(
        bottom_margin_inches=args.bottom_margin_in,
        block_gap_points=args.block_gap_pt,
        allow_additional_template_slides=args.allow_additional_template_slides,
    )
    try:
        payload = json.loads(args.data.read_text(encoding="utf-8"))
        result = generate_presentation(args.template, payload, args.output, config=config)
    except (OSError, json.JSONDecodeError, GenPptError) as exc:
        parser.exit(2, f"generation failed: {exc}\n")
    logging.info("generated %s", result)


if __name__ == "__main__":
    main()
