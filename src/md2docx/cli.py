from __future__ import annotations

import argparse
from pathlib import Path

from .converter import convert_markdown


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="md2docx",
        description="Convert a Markdown file to DOCX using YAML-defined styles.",
    )
    parser.add_argument("input", type=Path, help="input Markdown file")
    parser.add_argument("-o", "--output", type=Path, help="output DOCX file")
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=Path("config.yaml"),
        help="YAML style configuration (default: config.yaml)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    output = args.output or args.input.with_suffix(".docx")
    try:
        convert_markdown(args.input, output, args.config)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    print(f"Created {output}")
    return 0
