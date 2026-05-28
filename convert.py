# -*- coding: utf-8 -*-
"""
Markdown -> handwritten Word/PDF workflow entry point.

Usage:
  python convert.py input.md
  python convert.py input.md -o output_dir
  python convert.py input.md --formats docx pdf
  python convert.py input.md --font 美玉 --no-interactive
  python convert.py --list-fonts

See README.md for the homework workflow (solve_homework.py).
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

from handwriting_engine import (
    convert_to_handwriting,
    font_label_for,
    load_config,
    print_font_catalog,
    prompt_chinese_font,
    resolve_chinese_font,
    resolve_font_selection,
)


def check_prerequisites() -> None:
    if shutil.which("pandoc") is None:
        raise RuntimeError("Pandoc not found. Install: winget install JohnMacFarlane.Pandoc")

    try:
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("Missing pywin32. Run: pip install -r requirements.txt") from exc

    try:
        import pylatexenc  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("Missing pylatexenc. Run: pip install -r requirements.txt") from exc


def resolve_cli_chinese_font(
    config: dict,
    font_arg: str | None,
    no_interactive: bool,
) -> tuple[str, float, float]:
    if font_arg is not None:
        return resolve_font_selection(config, font_arg)
    if no_interactive or not sys.stdin.isatty():
        return resolve_chinese_font(config)
    return prompt_chinese_font(config)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Markdown (with LaTeX math) to handwritten Word/PDF documents."
    )
    parser.add_argument(
        "input",
        type=Path,
        nargs="?",
        default=None,
        help="Input Markdown file (.md / .txt)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Output directory (default: same directory as input file)",
    )
    parser.add_argument(
        "--name",
        type=str,
        default=None,
        help="Output file base name without extension (default: input file stem)",
    )
    parser.add_argument(
        "--formats",
        nargs="+",
        choices=["docx", "pdf", "html", "rtf"],
        default=None,
        help="Export formats (default: docx pdf html rtf from config.json)",
    )
    parser.add_argument(
        "--font",
        type=str,
        default=None,
        metavar="FONT",
        help="Chinese font: index, Word name, or Chinese label (skips interactive menu)",
    )
    parser.add_argument(
        "--no-interactive",
        action="store_true",
        help="Do not prompt for font; use chinese_font from config.json",
    )
    parser.add_argument(
        "--list-fonts",
        action="store_true",
        help="List available Chinese handwriting fonts and exit",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_config()

    if args.list_fonts:
        print_font_catalog(config)
        return 0

    if args.input is None:
        print("Error: input file required (unless using --list-fonts).", file=sys.stderr)
        return 1

    source = args.input.resolve()
    if not source.exists():
        print(f"Error: input file not found: {source}", file=sys.stderr)
        return 1

    output_dir = (args.output_dir or source.parent).resolve()

    try:
        chinese_font = resolve_cli_chinese_font(
            config, args.font, args.no_interactive
        )
        label = font_label_for(config, chinese_font[0])
        print(f"Using Chinese font: {label} ({chinese_font[0]})")

        check_prerequisites()
        outputs = convert_to_handwriting(
            source=source,
            output_dir=output_dir,
            output_stem=args.name,
            export_formats=args.formats,
            chinese_font=chinese_font,
        )
    except subprocess.CalledProcessError as exc:
        print(f"Error: external command failed: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print("\nConversion complete:")
    for fmt, path in outputs.items():
        size = path.stat().st_size if path.exists() else 0
        print(f"  [{fmt}] {path} ({size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
