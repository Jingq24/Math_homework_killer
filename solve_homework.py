# -*- coding: utf-8 -*-
"""
Homework solving workflow entry point.

Usage:
  python solve_homework.py question.png --number 1 2
  python solve_homework.py question.jpg -n 3 -o output --font 美玉 --no-interactive
  python solve_homework.py --merge-only -o output/homework
  python solve_homework.py --skip-ai --markdown q1.md q2.md -n 1 2 -o output
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

from convert import check_prerequisites, resolve_cli_chinese_font
from handwriting_engine import font_label_for, load_config, print_font_catalog
from homework_workflow import merge_homework_pdfs, run_homework_batch


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Solve homework from question image(s) via AI, export handwritten PDF, merge all."
    )
    parser.add_argument(
        "image",
        type=Path,
        nargs="?",
        default=None,
        help="Question image (.png / .jpg); optional with --skip-ai or --merge-only",
    )
    parser.add_argument(
        "-n",
        "--number",
        nargs="+",
        default=None,
        help="One or more question numbers (e.g. 1 2 3)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Output directory",
    )
    parser.add_argument(
        "--formats",
        nargs="+",
        choices=["docx", "pdf", "html", "rtf"],
        default=None,
        help="Export formats (default: pdf from homework_ai.export_formats)",
    )
    parser.add_argument(
        "--font",
        type=str,
        default=None,
        help="Chinese handwriting font (index / Word name / label)",
    )
    parser.add_argument(
        "--no-interactive",
        action="store_true",
        help="Do not prompt for font; use config default",
    )
    parser.add_argument(
        "--list-fonts",
        action="store_true",
        help="List fonts and exit",
    )
    parser.add_argument(
        "--skip-ai",
        action="store_true",
        help="Skip AI; use --markdown file(s) as solution source",
    )
    parser.add_argument(
        "--markdown",
        type=Path,
        nargs="+",
        default=None,
        help="Existing Markdown solution(s); one per --number when using --skip-ai",
    )
    parser.add_argument(
        "--merge-only",
        action="store_true",
        help="Merge existing q*.pdf in output dir without reprocessing",
    )
    parser.add_argument(
        "--merged-name",
        type=str,
        default=None,
        help="Merged PDF filename (default: homework_all.pdf from config)",
    )
    parser.add_argument(
        "--no-merge",
        action="store_true",
        help="Do not merge PDFs even when multiple questions are processed",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_config()

    if args.list_fonts:
        print_font_catalog(config)
        return 0

    if args.merge_only:
        if args.output_dir is None:
            print("Error: --merge-only requires -o / --output-dir.", file=sys.stderr)
            return 1
        output_dir = args.output_dir.resolve()
        try:
            merged = merge_homework_pdfs(output_dir, merged_name=args.merged_name)
        except Exception as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 1
        if merged is None:
            return 1
        print(f"\nMerged PDF: {merged}")
        return 0

    if not args.number:
        print("Error: at least one --number is required.", file=sys.stderr)
        return 1

    if args.skip_ai:
        if not args.markdown:
            print("Error: --skip-ai requires --markdown.", file=sys.stderr)
            return 1
        if len(args.markdown) != len(args.number):
            print("Error: provide one --markdown file per --number.", file=sys.stderr)
            return 1
        md_paths = [p.resolve() for p in args.markdown]
        for p in md_paths:
            if not p.is_file():
                print(f"Error: markdown not found: {p}", file=sys.stderr)
                return 1
        base_dir = md_paths[0].parent
        image_path = args.image.resolve() if args.image else md_paths[0]
    else:
        if args.image is None:
            print("Error: question image required.", file=sys.stderr)
            return 1
        image_path = args.image.resolve()
        if not image_path.is_file():
            print(f"Error: image not found: {image_path}", file=sys.stderr)
            return 1
        base_dir = image_path.parent
        md_paths = None

    output_dir = (args.output_dir or base_dir).resolve()

    try:
        chinese_font = resolve_cli_chinese_font(
            config, args.font, args.no_interactive or args.skip_ai
        )
        label = font_label_for(config, chinese_font[0])
        print(f"Using Chinese font: {label} ({chinese_font[0]})")

        check_prerequisites()
        outputs = run_homework_batch(
            image_path=image_path,
            question_numbers=args.number,
            output_dir=output_dir,
            chinese_font=chinese_font,
            export_formats=args.formats,
            skip_ai=args.skip_ai,
            markdown_paths=md_paths,
            merge=not args.no_merge,
            merged_name=args.merged_name,
        )
    except subprocess.CalledProcessError as exc:
        print(f"Error: external command failed: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print("\nHomework workflow complete:")
    questions = outputs.get("questions", [])
    if isinstance(questions, list):
        for md in questions:
            if isinstance(md, Path) and md.exists():
                print(f"  [markdown] {md}")

    pdfs = outputs.get("pdf")
    if isinstance(pdfs, list):
        for pdf in pdfs:
            if pdf.exists():
                print(f"  [pdf] {pdf} ({pdf.stat().st_size} bytes)")

    merged = outputs.get("merged_pdf")
    if isinstance(merged, Path) and merged.exists():
        print(f"\nMerged handwritten PDF: {merged} ({merged.stat().st_size} bytes)")
    elif isinstance(pdfs, list) and len(pdfs) == 1 and pdfs[0].exists():
        print(f"\nHandwritten PDF: {pdfs[0]}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
