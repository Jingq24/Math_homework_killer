# -*- coding: utf-8 -*-
"""Homework solving workflow: image + question number -> AI -> handwritten PDF."""

from __future__ import annotations

import re
from pathlib import Path

from ai_solver import normalize_solution_markdown, solve_from_image
from handwriting_engine import convert_to_handwriting, load_config

WORKFLOW_ROOT = Path(__file__).resolve().parent
QUESTION_PDF_PATTERN = re.compile(r"^q(.+)\.pdf$", re.IGNORECASE)


def sanitize_stem(question_number: str) -> str:
    """Filesystem-safe output stem from question number."""
    stem = re.sub(r'[<>:"/\\|?*]', "_", question_number.strip())
    stem = stem.replace(" ", "_")
    if not stem:
        raise ValueError("Question number cannot be empty.")
    return f"q{stem}"


def get_homework_ai_config(config: dict) -> dict:
    return config.get("homework_ai") or {}


def _natural_sort_key(text: str) -> tuple:
    parts = re.split(r"(\d+)", text)
    key: list = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.lower()))
    return tuple(key)


def list_question_pdfs(output_dir: Path) -> list[Path]:
    """List q*.pdf in output_dir, sorted by question number."""
    output_dir = output_dir.resolve()
    if not output_dir.is_dir():
        return []
    pdfs = [p for p in output_dir.iterdir() if p.is_file() and QUESTION_PDF_PATTERN.match(p.name)]
    return sorted(
        pdfs,
        key=lambda p: _natural_sort_key(QUESTION_PDF_PATTERN.match(p.name).group(1)),
    )


def merge_pdfs(pdf_paths: list[Path], output_path: Path) -> Path:
    """Merge PDF files into one document."""
    if not pdf_paths:
        raise ValueError("No PDF files to merge.")
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError as exc:
        raise RuntimeError("Missing pypdf. Run: pip install -r requirements.txt") from exc

    writer = PdfWriter()
    for path in pdf_paths:
        if not path.is_file():
            raise FileNotFoundError(f"PDF not found for merge: {path}")
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)

    output_path = output_path.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path


def merge_homework_pdfs(
    output_dir: Path,
    *,
    merged_name: str | None = None,
    pdf_paths: list[Path] | None = None,
) -> Path | None:
    """Merge completed question PDFs into one file. Returns merged path or None if skipped."""
    config = load_config()
    ai_cfg = get_homework_ai_config(config)
    if not ai_cfg.get("merge_pdf", True):
        return None

    paths = pdf_paths if pdf_paths is not None else list_question_pdfs(output_dir)
    if len(paths) < 2:
        if len(paths) == 1:
            print("[merge] Only one question PDF; skip merge.")
        else:
            print("[merge] No question PDFs found; skip merge.")
        return None

    name = merged_name or ai_cfg.get("merged_pdf_name") or "homework_all.pdf"
    merged_path = output_dir / name
    print(f"[merge] Merging {len(paths)} PDF(s) -> {merged_path.name}")
    for p in paths:
        print(f"    + {p.name}")
    return merge_pdfs(paths, merged_path)


def run_homework_workflow(
    image_path: Path,
    question_number: str,
    output_dir: Path,
    *,
    chinese_font: tuple[str, float, float] | None = None,
    export_formats: list[str] | None = None,
    skip_ai: bool = False,
    markdown_path: Path | None = None,
) -> dict[str, Path]:
    """
    Full pipeline:
      1. AI solves question from image -> Markdown
      2. Existing handwriting engine -> PDF (and optional formats)
    """
    config = load_config()
    ai_cfg = get_homework_ai_config(config)
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = sanitize_stem(question_number)

    print(f"\n=== Question {question_number} ===")
    print("[1/2] Generate solution (AI)")
    if skip_ai:
        if markdown_path is None or not markdown_path.is_file():
            raise ValueError("--skip-ai requires --markdown pointing to an existing file.")
        markdown_text = markdown_path.read_text(encoding="utf-8")
        markdown_text = normalize_solution_markdown(markdown_text, question_number)
        print(f"    - Using existing Markdown: {markdown_path}")
    else:
        if not ai_cfg.get("enabled", True):
            raise RuntimeError("homework_ai.enabled is false in config.json.")
        markdown_text = solve_from_image(image_path, question_number, ai_cfg)

    md_out = output_dir / f"{stem}.md"
    md_out.write_text(markdown_text, encoding="utf-8")
    print(f"    - Saved Markdown: {md_out}")

    formats = export_formats or ai_cfg.get("export_formats") or ["pdf"]
    print("[2/2] Convert to handwritten document")
    outputs = convert_to_handwriting(
        source=md_out,
        output_dir=output_dir,
        output_stem=stem,
        export_formats=formats,
        chinese_font=chinese_font,
    )
    return {"question": question_number, "markdown": md_out, **outputs}


def run_homework_batch(
    image_path: Path,
    question_numbers: list[str],
    output_dir: Path,
    *,
    chinese_font: tuple[str, float, float] | None = None,
    export_formats: list[str] | None = None,
    skip_ai: bool = False,
    markdown_paths: list[Path | None] | None = None,
    merge: bool = True,
    merged_name: str | None = None,
) -> dict[str, Path | list[Path]]:
    """Process multiple questions, then merge PDFs when enabled."""
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if skip_ai and markdown_paths is not None and len(markdown_paths) != len(question_numbers):
        raise ValueError("With --skip-ai, provide one --markdown file per --number.")

    per_question: list[dict[str, Path]] = []
    batch_pdfs: list[Path] = []

    for i, number in enumerate(question_numbers):
        md_path = None
        if skip_ai and markdown_paths:
            md_path = markdown_paths[i]
        result = run_homework_workflow(
            image_path=image_path,
            question_number=number,
            output_dir=output_dir,
            chinese_font=chinese_font,
            export_formats=export_formats,
            skip_ai=skip_ai,
            markdown_path=md_path,
        )
        per_question.append(result)
        pdf = result.get("pdf")
        if pdf and pdf.exists():
            batch_pdfs.append(pdf)

    outputs: dict[str, Path | list[Path]] = {"questions": [r["markdown"] for r in per_question]}
    for fmt in ("pdf", "docx", "html", "rtf"):
        paths = [r[fmt] for r in per_question if fmt in r]
        if paths:
            outputs[fmt] = paths

    if merge and batch_pdfs:
        merged = merge_homework_pdfs(
            output_dir, merged_name=merged_name, pdf_paths=batch_pdfs
        )
        if merged:
            outputs["merged_pdf"] = merged

    return outputs
