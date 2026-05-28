# -*- coding: utf-8 -*-
"""Install fonts and apply office_handwriting-style effects via Word COM."""

from __future__ import annotations

import json
import os
import random
import re
import shutil
import subprocess
import tempfile
import uuid
import winreg
from dataclasses import dataclass
from pathlib import Path

from math_to_handwriting import (
    MATRIX_CELL,
    MATRIX_CLOSE,
    MATRIX_OPEN,
    MATRIX_ROW,
    SQRT_CLOSE,
    SQRT_OPEN,
    SUB_CLOSE,
    SUB_OPEN,
    SUP_CLOSE,
    SUP_OPEN,
    prepare_markdown,
)

WORKFLOW_ROOT = Path(__file__).resolve().parent
WORD_FORMATS = {
    "pdf": 17,
    "rtf": 6,
    "html": 10,
}

DEFAULT_STYLE = {
    "paragraph_left_indent_max": 16,
    "paragraph_first_line_indent_max": 10,
    "paragraph_space_before_max": 8,
    "line_spacing_range": [13, 20],
    "math_line_spacing_range": [16, 22],
    "position_jitter": 5.0,
    "spacing_jitter": 5.0,
    "size_jitter": 0.12,
    "word_font_switch_chance": 0.12,
    "math_font_switch_multiplier": 0.3,
    "scaling_range": [90, 112],
    "heading_position_boost": 1.3,
}

MATH_CHARS = set(
    "÷∫∑∏∂∇→∞≤≥≠±·×^′″()[]|<>=+-−*/0123456789."
    "εαβγδθλμπσξ"
    "½⅓⅔¼¾"
    "∈⊂∪∩∅∀∃≈≡∝"
    "│√"
)
MINUS_CHARS = frozenset("-−")
DEFAULT_ONE_GLYPH = "│"
MATH_WORDS = ("lim", "sin", "cos", "tan", "log", "ln", "sup", "inf", "max", "min", "det")


@dataclass
class WordRunStyle:
    font_name: str
    base_size: float
    expanded: float
    size_ratio: float
    baseline: float
    scaling: float


def load_config() -> dict:
    config_path = WORKFLOW_ROOT / "config.json"
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
    font_dir = Path(config["font_dir"])
    if not font_dir.is_absolute():
        font_dir = (WORKFLOW_ROOT / font_dir).resolve()
    config["font_dir"] = font_dir

    math_dir = config.get("mathhand_font_dir")
    if math_dir:
        math_path = Path(math_dir)
        if not math_path.is_absolute():
            math_path = (WORKFLOW_ROOT / math_path).resolve()
        config["mathhand_font_dir"] = math_path
    return config


def install_fonts(*font_dirs: Path) -> None:
    if not font_dirs:
        raise FileNotFoundError("No font directories configured.")

    user_fonts = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts"
    user_fonts.mkdir(parents=True, exist_ok=True)
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows NT\CurrentVersion\Fonts",
        0,
        winreg.KEY_SET_VALUE,
    )
    installed_any = False
    try:
        for font_dir in font_dirs:
            if not font_dir.exists():
                continue
            for pattern in ("*.ttf", "*.otf"):
                for font_file in font_dir.glob(pattern):
                    dest = user_fonts / font_file.name
                    if not dest.exists():
                        shutil.copy2(font_file, dest)
                    reg_suffix = "(TrueType)" if font_file.suffix.lower() == ".ttf" else "(OpenType)"
                    try:
                        winreg.SetValueEx(
                            key,
                            f"{font_file.name} {reg_suffix}",
                            0,
                            winreg.REG_SZ,
                            font_file.name,
                        )
                    except OSError:
                        pass
                    installed_any = True
    finally:
        winreg.CloseKey(key)

    if not installed_any:
        raise FileNotFoundError(
            "No font files found. Ensure handwriting_font_config and mathhand_font exist."
        )


def replace_file(src: Path, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        try:
            dest.unlink()
        except PermissionError:
            alt = dest.with_name(dest.stem + "_handwritten" + dest.suffix)
            shutil.move(str(src), str(alt))
            print(f"Warning: {dest} is locked, saved as {alt}")
            return alt
    shutil.move(str(src), str(dest))
    return dest


def build_skip_ranges(doc) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []

    def append(start: int, end: int) -> None:
        if end > start:
            ranges.append((start, end))

    for i in range(1, doc.InlineShapes.Count + 1):
        shape = doc.InlineShapes(i)
        append(shape.Range.Start, shape.Range.End)

    for i in range(1, doc.Fields.Count + 1):
        field = doc.Fields(i)
        append(field.Code.Start, field.Result.End)

    for i in range(1, doc.Tables.Count + 1):
        table = doc.Tables(i)
        append(table.Range.Start, table.Range.End)

    for i in range(1, doc.OMaths.Count + 1):
        omath = doc.OMaths(i)
        append(omath.Range.Start, omath.Range.End)

    return ranges


def should_skip(pos: int, text: str, skip_ranges: list[tuple[int, int]]) -> bool:
    for start, end in skip_ranges:
        if start <= pos < end:
            return True
    return text in ("\x13", "\x14", "\x15", "\r")


def resolve_chinese_font(config: dict) -> tuple[str, float, float]:
    """统一中文手写体：优先 chinese_font，否则取 font_configs 中权重最高的一项。"""
    cf = config.get("chinese_font")
    if cf and cf.get("name"):
        return (
            cf["name"],
            float(cf.get("size", 16)),
            float(cf.get("expanded", 0)),
        )
    configs = config.get("font_configs") or []
    if not configs:
        return "MEIYUJW", 16.0, 0.0
    weighted = [c for c in configs if len(c) >= 4 and c[3] > 0]
    pick = max(weighted, key=lambda c: c[3]) if weighted else configs[0]
    return pick[0], float(pick[1]), float(pick[2])


def build_font_catalog(config: dict) -> list[dict]:
    """Build selectable font list from font_configs and font_labels."""
    labels = config.get("font_labels") or {}
    catalog: list[dict] = []
    for i, cfg in enumerate(config.get("font_configs") or [], start=1):
        if len(cfg) < 3:
            continue
        name = str(cfg[0])
        catalog.append(
            {
                "index": i,
                "name": name,
                "label": labels.get(name, name),
                "size": float(cfg[1]),
                "expanded": float(cfg[2]),
            }
        )
    return catalog


def font_label_for(config: dict, font_name: str) -> str:
    labels = config.get("font_labels") or {}
    return labels.get(font_name, font_name)


def catalog_entry_to_tuple(entry: dict) -> tuple[str, float, float]:
    return entry["name"], entry["size"], entry["expanded"]


def resolve_font_selection(config: dict, selection: str) -> tuple[str, float, float]:
    """Resolve --font by index, Word name, or Chinese label substring."""
    catalog = build_font_catalog(config)
    if not catalog:
        raise ValueError("No fonts in font_configs.")

    sel = selection.strip()
    if sel.isdigit():
        idx = int(sel)
        for entry in catalog:
            if entry["index"] == idx:
                return catalog_entry_to_tuple(entry)
        raise ValueError(
            f"Font index {idx} out of range (1–{len(catalog)}). "
            f"Use --list-fonts to see options."
        )

    sel_lower = sel.lower()
    name_matches = [e for e in catalog if e["name"].lower() == sel_lower]
    if len(name_matches) == 1:
        return catalog_entry_to_tuple(name_matches[0])

    label_matches = [e for e in catalog if sel in e["label"] or sel in e["name"]]
    if len(label_matches) == 1:
        return catalog_entry_to_tuple(label_matches[0])
    if len(label_matches) > 1:
        names = ", ".join(f"{e['label']} ({e['name']})" for e in label_matches)
        raise ValueError(f"Ambiguous font '{selection}'. Matches: {names}")

    if len(name_matches) > 1:
        names = ", ".join(e["name"] for e in name_matches)
        raise ValueError(f"Ambiguous font name '{selection}'. Matches: {names}")

    available = ", ".join(f"{e['index']}. {e['label']}" for e in catalog)
    raise ValueError(f"Unknown font '{selection}'. Available: {available}")


def print_font_catalog(config: dict) -> None:
    catalog = build_font_catalog(config)
    default_name = resolve_chinese_font(config)[0]
    print("Available Chinese handwriting fonts:\n")
    for entry in catalog:
        mark = " (default)" if entry["name"] == default_name else ""
        print(
            f"  {entry['index']:2}. {entry['label']} ({entry['name']})"
            f" — size {entry['size']}, spacing {entry['expanded']}{mark}"
        )


def prompt_chinese_font(config: dict) -> tuple[str, float, float]:
    catalog = build_font_catalog(config)
    if not catalog:
        return resolve_chinese_font(config)

    default_tuple = resolve_chinese_font(config)
    default_index = next(
        (e["index"] for e in catalog if e["name"] == default_tuple[0]),
        catalog[0]["index"],
    )

    print("\n请选择中文手写字体:")
    for entry in catalog:
        marker = " *" if entry["index"] == default_index else ""
        print(f"  {entry['index']:2}. {entry['label']} ({entry['name']}){marker}")

    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            raw = input(f"\n输入序号 [默认 {default_index}]: ").strip()
        except EOFError:
            raw = ""
        if not raw:
            for entry in catalog:
                if entry["index"] == default_index:
                    return catalog_entry_to_tuple(entry)
            return default_tuple
        try:
            return resolve_font_selection(config, raw)
        except ValueError as exc:
            remaining = max_attempts - attempt - 1
            print(f"  {exc}")
            if remaining <= 0:
                raise
            print(f"  请重试（剩余 {remaining} 次）。")

    return default_tuple


def pick_font_config(font_configs: list, total_probability: int):
    """Legacy: random pick from pool. Prefer resolve_chinese_font for unified Chinese."""
    target = random.randint(0, max(total_probability - 1, 0))
    current = 0
    for cfg in font_configs:
        name, size, expanded, probability = cfg
        if probability <= 0:
            continue
        current += probability
        if target < current:
            return name, size, expanded
    name, size, expanded, _ = font_configs[1]
    return name, size, expanded


FONT_COLOR_BLACK = 0  # Word Font.Color RGB black


def apply_font_black(font_obj) -> None:
    try:
        font_obj.Color = FONT_COLOR_BLACK
    except Exception:
        pass


def force_document_black(doc) -> None:
    """Ensure all text uses black (Pandoc headings/links may use theme colors)."""
    try:
        doc.Content.Font.Color = FONT_COLOR_BLACK
    except Exception:
        pass
    for style_name in ("Normal", "Title", "Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        try:
            doc.Styles(style_name).Font.Color = FONT_COLOR_BLACK
        except Exception:
            pass


def apply_font_name(font_obj, name: str) -> None:
    for prop, value in (
        ("Name", name),
        ("NameFarEast", name),
        ("NameAscii", name),
        ("NameOther", name),
    ):
        try:
            setattr(font_obj, prop, value)
        except Exception:
            pass
    apply_font_black(font_obj)


def apply_script_markers(doc) -> None:
    """将私有区标记转为 Word 真下标/上标并移除标记。"""
    _apply_script_markers_in_range(doc, doc.Content)


def _apply_script_markers_in_range(doc, search_range) -> None:
    blocks = [
        (SUP_OPEN, SUP_CLOSE, False, True),
        (SUB_OPEN, SUB_CLOSE, True, False),
    ]
    for open_tag, close_tag, subscript, superscript in blocks:
        while True:
            find = search_range.Find
            find.ClearFormatting()
            find.Text = open_tag
            find.Forward = True
            find.Wrap = 0

            if not find.Execute():
                break

            start_pos = find.Parent.Start
            content_start = start_pos + len(open_tag)

            end_search = doc.Range(content_start, search_range.End)
            find2 = end_search.Find
            find2.ClearFormatting()
            find2.Text = close_tag
            find2.Forward = True
            find2.Wrap = 0

            if not find2.Execute():
                break

            end_pos = find2.Parent.Start
            inner = doc.Range(content_start, end_pos)
            try:
                inner.Font.Subscript = subscript
                inner.Font.Superscript = superscript
                base_size = inner.Font.Size
                if base_size and base_size > 0:
                    inner.Font.Size = max(base_size * 0.82, 8)
                inner.Font.Position = 0
                inner.Font.Spacing = 0
                inner.Font.Scaling = 100
                apply_font_black(inner.Font)
            except Exception:
                pass

            doc.Range(end_pos, end_pos + len(close_tag)).Delete()
            doc.Range(start_pos, start_pos + len(open_tag)).Delete()


def apply_sqrt_markers(doc) -> None:
    """Render √ with overlined radicand from SQRT_OPEN/SQRT_CLOSE markers."""
    _apply_sqrt_markers_in_range(doc, doc.Content)


def _apply_sqrt_markers_in_range(doc, search_range) -> None:
    while True:
        find = search_range.Find
        find.ClearFormatting()
        find.Text = SQRT_OPEN
        find.Forward = True
        find.Wrap = 0

        if not find.Execute():
            break

        start_pos = find.Parent.Start
        content_start = start_pos + len(SQRT_OPEN)

        end_search = doc.Range(content_start, search_range.End)
        find2 = end_search.Find
        find2.ClearFormatting()
        find2.Text = SQRT_CLOSE
        find2.Forward = True
        find2.Wrap = 0

        if not find2.Execute():
            break

        end_pos = find2.Parent.Start
        doc.Range(end_pos, end_pos + len(SQRT_CLOSE)).Delete()
        doc.Range(start_pos, start_pos + len(SQRT_OPEN)).Delete()

        radicand = doc.Range(start_pos, end_pos - len(SQRT_OPEN))
        try:
            radicand.Font.Overline = True
            apply_font_black(radicand.Font)
        except Exception:
            pass


def _matrix_delim_chars(kind: str) -> tuple[str, str]:
    if kind in ("p", "c"):
        return "(", ")"
    if kind == "b":
        return "[", "]"
    if kind == "B":
        return "{", "}"
    if kind == "V":
        return "\u2016", "\u2016"
    if kind == "v":
        return "|", "|"
    return "", ""


def _apply_matrix_cell_font(font_obj, font_name: str, font_size: float) -> None:
    apply_font_name(font_obj, font_name)
    try:
        font_obj.Size = font_size
        font_obj.Position = 0
        font_obj.Spacing = 0
        font_obj.Scaling = 100
    except Exception:
        pass


def _apply_matrix_data_cell(
    cell,
    cell_text: str,
    font_name: str,
    font_size: float,
    doc,
    math_cfg: dict | None = None,
    digit_one_cfg: dict | None = None,
    style: dict | None = None,
) -> None:
    display = _cell_display_text(cell_text, digit_one_cfg)
    cell.Range.Text = display
    _apply_sqrt_markers_in_range(doc, cell.Range)
    _apply_script_markers_in_range(doc, cell.Range)
    try:
        cell.Range.ParagraphFormat.Alignment = 1
        cell.VerticalAlignment = 1
    except Exception:
        pass
    math_cfg = math_cfg or {}
    style = style or DEFAULT_STYLE
    if math_cfg.get("enabled", False):
        _apply_math_handwriting_range(
            cell.Range, font_name, font_size, math_cfg, style, digit_one_cfg
        )
    else:
        _apply_matrix_cell_font(cell.Range.Font, font_name, font_size)
    if math_cfg.get("align_negative_numbers", True):
        _align_negative_number_in_range(
            cell.Range, font_name, font_size, math_cfg, digit_one_cfg
        )


def _resolve_math_jitter(math_cfg: dict, style: dict) -> dict:
    """Jitter parameters for formula / number handwriting."""
    scaling = math_cfg.get("scaling_range") or style.get("math_scaling_range") or style["scaling_range"]
    return {
        "size": float(math_cfg.get("size_jitter", style.get("math_size_jitter", 0.1))),
        "number_size": float(math_cfg.get("number_size_jitter", 0.06)),
        "position": float(math_cfg.get("position_jitter", style.get("math_position_jitter", 3.5))),
        "spacing": float(math_cfg.get("spacing_jitter", style.get("math_spacing_jitter", 4.0))),
        "scaling_min": int(scaling[0]),
        "scaling_max": int(scaling[1]),
        "italic_chance": float(math_cfg.get("italic_chance", 0.12)),
    }


def _apply_math_handwriting_char(
    font_obj,
    font_name: str,
    base_size: float,
    math_cfg: dict,
    style: dict,
    *,
    per_digit: bool = False,
    minus_prev: str = "",
    minus_next: str | None = None,
    digit_one_cfg: dict | None = None,
    one_glyph: str | None = None,
    ch: str = "",
) -> None:
    jitter = _resolve_math_jitter(math_cfg, style)
    size_j = jitter["number_size"] if per_digit else jitter["size"]
    font_size = base_size * random.uniform(1 - size_j, 1 + size_j)
    if per_digit:
        font_size *= random.uniform(1 - jitter["size"] * 0.35, 1 + jitter["size"] * 0.35)
    position = random.uniform(-jitter["position"], jitter["position"])
    spacing = float(math_cfg.get("expanded", 0)) + random.uniform(
        -jitter["spacing"], jitter["spacing"]
    )
    scaling = random.randint(jitter["scaling_min"], jitter["scaling_max"])
    italic = random.random() < jitter["italic_chance"]

    if ch in MINUS_CHARS and minus_prev is not None:
        position, spacing = _minus_style(math_cfg, minus_prev, minus_next, one_glyph)
    if one_glyph and ch == one_glyph:
        position, spacing, scaling, italic = _digit_one_style(
            digit_one_cfg or {}, minus_prev
        )

    apply_font_name(font_obj, font_name)
    try:
        font_obj.Size = max(font_size, 8)
        font_obj.Position = position
        font_obj.Spacing = spacing
        font_obj.Scaling = scaling
        font_obj.Italic = italic
    except Exception:
        pass


def _apply_math_handwriting_range(
    rng,
    font_name: str,
    base_size: float,
    math_cfg: dict,
    style: dict,
    digit_one_cfg: dict | None = None,
) -> None:
    """Per-character Patrick Hand + jitter inside a range (e.g. matrix cell)."""
    cfg = digit_one_cfg or {}
    one_enabled = bool(cfg.get("enabled", False))
    one_glyph = cfg.get("glyph", DEFAULT_ONE_GLYPH) if one_enabled else None
    text = rng.Text.replace("\r", "").replace("\x07", "")
    for idx in range(1, len(text) + 1):
        try:
            char_range = rng.Characters(idx)
            ch = char_range.Text
            if ch in ("\r", "\x07", "\n"):
                continue
            prev_ch = text[idx - 2] if idx > 1 else ""
            next_ch = text[idx] if idx < len(text) else ""
            _apply_math_handwriting_char(
                char_range.Font,
                font_name,
                base_size,
                math_cfg,
                style,
                per_digit=ch.isdigit() or (one_glyph and ch == one_glyph),
                minus_prev=prev_ch,
                minus_next=next_ch,
                digit_one_cfg=cfg if one_enabled else None,
                one_glyph=one_glyph,
                ch=ch,
            )
        except Exception:
            pass


def _resolve_matrix_font(
    matrix_style: dict,
    chinese_font: tuple[str, float, float] | None,
    math_cfg: dict | None = None,
) -> tuple[str, float]:
    math_cfg = math_cfg or {}
    if math_cfg.get("enabled", False):
        return (
            math_cfg.get("name", "Patrick Hand"),
            float(math_cfg.get("size", 15)),
        )
    if matrix_style.get("use_handwriting_fonts") and chinese_font:
        name, size, _ = chinese_font
        return name, float(size)

    name = matrix_style.get("font_name") or "Patrick Hand"
    return name, float(matrix_style.get("font_size", 15))


def _insert_matrix_table(
    doc,
    insert_pos: int,
    kind: str,
    rows: list[list[str]],
    matrix_style: dict,
    chinese_font: tuple[str, float, float] | None = None,
    math_cfg: dict | None = None,
    digit_one_cfg: dict | None = None,
    style: dict | None = None,
) -> None:
    if not rows:
        return

    nrows = len(rows)
    ncols = max(len(row) for row in rows)
    padded = [row + [""] * (ncols - len(row)) for row in rows]

    left_char, right_char = _matrix_delim_chars(kind)
    has_delim = kind != "m" and (left_char or right_char)

    if has_delim:
        total_cols = ncols + 2
        data_col_start = 2
    else:
        total_cols = ncols
        data_col_start = 1

    font_name, font_size = _resolve_matrix_font(matrix_style, chinese_font, math_cfg)
    bracket_size = font_size * float(matrix_style.get("bracket_size_scale", 1.35 + 0.25 * nrows))
    cell_style = style or DEFAULT_STYLE

    rng = doc.Range(insert_pos, insert_pos)
    table = doc.Tables.Add(Range=rng, NumRows=nrows, NumColumns=total_cols)

    if has_delim:
        left_cell = table.Cell(1, 1)
        right_cell = table.Cell(1, total_cols)
        if nrows > 1:
            left_cell.Merge(table.Cell(nrows, 1))
            right_cell.Merge(table.Cell(nrows, total_cols))

        left_cell.Range.Text = left_char
        right_cell.Range.Text = right_char
        for bracket_cell in (left_cell, right_cell):
            try:
                bracket_cell.VerticalAlignment = 1
                bracket_cell.Range.ParagraphFormat.Alignment = 1
            except Exception:
                pass
            _apply_matrix_cell_font(bracket_cell.Range.Font, font_name, bracket_size)

    for row_idx in range(nrows):
        for col_idx in range(ncols):
            data_cell = table.Cell(row_idx + 1, col_idx + data_col_start)
            _apply_matrix_data_cell(
                data_cell,
                padded[row_idx][col_idx],
                font_name,
                font_size,
                doc,
                math_cfg,
                digit_one_cfg,
                cell_style,
            )

    try:
        table.Borders.Enable = False
        table.Rows.AllowBreakAcrossPages = False
        table.Range.ParagraphFormat.SpaceAfter = 0
        table.Range.ParagraphFormat.SpaceBefore = 0
        if has_delim:
            table.Columns(1).Width = 16
            table.Columns(total_cols).Width = 16
        for col_idx in range(data_col_start, data_col_start + ncols):
            table.Columns(col_idx).Width = 30
    except Exception:
        pass


def _cell_display_text(cell: str, digit_one_cfg: dict | None = None) -> str:
    cfg = digit_one_cfg or {}
    enabled = cfg.get("enabled", True)
    one_glyph = cfg.get("glyph", DEFAULT_ONE_GLYPH)

    def sub_repl(match: re.Match[str]) -> str:
        content = match.group(1)
        if enabled and content in ("1", one_glyph, "l", "│", "丨"):
            return f"{SUB_OPEN}{one_glyph}{SUB_CLOSE}"
        if len(content) == 1 and content.isdigit():
            sub_digits = "₀₁₂₃₄₅₆₇₈₉"
            return sub_digits[int(content)]
        return content

    text = re.sub(f"{re.escape(SUB_OPEN)}(.*?){re.escape(SUB_CLOSE)}", sub_repl, cell)
    text = re.sub(f"{re.escape(SUP_OPEN)}(.*?){re.escape(SUP_CLOSE)}", r"\1", text)
    if enabled and one_glyph:
        text = text.replace("1", one_glyph)
    return text.strip()


def apply_matrix_markers(
    doc,
    matrix_style: dict | None = None,
    chinese_font: tuple[str, float, float] | None = None,
    math_cfg: dict | None = None,
    digit_one_cfg: dict | None = None,
    style: dict | None = None,
) -> None:
    """将矩阵标记转为 Word 表格方阵（合并括号 + 手写体数字）。"""
    matrix_style = matrix_style or {}
    if not matrix_style.get("use_word_table", True):
        return

    while True:
        find = doc.Content.Find
        find.ClearFormatting()
        find.Text = MATRIX_OPEN
        find.Forward = True
        find.Wrap = 0

        if not find.Execute():
            break

        start_pos = find.Parent.Start
        kind_pos = start_pos + len(MATRIX_OPEN)
        row_start = kind_pos + 1 + len(MATRIX_ROW)
        close_search = doc.Range(row_start, doc.Content.End)
        find_close = close_search.Find
        find_close.ClearFormatting()
        find_close.Text = MATRIX_CLOSE
        find_close.Forward = True
        find_close.Wrap = 0

        if not find_close.Execute():
            break

        close_pos = find_close.Parent.Start
        kind = doc.Range(kind_pos, kind_pos + 1).Text
        body_text = doc.Range(row_start, close_pos).Text
        rows = [
            [cell for cell in row_text.split(MATRIX_CELL)]
            for row_text in body_text.split(MATRIX_ROW)
            if row_text != ""
        ]

        end_pos = close_pos + len(MATRIX_CLOSE)
        block = doc.Range(start_pos, end_pos)
        block.Delete()
        try:
            _insert_matrix_table(
                doc,
                start_pos,
                kind,
                rows,
                matrix_style,
                chinese_font=chinese_font,
                math_cfg=math_cfg,
                digit_one_cfg=digit_one_cfg,
                style=style,
            )
        except Exception:
            fallback = doc.Range(start_pos, start_pos)
            fallback.Text = "[matrix]"


def is_cjk_char(ch: str) -> bool:
    if not ch or len(ch) != 1:
        return False
    code = ord(ch)
    return (
        0x4E00 <= code <= 0x9FFF
        or 0x3400 <= code <= 0x4DBF
        or 0x3000 <= code <= 0x303F
        or 0xFF00 <= code <= 0xFFEF
    )


def is_script_char(char_range) -> bool:
    try:
        return bool(char_range.Font.Subscript or char_range.Font.Superscript)
    except Exception:
        return False


def load_style(config: dict) -> dict:
    style = dict(DEFAULT_STYLE)
    style.update(config.get("handwriting_style", {}))
    return style


def is_word_break(ch: str) -> bool:
    return ch in (" ", "\t", "\n", "\r", "\x07", "\x0b")


def is_math_paragraph(text: str) -> bool:
    if any(ch in MATH_CHARS for ch in text):
        return True
    lower = text.lower()
    return any(word in lower for word in MATH_WORDS)


def is_math_char(ch: str, para_text: str, *, handwriting_digits: bool = True) -> bool:
    if is_cjk_char(ch):
        return False
    if ch in MATH_CHARS:
        return True
    if handwriting_digits and ch.isdigit():
        return is_math_paragraph(para_text) or any(
            c in MATH_CHARS for c in para_text if not c.isdigit()
        )
    if ch.isalpha() and len(ch) == 1 and ch.isascii():
        return is_math_paragraph(para_text)
    return False


def get_char_text(doc, index: int) -> str:
    try:
        return doc.Characters(index).Text
    except Exception:
        return ""


def _looks_like_digit_one(prev_ch: str, next_ch: str, one_glyph: str) -> bool:
    if prev_ch in MINUS_CHARS or prev_ch in "023456789" or prev_ch == one_glyph:
        return True
    if prev_ch in "=([{,;" or not prev_ch:
        return True
    if next_ch in "023456789.)]}," or not next_ch:
        return True
    return False


def _is_digit_at(
    doc,
    index: int,
    char_count: int,
    one_glyph: str | None = None,
) -> bool:
    ch = get_char_text(doc, index)
    prev_ch = get_char_text(doc, index - 1) if index > 1 else ""
    next_ch = get_char_text(doc, index + 1) if index < char_count else ""
    if ch in "023456789" or ch == "1":
        return True
    if one_glyph and ch == one_glyph:
        return _looks_like_digit_one(prev_ch, next_ch, one_glyph)
    return False


def _minus_context(prev_ch: str) -> str:
    if not prev_ch or prev_ch in " (=+[*/^,;&|\t\n\r\x07":
        return "unary"
    if prev_ch.isalpha() or prev_ch in "εαβγδθλμπσξ":
        return "binary"
    if prev_ch in MINUS_CHARS or prev_ch in "+*/^=":
        return "unary"
    return "none"


def _is_number_minus_at(
    doc, index: int, char_count: int, one_glyph: str | None = None
) -> bool:
    ch = get_char_text(doc, index)
    if ch not in MINUS_CHARS:
        return False
    next_ch = get_char_text(doc, index + 1) if index < char_count else ""
    if not _is_digit_at(doc, index + 1, char_count, one_glyph):
        return False
    return _minus_context(get_char_text(doc, index - 1) if index > 1 else "") != "none"


def _number_literal_start(
    doc, index: int, char_count: int, one_glyph: str | None = None
) -> int | None:
    ch = get_char_text(doc, index)
    if _is_number_minus_at(doc, index, char_count, one_glyph):
        return index
    if _is_digit_at(doc, index, char_count, one_glyph):
        start = index
        while start > 1 and _is_digit_at(doc, start - 1, char_count, one_glyph):
            start -= 1
        if start > 1 and _is_number_minus_at(doc, start - 1, char_count, one_glyph):
            return start - 1
        return start
    return None


def _number_literal_end(
    doc, start: int, char_count: int, one_glyph: str | None = None
) -> int:
    pos = start
    if get_char_text(doc, pos) in MINUS_CHARS:
        pos += 1
    while pos <= char_count and _is_digit_at(doc, pos, char_count, one_glyph):
        pos += 1
    if pos <= char_count and get_char_text(doc, pos) == ".":
        pos += 1
        while pos <= char_count and _is_digit_at(doc, pos, char_count, one_glyph):
            pos += 1
    return pos - 1


def _minus_style(
    math_cfg: dict,
    prev_ch: str,
    next_ch: str | None = None,
    one_glyph: str | None = None,
) -> tuple[float, float]:
    if one_glyph and next_ch == one_glyph:
        return (
            float(math_cfg.get("minus_with_one_position", 1.2)),
            float(math_cfg.get("minus_with_one_spacing", -1.0)),
        )
    ctx = _minus_context(prev_ch)
    if ctx == "unary":
        return (
            float(math_cfg.get("minus_position", 1.5)),
            float(math_cfg.get("minus_spacing", -1.2)),
        )
    if ctx == "binary":
        return (
            float(math_cfg.get("binary_minus_position", 0.8)),
            float(math_cfg.get("binary_minus_spacing", -0.8)),
        )
    return 0.0, 0.0


def _digit_one_style(
    digit_one_cfg: dict,
    prev_ch: str,
) -> tuple[float, float, int, bool]:
    scaling = int(digit_one_cfg.get("scaling", 85))
    italic = bool(digit_one_cfg.get("italic", True))
    if prev_ch in MINUS_CHARS:
        return (
            float(digit_one_cfg.get("after_minus_position", 1.2)),
            float(digit_one_cfg.get("after_minus_spacing", -0.5)),
            scaling,
            italic,
        )
    return (
        float(digit_one_cfg.get("position", 0.4)),
        float(digit_one_cfg.get("spacing", 0)),
        scaling,
        italic,
    )


def _apply_digit_one_font(
    font_obj,
    font_name: str,
    font_size: float,
    digit_one_cfg: dict,
    prev_ch: str,
) -> None:
    position, spacing, scaling, italic = _digit_one_style(digit_one_cfg, prev_ch)
    apply_font_name(font_obj, font_name)
    try:
        font_obj.Size = font_size
        font_obj.Position = position
        font_obj.Spacing = spacing
        font_obj.Scaling = scaling
        font_obj.Italic = italic
    except Exception:
        pass


def _align_negative_number_in_range(
    rng,
    font_name: str,
    font_size: float,
    math_cfg: dict,
    digit_one_cfg: dict | None = None,
) -> None:
    cfg = digit_one_cfg or {}
    one_glyph = cfg.get("glyph", DEFAULT_ONE_GLYPH)
    text = rng.Text.replace("\r", "").replace("\x07", "")
    if len(text) < 2 or text[0] not in MINUS_CHARS:
        return
    tail = text[1]
    if not (tail.isdigit() or (one_glyph and tail == one_glyph)):
        return
    try:
        minus = rng.Characters(1)
        apply_font_name(minus.Font, font_name)
        minus.Font.Size = font_size
        minus.Font.Scaling = 100
        minus.Font.Italic = False
        pos, spacing = _minus_style(math_cfg, "", tail, one_glyph)
        minus.Font.Position = pos
        minus.Font.Spacing = spacing
        for idx in range(2, len(text) + 1):
            glyph = rng.Characters(idx)
            ch = text[idx - 1]
            if one_glyph and ch == one_glyph and idx == 2:
                _apply_digit_one_font(glyph.Font, font_name, font_size, cfg, "-")
            else:
                apply_font_name(glyph.Font, font_name)
                glyph.Font.Size = font_size
                glyph.Font.Position = 0
                glyph.Font.Spacing = 0
                glyph.Font.Scaling = 100
                glyph.Font.Italic = False
    except Exception:
        pass


def apply_paragraph_jitter(doc, style: dict) -> None:
    left_max = style["paragraph_left_indent_max"]
    first_max = style["paragraph_first_line_indent_max"]
    space_max = style["paragraph_space_before_max"]
    line_min, line_max = style["line_spacing_range"]
    math_line_min, math_line_max = style.get("math_line_spacing_range", [16, 22])

    for para in doc.Paragraphs:
        text = para.Range.Text.strip()
        if not text:
            continue

        math_para = is_math_paragraph(text)
        fmt = para.Format
        try:
            if math_para:
                fmt.LeftIndent = random.uniform(0, left_max * 0.25)
                fmt.FirstLineIndent = 0
                fmt.SpaceBefore = random.uniform(0, space_max * 0.5)
                fmt.SpaceAfter = random.uniform(0, space_max * 0.4)
            else:
                fmt.LeftIndent = random.uniform(0, left_max)
                fmt.FirstLineIndent = random.uniform(-first_max * 0.4, first_max)
                fmt.SpaceBefore = random.uniform(0, space_max)
                fmt.SpaceAfter = random.uniform(0, space_max * 0.6)
        except Exception:
            pass
        try:
            fmt.LineSpacingRule = 5  # wdLineSpaceMultiple
            if math_para:
                fmt.LineSpacing = random.uniform(math_line_min, math_line_max)
            else:
                fmt.LineSpacing = random.uniform(line_min, line_max)
        except Exception:
            try:
                fmt.LineSpacingRule = 0
            except Exception:
                pass


def create_word_style(
    chinese_font: tuple[str, float, float],
    last_font_ratio: float,
    style: dict,
) -> tuple[WordRunStyle, float]:
    name, base_size, expanded = chinese_font
    font_ratio = last_font_ratio + (0.12 * random.random() - 0.06)
    font_ratio = min(max(font_ratio, 0.12), 0.32)
    return (
        WordRunStyle(
            font_name=name,
            base_size=base_size,
            expanded=expanded,
            size_ratio=font_ratio,
            baseline=random.uniform(-2.5, 1.5),
            scaling=random.uniform(*style["scaling_range"]),
        ),
        font_ratio,
    )


def apply_handwriting(
    doc,
    chinese_font: tuple[str, float, float],
    style: dict | None = None,
    math_font: dict | None = None,
    digit_one: dict | None = None,
) -> None:
    style = style or DEFAULT_STYLE
    math_cfg = math_font or {}
    digit_one_cfg = digit_one or {}
    one_enabled = bool(digit_one_cfg.get("enabled", False))
    one_glyph = digit_one_cfg.get("glyph", DEFAULT_ONE_GLYPH) if one_enabled else None
    math_enabled = bool(math_cfg.get("enabled", False))
    math_font_name = math_cfg.get("name", "Patrick Hand")
    math_font_size = float(math_cfg.get("size", 14))
    math_whole_para = bool(math_cfg.get("apply_to_whole_math_paragraph", True))
    handwriting_digits = bool(math_cfg.get("handwriting_digits", True))
    math_jitter = _resolve_math_jitter(math_cfg, style)

    skip_ranges = build_skip_ranges(doc)
    last_font_ratio = 0.2
    word_style: WordRunStyle | None = None
    math_number_styles: dict[int, float] = {}

    para_math_cache: dict[int, bool] = {}
    char_count = doc.Characters.Count

    for i in range(1, char_count + 1):
        char_range = doc.Characters(i)
        ch = char_range.Text
        if should_skip(char_range.Start, ch, skip_ranges):
            continue

        if one_glyph and ch == "1":
            try:
                char_range.Text = one_glyph
                ch = one_glyph
            except Exception:
                pass

        try:
            para_range = char_range.Paragraphs(1).Range
            para_text = para_range.Text
            para_idx = para_range.Start
        except Exception:
            para_text = ""
            para_idx = -1

        if para_idx not in para_math_cache and para_idx >= 0:
            para_math_cache[para_idx] = is_math_paragraph(para_text)
        math_para = para_math_cache.get(para_idx, False) if para_idx >= 0 else False

        if is_word_break(ch):
            word_style = None
            if ch == " " and not math_para and random.random() < 0.18:
                try:
                    char_range.Font.Spacing = random.uniform(1.5, 4.0)
                except Exception:
                    pass
            continue

        if is_script_char(char_range):
            if math_enabled and math_para and not is_cjk_char(ch):
                script_size = math_font_size * random.uniform(0.78, 0.88)
                _apply_math_handwriting_char(
                    char_range.Font,
                    math_font_name,
                    script_size,
                    math_cfg,
                    style,
                    per_digit=ch.isdigit(),
                    ch=ch,
                )
            else:
                if word_style is None or random.random() < style["word_font_switch_chance"] * 0.5:
                    word_style, last_font_ratio = create_word_style(
                        chinese_font, last_font_ratio, style
                    )
                apply_font_name(char_range.Font, word_style.font_name)
                script_size = word_style.base_size * (1 + word_style.size_ratio) * 0.82
                try:
                    char_range.Font.Size = max(script_size, 8)
                    char_range.Font.Position = 0
                    char_range.Font.Spacing = 0
                    char_range.Font.Scaling = 100
                except Exception:
                    pass
            continue

        math_char = is_math_char(
            ch, para_text, handwriting_digits=handwriting_digits
        )
        prev_ch = get_char_text(doc, i - 1) if i > 1 else ""
        next_ch = get_char_text(doc, i + 1) if i < char_count else ""
        if (
            one_glyph
            and ch == one_glyph
            and _looks_like_digit_one(prev_ch, next_ch, one_glyph)
        ):
            math_char = True
        if not math_char and not is_cjk_char(ch) and (
            prev_ch in MATH_CHARS or next_ch in MATH_CHARS
        ):
            math_char = True

        use_math_font = (
            math_enabled
            and not is_cjk_char(ch)
            and (
                math_char
                or (math_whole_para and math_para and (ch.isascii() or ch in MATH_CHARS))
            )
        )
        if use_math_font:
            num_start = _number_literal_start(doc, i, char_count, one_glyph)
            in_number = num_start is not None and num_start <= i <= _number_literal_end(
                doc, num_start, char_count, one_glyph
            )
            if in_number and num_start not in math_number_styles:
                base = math_font_size * random.uniform(
                    1 - math_jitter["number_size"], 1 + math_jitter["number_size"]
                )
                math_number_styles[num_start] = base
            base_size = (
                math_number_styles[num_start]
                if in_number and num_start in math_number_styles
                else math_font_size
            )
            _apply_math_handwriting_char(
                char_range.Font,
                math_font_name,
                base_size,
                math_cfg,
                style,
                per_digit=in_number or ch.isdigit(),
                minus_prev=prev_ch,
                minus_next=next_ch,
                digit_one_cfg=digit_one_cfg if one_enabled else None,
                one_glyph=one_glyph,
                ch=ch,
            )
            continue

        switch_chance = style["word_font_switch_chance"]
        if math_char or math_para:
            switch_chance *= style.get("math_font_switch_multiplier", 0.3)

        if word_style is None or random.random() < switch_chance:
            word_style, last_font_ratio = create_word_style(
                chinese_font, last_font_ratio, style
            )

        size_jitter = style["size_jitter"]
        pos_jitter = style["position_jitter"]
        spacing_jitter = style["spacing_jitter"]

        font_size = word_style.base_size * (1 + word_style.size_ratio)

        if math_char:
            font_size *= random.uniform(1 - size_jitter * 0.25, 1 + size_jitter * 0.25)
            position = 0
            spacing = 0
            scaling = 100
        else:
            font_size *= random.uniform(1 - size_jitter, 1 + size_jitter)
            position = word_style.baseline + random.uniform(-pos_jitter, pos_jitter) * 0.35
            position -= random.uniform(0.05, 0.25) * max(font_size - 14, 0)
            spacing = word_style.expanded + random.uniform(-spacing_jitter, spacing_jitter)
            scaling = int(word_style.scaling + random.uniform(-4, 4))

        apply_font_name(char_range.Font, word_style.font_name)
        italic = False
        if one_glyph and ch == one_glyph and _looks_like_digit_one(
            prev_ch, next_ch, one_glyph
        ):
            position, spacing, scaling, italic = _digit_one_style(
                digit_one_cfg, prev_ch
            )
        try:
            char_range.Font.Size = font_size
            char_range.Font.Position = position
            char_range.Font.Spacing = spacing
            char_range.Font.Scaling = scaling
            char_range.Font.Italic = italic
        except Exception:
            pass

    apply_paragraph_jitter(doc, style)


def pandoc_to_docx(source: Path, output: Path) -> None:
    subprocess.run(
        ["pandoc", str(source), "-f", "markdown", "-t", "docx", "-o", str(output)],
        check=True,
    )


def convert_to_handwriting(
    source: Path,
    output_dir: Path,
    output_stem: str | None = None,
    export_formats: list[str] | None = None,
    chinese_font: tuple[str, float, float] | None = None,
) -> dict[str, Path]:
    import win32com.client

    config = load_config()
    if chinese_font is None:
        chinese_font = resolve_chinese_font(config)
    digit_one_cfg = config.get("digit_one", {"enabled": False})
    formats = export_formats or config.get("export_formats", ["docx", "pdf"])
    stem = output_stem or source.stem

    font_dirs = [config["font_dir"]]
    math_dir = config.get("mathhand_font_dir")
    if math_dir and math_dir not in font_dirs:
        font_dirs.append(math_dir)
    install_fonts(*font_dirs)

    temp_dir = Path(tempfile.gettempdir())
    tag = uuid.uuid4().hex[:8]
    prepared_md = temp_dir / f"{stem}_{tag}_prepared.md"
    tmp_docx = temp_dir / f"{stem}_{tag}.docx"

    print("[1/4] LaTeX -> handwritten plain text")
    prepare_markdown(
        str(source),
        str(prepared_md),
        digit_one_glyph=digit_one_cfg.get("glyph", DEFAULT_ONE_GLYPH),
        digit_one_enabled=bool(digit_one_cfg.get("enabled", False)),
    )

    print("[2/4] Markdown -> Word (Pandoc)")
    pandoc_to_docx(prepared_md, tmp_docx)

    print("[3/4] Apply handwriting fonts (Word)")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.ScreenUpdating = False

    doc = word.Documents.Open(str(tmp_docx), ReadOnly=False)
    style = load_style(config)
    math_font_cfg = config.get("math_font", {})
    random.seed()
    print("    - Apply sqrt overline (radicand)")
    apply_sqrt_markers(doc)
    print("    - Apply subscript/superscript (lim, ∫, ∑, ...)")
    apply_script_markers(doc)
    print("    - Apply matrix layout (Word table)")
    apply_matrix_markers(
        doc,
        config.get("matrix_style", {}),
        chinese_font,
        math_font_cfg,
        digit_one_cfg,
        style,
    )
    if math_font_cfg.get("enabled", False):
        print(f"    - Apply formula font ({math_font_cfg.get('name', 'Patrick Hand')})")
    if digit_one_cfg.get("enabled", True):
        print(f"    - Digit one glyph ({digit_one_cfg.get('glyph', DEFAULT_ONE_GLYPH)})")
    cn_label = font_label_for(config, chinese_font[0])
    print(f"    - Apply Chinese font ({cn_label} / {chinese_font[0]})")
    apply_handwriting(
        doc, chinese_font, style, math_font=math_font_cfg, digit_one=digit_one_cfg
    )
    print("    - Set font color to black")
    force_document_black(doc)
    doc.Save()

    print("[4/4] Export files")
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs: dict[str, Path] = {}
    export_jobs: list[tuple[Path, Path, str]] = []

    for fmt in formats:
        final_path = output_dir / f"{stem}.{fmt}"
        if fmt == "docx":
            tmp_path = temp_dir / f"{stem}_{tag}_out.docx"
            doc.SaveAs2(str(tmp_path))
            export_jobs.append((tmp_path, final_path, fmt))
        elif fmt in WORD_FORMATS:
            tmp_path = temp_dir / f"{stem}_{tag}.{fmt}"
            doc.SaveAs2(str(tmp_path), FileFormat=WORD_FORMATS[fmt])
            export_jobs.append((tmp_path, final_path, fmt))

    doc.Close(False)
    word.ScreenUpdating = True
    word.Quit()

    for tmp_path, final_path, fmt in export_jobs:
        outputs[fmt] = replace_file(tmp_path, final_path)

    prepared_md.unlink(missing_ok=True)
    tmp_docx.unlink(missing_ok=True)
    return outputs
