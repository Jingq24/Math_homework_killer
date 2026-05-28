# -*- coding: utf-8 -*-
"""Convert LaTeX math in markdown to plain handwritten-friendly text."""

from __future__ import annotations

import re

from pylatexenc.latex2text import LatexNodes2Text

LATEX2TEXT = LatexNodes2Text(math_mode="text")

# Word 中由 handwriting_engine 解析（避免 << >> 被 Word 特殊处理）
SUB_OPEN, SUB_CLOSE = "\uE010", "\uE011"
SUP_OPEN, SUP_CLOSE = "\uE012", "\uE013"
SQRT_OPEN, SQRT_CLOSE = "\uE030", "\uE031"
MATRIX_OPEN, MATRIX_ROW, MATRIX_CLOSE = "\uE020", "\uE021", "\uE022"
MATRIX_CELL = "\uE024"

MATRIX_ENV_KIND = {
    "pmatrix": "p",
    "bmatrix": "b",
    "Bmatrix": "B",
    "vmatrix": "v",
    "Vmatrix": "V",
    "matrix": "m",
}

LATEX_REPLACEMENTS = [
    (r"\\displaystyle\b", ""),
    (r"\\,", " "),
    (r"\\;", " "),
    (r"\\!", ""),
    (r"\\cdot", "·"),
    (r"\\times", "×"),
    (r"\\pm", "±"),
    (r"\\mp", "∓"),
    (r"\\leq", "≤"),
    (r"\\geq", "≥"),
    (r"\\neq", "≠"),
    (r"\\approx", "≈"),
    (r"\\equiv", "≡"),
    (r"\\propto", "∝"),
    (r"\\sim", "~"),
    (r"\\to", "→"),
    (r"\\rightarrow", "→"),
    (r"\\Rightarrow", "⇒"),
    (r"\\leftrightarrow", "↔"),
    (r"\\Leftrightarrow", "⇔"),
    (r"\\infty", "∞"),
    (r"\\partial", "∂"),
    (r"\\nabla", "∇"),
    (r"\\varepsilon", "ε"),
    (r"\\alpha", "α"),
    (r"\\beta", "β"),
    (r"\\gamma", "γ"),
    (r"\\delta", "δ"),
    (r"\\Delta", "Δ"),
    (r"\\pi", "π"),
    (r"\\theta", "θ"),
    (r"\\lambda", "λ"),
    (r"\\mu", "μ"),
    (r"\\sigma", "σ"),
    (r"\\xi", "ξ"),
    (r"\\omega", "ω"),
    (r"\\Omega", "Ω"),
    (r"\\forall", "∀"),
    (r"\\exists", "∃"),
    (r"\\notin", "∉"),
    (r"\\in\b", "∈"),
    (r"\\subset", "⊂"),
    (r"\\subseteq", "⊆"),
    (r"\\supset", "⊃"),
    (r"\\cup", "∪"),
    (r"\\cap", "∩"),
    (r"\\emptyset", "∅"),
    (r"\\oplus", "⊕"),
    (r"\\otimes", "⊗"),
    (r"\\perp", "⊥"),
    (r"\\parallel", "∥"),
    (r"\\angle", "∠"),
    (r"\\mathbb\s*\{?\s*R\s*\}?", "ℝ"),
    (r"\\mathbb\s*\{?\s*Z\s*\}?", "ℤ"),
    (r"\\mathbb\s*\{?\s*N\s*\}?", "ℕ"),
    (r"\\mathbb\s*\{?\s*Q\s*\}?", "ℚ"),
    (r"\\mathbb\s*\{?\s*C\s*\}?", "ℂ"),
    (r"\\iiint", "∭"),
    (r"\\oint", "∮"),
    (r"\\iint", "∬"),
    (r"\\int\b", "∫"),
    (r"\\sum", "∑"),
    (r"\\prod", "∏"),
    (r"\\limsup", "lim sup"),
    (r"\\liminf", "lim inf"),
    (r"\\lim\\limits_", "lim_"),
    (r"\\lim\b", "lim"),
    (r"\\max\b", "max"),
    (r"\\min\b", "min"),
    (r"\\sup\b", "sup"),
    (r"\\inf\b", "inf"),
    (r"\\det\b", "det"),
    (r"\\rank\b", "rank"),
    (r"\\tr\b", "tr"),
    (r"\\sin", "sin"),
    (r"\\cos", "cos"),
    (r"\\tan", "tan"),
    (r"\\cot", "cot"),
    (r"\\sec", "sec"),
    (r"\\csc", "csc"),
    (r"\\ln", "ln"),
    (r"\\log", "log"),
    (r"\\prime", "′"),
    (r"\\cdots", "⋯"),
    (r"\\arcsin", "arcsin"),
    (r"\\arccos", "arccos"),
    (r"\\arctan", "arctan"),
    (r"\\bigg\|", "|"),
    (r"\\big\|", "|"),
    (r"\\Bigg\|", "|"),
    (r"\\Big\|", "|"),
]

UNWRAP_COMMANDS = (
    "boldsymbol", "mathrm", "text", "displaystyle", "mathbf", "mathit", "mathsf", "mathtt"
)
FRAC_COMMANDS = ("frac", "dfrac", "tfrac")


def _extract_braced(text: str) -> tuple[str, str] | None:
    text = text.lstrip()
    if not text.startswith("{"):
        return None
    depth = 0
    for i, ch in enumerate(text):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[1:i], text[i + 1 :]
    return None


def _unwrap_braced_commands(expr: str) -> str:
    result = expr
    changed = True
    while changed:
        changed = False
        for cmd in UNWRAP_COMMANDS:
            token = f"\\{cmd}"
            pos = result.find(token)
            while pos != -1:
                tail = result[pos + len(token) :]
                extracted = _extract_braced(tail)
                if not extracted:
                    break
                inner, after = extracted
                result = result[:pos] + inner + after
                changed = True
                pos = result.find(token, pos + len(inner))
    return result


def _normalize_matrix_cell(cell: str) -> str:
    cell = cell.strip()
    cell = _replace_symbols_for_fracs(cell)
    cell = re.sub(r"\\(lim|sin|cos|tan|log|ln|exp|max|min|sup|inf|det)\b", r"\1", cell)
    cell = _replace_fracs(cell)
    for pattern, repl in sorted(LATEX_REPLACEMENTS, key=lambda x: len(x[0]), reverse=True):
        cell = re.sub(pattern, repl, cell)
    cell = re.sub(r"\\mathrm\b", "", cell)
    if re.search(r"\\[a-zA-Z]", cell):
        try:
            cell = LATEX2TEXT.latex_to_text(cell)
        except Exception:
            pass
    cell = wrap_script_markers(cell)
    cell = normalize_divisions(cell)
    return re.sub(r"\s+", " ", cell).strip()


def _parse_matrix_body(body: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for row_text in re.split(r"\\\\", body):
        row_text = row_text.strip()
        if not row_text:
            continue
        rows.append([part.strip() for part in row_text.split("&") if part.strip() or part == ""])
    return rows


def _format_matrix_rows(rows: list[list[str]], kind: str) -> str:
    if not rows:
        return ""

    normalized = [[_normalize_matrix_cell(cell) for cell in row] for row in rows]
    ncol = max(len(row) for row in normalized)
    row_payloads: list[str] = []
    for row in normalized:
        cells = [row[idx] if idx < len(row) else "" for idx in range(ncol)]
        row_payloads.append(MATRIX_CELL.join(cells))

    return f"{MATRIX_OPEN}{kind}{MATRIX_ROW}{MATRIX_ROW.join(row_payloads)}{MATRIX_CLOSE}"


def _replace_matrix_environments(expr: str) -> str:
    pattern = (
        r"\\begin\{(pmatrix|bmatrix|Bmatrix|vmatrix|Vmatrix|matrix)\}"
        r"([\s\S]*?)"
        r"\\end\{\1\}"
    )

    def repl(match: re.Match[str]) -> str:
        env = match.group(1)
        rows = _parse_matrix_body(match.group(2))
        return _format_matrix_rows(rows, MATRIX_ENV_KIND[env])

    return re.sub(pattern, repl, expr)


def _replace_binom(expr: str) -> str:
    def repl(match: re.Match[str]) -> str:
        top = match.group(1)
        bottom = match.group(2)
        return _format_matrix_rows([[top], [bottom]], "c")

    return re.sub(r"\\binom\{([^}]*)\}\{([^}]*)\}", repl, expr)


def _strip_fraction_part(part: str) -> str:
    return part.strip().strip("()")


def _numeric_fraction_to_decimal(num: str, den: str) -> str | None:
    num_c = _strip_fraction_part(num)
    den_c = _strip_fraction_part(den)
    if not re.fullmatch(r"-?\d+", num_c) or not re.fullmatch(r"-?\d+", den_c):
        return None
    n, d = int(num_c), int(den_c)
    if d == 0:
        return None
    if n % d == 0:
        return str(n // d)
    value = n / d
    text = f"{value:.6f}".rstrip("0").rstrip(".")
    return text or "0"


def _format_fraction_div(num: str, den: str) -> str:
    num = num.strip()
    den = den.strip()

    def wrap(part: str) -> str:
        if re.fullmatch(r"[\w\d]+", _strip_fraction_part(part)):
            return _strip_fraction_part(part)
        return f"({part.strip()})"

    return f"{wrap(num)}÷{wrap(den)}"


def _format_fraction(num: str, den: str) -> str:
    decimal = _numeric_fraction_to_decimal(num, den)
    if decimal is not None:
        return decimal
    return _format_fraction_div(num, den)


def _replace_fracs(expr: str) -> str:
    result = expr
    for cmd in FRAC_COMMANDS:
        token = f"\\{cmd}"
        while True:
            pos = result.find(token)
            if pos == -1:
                break
            tail = result[pos + len(token) :]

            braced_num = _extract_braced(tail)
            if braced_num:
                num_text, tail2 = braced_num
                braced_den = _extract_braced(tail2)
                if braced_den:
                    den_text, tail3 = braced_den
                    replacement = _format_fraction(num_text, den_text)
                    result = result[:pos] + replacement + tail3
                    continue

            shorthand = re.match(r"(\d)(\d)", tail)
            if shorthand:
                replacement = _format_fraction(shorthand.group(1), shorthand.group(2))
                result = result[:pos] + replacement + tail[shorthand.end() :]
                continue

            mixed = re.match(r"(\d+)([a-zA-Z]+)", tail)
            if mixed:
                replacement = _format_fraction(mixed.group(1), mixed.group(2))
                result = result[:pos] + replacement + tail[mixed.end() :]
                continue

            break
    return result


def _sub(content: str) -> str:
    return f"{SUB_OPEN}{content.strip()}{SUB_CLOSE}"


def _sup(content: str) -> str:
    return f"{SUP_OPEN}{content.strip()}{SUP_CLOSE}"


def wrap_script_markers(text: str) -> str:
    """将 _{..} / ^{..} 转为 Word 可识别的上下标标记（整段内容位于右下角/右上角）。"""
    text = re.sub(r"_\{([^}]+)\}", lambda m: _sub(m.group(1)), text)
    text = re.sub(r"\^\{([^}]+)\}", lambda m: _sup(m.group(1)), text)
    text = re.sub(r"_(?!\{)(\d+)", lambda m: _sub(m.group(1)), text)
    text = re.sub(r"\^(?!\{)(\d+)", lambda m: _sup(m.group(1)), text)
    text = re.sub(r"_(?!\{)([a-zA-Z])", lambda m: _sub(m.group(1)), text)
    text = re.sub(r"\^(?!\{)([a-zA-Z])", lambda m: _sup(m.group(1)), text)
    text = re.sub(r"(?<=\w)''", lambda _m: _sup("″"), text)
    text = re.sub(r"(?<=\w)'", lambda _m: _sup("′"), text)
    return text


def normalize_math_minus(text: str) -> str:
    """将公式中的 ASCII 连字符改为 Unicode 减号，负号与数字更易对齐。"""
    greek = set("εαβγδθλμπσξ")
    out: list[str] = []
    for i, ch in enumerate(text):
        if ch == "-" and i + 1 < len(text) and text[i + 1].isdigit():
            prev = text[i - 1] if i > 0 else ""
            if (
                not prev
                or prev in " (=+[*/^,;&|"
                or prev.isalpha()
                or prev in greek
            ):
                out.append("−")
                continue
        out.append(ch)
    return "".join(out)


def normalize_digit_one(text: str, glyph: str = "│", enabled: bool = True) -> str:
    """数字 1 改为无钩竖线（│），上下均不带弯钩。"""
    if not enabled or not glyph:
        return text
    return text.replace("1", glyph)


def normalize_divisions(text: str) -> str:
    def repl_paren(match: re.Match[str]) -> str:
        return _format_fraction(match.group(1), match.group(2))

    text = re.sub(r"\(([^()]+)\)/\(([^()]+)\)", repl_paren, text)
    text = re.sub(
        r"(?<![÷\w>])(\d+)/(\d+)(?![\w/])",
        lambda m: _format_fraction(m.group(1), m.group(2)),
        text,
    )
    text = re.sub(
        r"(?<![÷\w>])(\d+)/([a-zA-Z]+)(?![\w/])",
        lambda m: _format_fraction(m.group(1), m.group(2)),
        text,
    )
    text = re.sub(
        r"(?<![\w.])(\d+)÷(\d+)(?![\w.])",
        lambda m: _format_fraction(m.group(1), m.group(2)),
        text,
    )
    return text


def _replace_symbols_for_fracs(expr: str) -> str:
    """分式解析前先替换希腊字母等，便于识别简单分子/分母。"""
    for pattern, repl in sorted(LATEX_REPLACEMENTS, key=lambda x: len(x[0]), reverse=True):
        if pattern.startswith(r"\\") and (
            "mathbb" in pattern
            or pattern in (r"\\int\b", r"\\sum", r"\\prod", r"\\lim\b")
        ):
            continue
        expr = re.sub(pattern, repl, expr)
    return expr


def _replace_sqrt(expr: str) -> str:
    """Convert \\sqrt{...} to √ with marked radicand (overline applied in Word)."""
    result = expr
    while True:
        pos = result.find(r"\sqrt")
        if pos == -1:
            break
        tail = result[pos + 5 :]
        tail_ls = tail.lstrip()
        lead = len(tail) - len(tail_ls)
        if tail_ls.startswith("{"):
            extracted = _extract_braced(tail_ls)
            if not extracted:
                break
            inner, after = extracted
            inner_p = preprocess_latex(inner)
            replacement = f"√{SQRT_OPEN}{inner_p}{SQRT_CLOSE}"
            used = lead + len(tail_ls) - len(after)
            result = result[:pos] + replacement + tail[used:]
            continue
        paren = re.match(r"\s*\(([^()]+)\)", tail)
        if paren:
            inner_p = preprocess_latex(paren.group(1))
            replacement = f"√{SQRT_OPEN}{inner_p}{SQRT_CLOSE}"
            result = result[:pos] + replacement + tail[paren.end() :]
            continue
        bare = re.match(r"\s*(\d+)", tail)
        if bare:
            replacement = f"√{SQRT_OPEN}{bare.group(1)}{SQRT_CLOSE}"
            result = result[:pos] + replacement + tail[bare.end() :]
            continue
        break
    return result


def preprocess_latex(expr: str) -> str:
    s = expr.strip()
    s = re.sub(r"\s+", " ", s)
    s = _replace_matrix_environments(s)
    s = _replace_binom(s)
    s = _unwrap_braced_commands(s)
    s = re.sub(r"\\dfrac", r"\\frac", s)
    s = re.sub(r"\\frac(\d)(\d)", r"\\frac{\1}{\2}", s)
    s = _replace_sqrt(s)
    s = _replace_symbols_for_fracs(s)
    s = _replace_fracs(s)
    for pattern, repl in sorted(LATEX_REPLACEMENTS, key=lambda x: len(x[0]), reverse=True):
        s = re.sub(pattern, repl, s)
    s = re.sub(r"\\int\b", "∫", s)
    s = re.sub(r"\\iint\b", "∬", s)
    s = re.sub(r"\\iiint\b", "∭", s)
    s = re.sub(r"\\oint\b", "∮", s)
    s = re.sub(r"\\mathrm\b", "", s)
    s = re.sub(r"\\(lim|sin|cos|tan|log|ln|exp|max|min|sup|inf|det)\b", r"\1", s)
    return s


def latex_to_handwritten(expr: str) -> str:
    prepared = preprocess_latex(expr)
    text = re.sub(r"\\(lim|sin|cos|tan|log|ln|exp|max|min|sup|inf|det)\b", r"\1", prepared)

    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = _replace_matrix_environments(text)
    text = _replace_binom(text)
    text = wrap_script_markers(text)

    if re.search(r"\\[a-zA-Z]", text):
        try:
            text = LATEX2TEXT.latex_to_text(text)
        except Exception:
            pass
        text = re.sub(r"\\(lim|sin|cos|tan|log|ln|exp|max|min|sup|inf|det)\b", r"\1", text)
        text = wrap_script_markers(text)

    text = normalize_divisions(text)
    text = normalize_math_minus(text)
    return text


def finalize_handwritten_text(
    text: str,
    digit_one_glyph: str = "│",
    digit_one_enabled: bool = False,
) -> str:
    text = normalize_digit_one(text, digit_one_glyph, digit_one_enabled)
    return text


def prepare_markdown(
    source_path: str,
    output_path: str,
    digit_one_glyph: str = "│",
    digit_one_enabled: bool = False,
) -> None:
    with open(source_path, "r", encoding="utf-8") as f:
        content = f.read()

    def repl_display(match: re.Match[str]) -> str:
        return "\n\n" + latex_to_handwritten(match.group(1)) + "\n\n"

    def repl_inline(match: re.Match[str]) -> str:
        return latex_to_handwritten(match.group(1))

    content = re.sub(r"\$\$(.+?)\$\$", repl_display, content, flags=re.DOTALL)
    content = re.sub(r"\\\[(.+?)\\\]", repl_display, content, flags=re.DOTALL)
    content = re.sub(r"\\\((.+?)\\\)", repl_inline, content)
    content = re.sub(r"\$(.+?)\$", repl_inline, content)
    content = re.sub(r"\n{3,}", "\n\n", content)
    content = finalize_handwritten_text(content, digit_one_glyph, digit_one_enabled)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
