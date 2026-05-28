# -*- coding: utf-8 -*-
"""AI homework solver — upload an image, get handwritten solution PDF.

Launch: python app.py
"""

from __future__ import annotations

import json
import shutil
import subprocess
import traceback
from pathlib import Path

import gradio as gr

from handwriting_engine import font_label_for, load_config, resolve_font_selection
from homework_workflow import run_homework_batch

WORKFLOW_ROOT = Path(__file__).resolve().parent

PROVIDER_PRESETS: dict[str, dict[str, str]] = {
    "dashscope": {
        "api_key_env": "DASHSCOPE_API_KEY",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen-vl-plus",
    },
    "siliconflow": {
        "api_key_env": "SILICONFLOW_API_KEY",
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen2.5-VL-72B-Instruct",
    },
    "zhipu": {
        "api_key_env": "ZHIPU_API_KEY",
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "model": "glm-4v",
    },
    "ollama": {
        "api_key_env": "OLLAMA_API_KEY",
        "base_url": "http://localhost:11434/v1",
        "model": "llava",
    },
}


def check_prerequisites() -> None:
    if shutil.which("pandoc") is None:
        raise RuntimeError(
            "Pandoc not found. Install: winget install JohnMacFarlane.Pandoc"
        )
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        raise RuntimeError("Missing pywin32. Run: pip install -r requirements.txt")
    try:
        import pylatexenc  # noqa: F401
    except ImportError:
        raise RuntimeError("Missing pylatexenc. Run: pip install -r requirements.txt")


def build_font_choices(config: dict) -> list[tuple[str, str]]:
    labels = config.get("font_labels", {})
    configured = {fc[0] for fc in config.get("font_configs", [])}
    choices: list[tuple[str, str]] = []
    for font_name, label in labels.items():
        if font_name in configured:
            choices.append((f"{label} ({font_name})", font_name))
    return choices


def get_default_font(config: dict) -> str:
    return config.get("chinese_font", {}).get("name", "MEIYUJW")


def _apply_ai_provider(provider: str) -> dict | None:
    config_path = WORKFLOW_ROOT / "config.json"
    original = json.loads(config_path.read_text(encoding="utf-8"))
    preset = PROVIDER_PRESETS.get(provider)
    if preset is None:
        return None

    modified = json.loads(json.dumps(original))
    hw = modified.setdefault("homework_ai", {})
    hw["provider"] = provider
    hw["model"] = preset.get("model", hw.get("model"))
    hw["base_url"] = preset.get("base_url", hw.get("base_url"))
    hw["api_key_env"] = preset.get("api_key_env", hw.get("api_key_env"))
    config_path.write_text(
        json.dumps(modified, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return original


def _restore_config(original: dict | None) -> None:
    if original is None:
        return
    config_path = WORKFLOW_ROOT / "config.json"
    config_path.write_text(
        json.dumps(original, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def solve_homework(
    input_image,
    question_numbers: str,
    output_dir: str,
    font: str,
    ai_provider: str,
    formats: list[str],
    merge_pdf: bool,
):
    if input_image is None:
        return [], "**错误：** 请上传题目图片。"

    numbers = question_numbers.strip().split()
    if not numbers:
        return [], "**错误：** 请输入至少一个题号。"

    image_path = Path(input_image)
    if not image_path.exists():
        return [], f"**错误：** 图片未找到：{image_path}"

    out_dir = Path(output_dir) if output_dir else image_path.parent
    out_dir = out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not formats:
        formats = ["pdf"]

    original_config = _apply_ai_provider(ai_provider)
    try:
        config = load_config()
        check_prerequisites()
        chinese_font = resolve_font_selection(config, font)
        label = font_label_for(config, chinese_font[0])

        lines = [
            f"**字体：** {label} ({chinese_font[0]})",
            f"**AI 模型：** {ai_provider}",
            f"**题号：** {', '.join(numbers)}",
            "",
            "AI 解题中，正在生成手写 PDF...",
            "（每题可能需要几分钟，请耐心等待）",
        ]

        outputs = run_homework_batch(
            image_path=image_path,
            question_numbers=numbers,
            output_dir=out_dir,
            chinese_font=chinese_font,
            export_formats=formats,
            skip_ai=False,
            markdown_paths=None,
            merge=merge_pdf,
        )

        output_files: list[str] = []
        questions = outputs.get("questions", [])
        if isinstance(questions, list):
            for md in questions:
                if isinstance(md, Path) and md.exists():
                    lines.append(f"- [md] {md.name}")
                    output_files.append(str(md))

        pdfs = outputs.get("pdf")
        if isinstance(pdfs, list):
            for pdf in pdfs:
                if pdf.exists():
                    lines.append(f"- [pdf] {pdf.name}")
                    output_files.append(str(pdf))

        merged = outputs.get("merged_pdf")
        if isinstance(merged, Path) and merged.exists():
            lines.append(f"- [merged] {merged.name}")
            output_files.append(str(merged))

        lines.insert(0, "### 完成！")
        return output_files, "\n".join(lines)

    except subprocess.CalledProcessError as e:
        return [], f"**错误：** 外部命令执行失败：\n```\n{e}\n```"
    except Exception as e:
        return [], f"**错误：** {e}\n\n```\n{traceback.format_exc()}\n```"
    finally:
        _restore_config(original_config)


def create_ui() -> gr.Blocks:
    config = load_config()
    font_choices = build_font_choices(config)
    default_font = get_default_font(config)

    if font_choices and not any(fv == default_font for _, fv in font_choices):
        default_font = font_choices[0][1]

    hw_config = config.get("homework_ai", {})
    default_formats = hw_config.get("export_formats", ["pdf"])
    default_provider = hw_config.get("provider", "dashscope")

    with gr.Blocks(title="高数作业杀手") as app:
        gr.Markdown(
            "# 高数作业杀手\n"
            "上传题目图片，AI自动解题并导出为手写体PDF。"
        )

        with gr.Row():
            with gr.Column(scale=1):
                hw_image = gr.Image(
                    label="上传题目图片",
                    type="filepath",
                )
                hw_numbers = gr.Textbox(
                    label="题号",
                    placeholder="例如：1 2 3（空格分隔）",
                )
                hw_output_dir = gr.Textbox(
                    label="输出目录（可选）",
                    placeholder="默认：与图片相同目录",
                )
                hw_font = gr.Dropdown(
                    label="中文字体",
                    choices=font_choices,
                    value=default_font,
                )
                hw_provider = gr.Dropdown(
                    label="AI 模型",
                    choices=list(PROVIDER_PRESETS.keys()),
                    value=default_provider,
                )
                hw_formats = gr.CheckboxGroup(
                    label="导出格式",
                    choices=["docx", "pdf", "html", "rtf"],
                    value=default_formats,
                )
                hw_merge = gr.Checkbox(
                    label="合并所有题目为单个 PDF",
                    value=hw_config.get("merge_pdf", True),
                )
                hw_btn = gr.Button(
                    "开始解题", variant="primary", size="lg"
                )

            with gr.Column(scale=1):
                hw_status = gr.Markdown(
                    "准备就绪。上传题目图片，输入题号，点击 **开始解题**。"
                )
                hw_output = gr.File(
                    label="下载输出文件",
                    file_count="multiple",
                )

        hw_btn.click(
            fn=solve_homework,
            inputs=[
                hw_image, hw_numbers, hw_output_dir, hw_font,
                hw_provider, hw_formats, hw_merge,
            ],
            outputs=[hw_output, hw_status],
        )

    return app


if __name__ == "__main__":
    check_prerequisites()
    app = create_ui()
    app.launch(inbrowser=True)
