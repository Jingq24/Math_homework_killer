# -*- coding: utf-8 -*-
"""Call vision-capable LLM APIs to solve homework from question images."""

from __future__ import annotations

import base64
import mimetypes
import os
import re
from pathlib import Path

WORKFLOW_ROOT = Path(__file__).resolve().parent


def _load_dotenv() -> None:
    try:
        from dotenv import load_dotenv

        load_dotenv(WORKFLOW_ROOT / ".env")
    except ImportError:
        pass


DEFAULT_SYSTEM_PROMPT = """你是一位严谨的数学/理科作业辅导老师。
用户会提供一道题目的图片和题号。请阅读图片中的题目，给出完整、清晰的解答过程。

输出要求（必须严格遵守）：
1. 只输出 Markdown 正文，不要输出代码围栏（不要用 ``` 包裹）。
2. 数学公式使用 LaTeX：行内用 $...$，独立公式块用 $$...$$；根号用 $\\sqrt{{...}}$ 包住整个被开方数。
3. 第一行必须是二级标题，格式：## 第{题号}题
4. 不要抄写或复述题目正文，不要写「题目」小节；标题后直接写解答步骤。
5. 不要写「解答」小标题、不要写「答案：」、不要写结尾总结或说明性文字（如「以上是…」「若…则…」等）。
6. 使用中文，步骤清楚，适合学生手写作业风格；不要编造图片中未出现的条件。
7. 若图片模糊无法辨认，仅简短说明无法识别，不要虚构答案。"""

DEFAULT_USER_TEMPLATE = "题号：{question_number}\n请根据附图完成解答，并按要求的 Markdown 格式输出。"

# 国产 / 本地 OpenAI 兼容接口预设（config.json 中 homework_ai.provider 选用）
PROVIDER_PRESETS: dict[str, dict] = {
    # 阿里云百炼 · 通义千问视觉（新用户有免费额度）
    "dashscope": {
        "api_key_env": "DASHSCOPE_API_KEY",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen-vl-plus",
    },
    # 硅基流动 · 开源视觉模型（常有免费额度）
    "siliconflow": {
        "api_key_env": "SILICONFLOW_API_KEY",
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen2-VL-7B-Instruct",
    },
    # 智谱 AI · GLM-4V 闪版（有免费额度）
    "zhipu": {
        "api_key_env": "ZHIPU_API_KEY",
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "model": "glm-4v-flash",
    },
    # 本地 Ollama（需先 ollama pull qwen2-vl 等视觉模型）
    "ollama": {
        "api_key_env": "OLLAMA_API_KEY",
        "base_url": "http://127.0.0.1:11434/v1",
        "model": "qwen2-vl",
        "api_key_optional": True,
    },
}


def resolve_ai_config(ai_cfg: dict) -> dict:
    """Apply provider preset, then homework_ai field overrides."""
    provider = ai_cfg.get("provider")
    overrides = {
        k: v
        for k, v in ai_cfg.items()
        if k != "provider" and v is not None and v != ""
    }
    if not provider:
        return overrides if overrides else dict(ai_cfg)
    preset = PROVIDER_PRESETS.get(provider)
    if preset is None:
        known = ", ".join(PROVIDER_PRESETS)
        raise ValueError(f"Unknown homework_ai.provider '{provider}'. Known: {known}")
    merged = dict(preset)
    merged.update(overrides)
    merged["provider"] = provider
    return merged


def _encode_image(image_path: Path) -> tuple[str, str]:
    mime, _ = mimetypes.guess_type(str(image_path))
    if not mime or not mime.startswith("image/"):
        suffix = image_path.suffix.lower()
        mime_map = {
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png",
            ".gif": "image/gif",
            ".webp": "image/webp",
            ".bmp": "image/bmp",
        }
        mime = mime_map.get(suffix, "image/jpeg")
    data = base64.standard_b64encode(image_path.read_bytes()).decode("ascii")
    return mime, data


def _strip_markdown_fences(text: str) -> str:
    text = text.strip()
    match = re.match(r"^```(?:markdown|md)?\s*\n(.*)\n```\s*$", text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return text


def normalize_solution_markdown(text: str, question_number: str) -> str:
    """Remove question restatement, answer labels, and trailing commentary."""
    text = _strip_markdown_fences(text)

    text = re.sub(
        r"(?ms)^#{1,4}\s*\*?\*?题目\*?\*?\s*\n.*?(?=^#{1,4}\s|\*\*解答\*\*|\Z)",
        "",
        text,
    )
    text = re.sub(
        r"(?ms)^\*\*题目\*\*\s*\n.*?(?=^\*\*解答\*\*|^#{1,4}\s|\Z)",
        "",
        text,
    )
    text = re.sub(r"(?m)^#{1,4}\s*\*?\*?解答\*?\*?\s*\n", "", text)
    text = re.sub(r"(?m)^\*\*解答\*\*\s*\n", "", text)
    text = re.sub(r"(?ms)^#{1,4}\s*\*?\*?补充.*?(?=^#{1,4}\s|\Z)", "", text)

    text = re.sub(r"(?ms)\n---\s*\n\s*以上是.*$", "", text)
    text = re.sub(r"(?ms)\n\*\*答案：?\*\*.*$", "", text)
    text = re.sub(
        r"(?ms)\n---\s*\n\s*若\s*\$?A\$?\s*不是单位矩阵.*$",
        "",
        text,
    )
    text = re.sub(r"\n---\s*$", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()

    if not re.match(r"^##\s*第", text):
        text = f"## 第{question_number}题\n\n{text}"
    else:
        text = re.sub(r"^##\s*第[^\n]+", f"## 第{question_number}题", text, count=1)
    return text.strip()


def _resolve_api_key(ai_cfg: dict) -> str:
    env_name = ai_cfg.get("api_key_env") or "DASHSCOPE_API_KEY"
    key = os.environ.get(env_name) or ai_cfg.get("api_key")
    if not key:
        if ai_cfg.get("api_key_optional"):
            return "not-needed"
        raise RuntimeError(
            f"AI API key not set. Set environment variable {env_name} "
            f"(see .env.example). Provider: {ai_cfg.get('provider', 'custom')}."
        )
    return key


def solve_from_image(
    image_path: Path,
    question_number: str,
    ai_cfg: dict,
) -> str:
    """
    Send question image to a vision LLM and return Markdown solution text.
    Uses OpenAI-compatible chat completions API.
    """
    _load_dotenv()
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise RuntimeError(
            "Missing openai package. Run: pip install -r requirements.txt"
        ) from exc

    ai_cfg = resolve_ai_config(ai_cfg)

    image_path = image_path.resolve()
    if not image_path.is_file():
        raise FileNotFoundError(f"Question image not found: {image_path}")

    mime, b64 = _encode_image(image_path)
    api_key = _resolve_api_key(ai_cfg)
    base_url = ai_cfg.get("base_url")
    if not base_url:
        raise RuntimeError(
            "homework_ai.base_url is required (or set homework_ai.provider to "
            "dashscope / siliconflow / zhipu / ollama)."
        )
    client = OpenAI(api_key=api_key, base_url=base_url)
    model = ai_cfg.get("model") or "qwen-vl-plus"
    system_prompt = ai_cfg.get("system_prompt") or DEFAULT_SYSTEM_PROMPT
    user_template = ai_cfg.get("user_prompt_template") or DEFAULT_USER_TEMPLATE
    user_text = user_template.format(question_number=question_number)

    print(f"[AI] Calling model {model} for question {question_number}...")
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": user_text},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime};base64,{b64}"},
                    },
                ],
            },
        ],
        temperature=float(ai_cfg.get("temperature", 0.3)),
        max_tokens=int(ai_cfg.get("max_tokens", 4096)),
        timeout=float(ai_cfg.get("timeout_seconds", 120)),
    )

    content = response.choices[0].message.content
    if not content or not content.strip():
        raise RuntimeError("AI returned empty response.")
    return normalize_solution_markdown(content, question_number)
