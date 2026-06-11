"""
AI structure recognition service.

Two modes:
1. AI-powered: calls OpenAI-compatible API (GPT, Qwen, DeepSeek) for
   paragraph classification with confidence scores.
2. Rule-based fallback: uses heuristics when AI is disabled or unavailable.
"""

from __future__ import annotations

import json
import re
from typing import Any

from app.config import (
    AI_API_KEY,
    AI_BASE_URL,
    AI_ENABLED,
    AI_MODEL,
    AI_PROVIDER,
)
from app.schemas.models import (
    ArticleContent,
    ConfidenceLevel,
    ContentDocument,
    ContentParagraph,
    ParagraphType,
)
from app.utils.word_utils import (
    is_blue_color,
    is_red_color,
)


def recognize_structure(
    paragraphs: list[dict[str, Any]],
    filename: str = "",
    style_brief: str = "",
) -> ContentDocument:
    """
    Recognize the structure of a parsed document via the cloud LLM.
    Returns a ContentDocument with typed paragraphs and confidence scores.

    Pure-AI mode: classification is delegated entirely to the configured
    OpenAI-compatible API. Rule-based fallback has been removed; failures
    propagate to the caller so the operator can act on them.

    Args:
        paragraphs: parsed paragraph dicts from `content_parser.parse_content`.
        filename: original filename, for diagnostics.
        style_brief: optional one-line-per-style description produced by
            `template_service.build_style_brief()`. When present it is
            injected into the AI prompt so the model classifies against
            the *target* style of the template batch, not just the format
            hints in the content.

    Raises:
        RuntimeError: when AI recognition is not configured, or the API call fails.
    """
    if not (AI_ENABLED and AI_API_KEY):
        raise RuntimeError(
            "AI recognition is required but not configured. "
            "Set AI_ENABLED=true and AI_API_KEY=... in the environment."
        )

    doc = ContentDocument(original_filename=filename)
    classified = _ai_classify(paragraphs, style_brief=style_brief)

    # Post-process: ensure structural invariants
    content_paragraphs = _post_process(classified)

    # Group into articles
    doc.articles = _split_articles(content_paragraphs)

    # Check for low confidence
    has_low = any(
        p.confidence == ConfidenceLevel.low
        for article in doc.articles
        for p in article.paragraphs
    )
    doc.has_low_confidence = has_low

    # Collect warnings
    doc.warnings = _collect_warnings(doc)

    return doc


# ── AI-Powered Classification ─────────────────────────────

def _describe_format(hints: dict) -> str:
    """Build a human-readable format description from paragraph hints."""
    parts = []
    first_run = hints.get("first_run", {})

    font_name = first_run.get("font_name", "")
    if font_name:
        parts.append(font_name)
    font_size = first_run.get("font_size_pt")
    if font_size:
        parts.append(f"{font_size}pt")
    if hints.get("is_bold"):
        parts.append("加粗")
    if hints.get("is_centered"):
        parts.append("居中")
    color = first_run.get("color", "")
    if color:
        if is_blue_color(color):
            parts.append("蓝色文字")
        elif is_red_color(color):
            parts.append("红色文字")
        elif color != "000000":
            parts.append(f"#{color}")
    # Detect first-line indent from runs if available
    indent_pt = hints.get("first_line_indent_pt", 0)
    if indent_pt and indent_pt > 10:
        parts.append("首行缩进")

    return f"【格式：{', '.join(parts)}】" if parts else ""


def _ai_classify(
    paragraphs: list[dict[str, Any]],
    style_brief: str = "",
) -> list[ContentParagraph]:
    """
    Use a cloud LLM (DeepSeek / OpenAI-compatible) to classify paragraphs.

    Key improvements over the original:
    - Passes format metadata (bold, color, alignment, font, size) alongside text
    - Uses few-shot examples in the system prompt so the model learns
      the mapping from (text + format) → paragraph type
    - Optionally appends a target-style brief derived from the template
      batch (see `template_service.build_style_brief`) so the model
      classifies against the *target* style, not just the input format
    - Handles JSON parsing edge cases (markdown fences, trailing commas)
    """
    import urllib.request
    import urllib.error

    # ── Build format-aware paragraph descriptions ──
    para_lines = []
    for i, p in enumerate(paragraphs):
        txt = p.get("text", "")
        hints = p.get("format_hints", {})

        # Type hint for special elements
        type_tag = ""
        if p.get("is_table"):
            type_tag = " [表格]"
        elif p.get("has_image"):
            type_tag = " [图片]"

        format_str = _describe_format(hints)
        para_lines.append(f"[{i}]{type_tag} \"{txt}\" {format_str}")

    text_block = "\n".join(para_lines)

    # ── System prompt with few-shot examples ──
    valid_types = [t.value for t in ParagraphType]

    system_prompt = f"""你是一个专业的文档结构分析专家。根据段落的**文本内容**和**格式信息**（字号、加粗、颜色、居中、字体、缩进）综合判断每个段落的类型。

## 段落类型

- main_title: 文档主标题。特征：第一段，字号≥16pt，加粗+居中，黑体/宋体
- article_title: 二级标题。特征：红色文字，或加粗+居中，中长文本
- section_header: 章节小标题。特征：蓝色文字，或加粗非居中，较短(<30字)
- body_text: 正文内容。特征：长文本(>80字)，仿宋，首行缩进
- author_name: 作者姓名。特征：2-4汉字（可带"教授/研究员"），加粗或居中
- author_note: 作者简介。特征："姓名，单位，职称，主要从事..." 格式，较长
- affiliation: 作者单位。特征：括号开头含"作者为/作者单位/作者系"
- source: 来源信息。特征："来源：", "《XX日报》", "摘自..."
- publish_date: 发布日期。特征："2024年1月15日" 日期格式
- edition: 版次信息。特征："第09版"，含日期+版次的括号内容
- editor_note: 编者按。特征："编者按"或"【编者按】"开头
- tag_label: 标签/导读/推荐。特征：短文本，含"理论文章："、"合集"、"连发"等
- image: 图片段落（标注为 [图片]）
- image_caption: 图片说明文字
- footer: 页脚
- other: 无法分类

## 分类示例（文本 + 格式 → 类型）

1. [0] "深刻理解'两个确立'的决定性意义" 【格式：黑体, 22pt, 加粗, 居中】 → main_title
2. [1] "朱继东  教授" 【格式：宋体, 16pt, 加粗, 居中】 → author_name
3. [2] "（作者为中国社会科学院马克思主义研究院研究员）" 【格式：仿宋, 14pt】 → affiliation
4. [3] "来源：《马克思主义研究》2024年第1期" 【格式：仿宋, 14pt, 居中】 → source
5. [4] "2024年01月15日  第09版" 【格式：仿宋, 14pt】 → publish_date
6. [5] "党的二十大报告指出，中国式现代化是中国共产党领导的社会主义现代化..." 【格式：仿宋, 14pt, 首行缩进】 → body_text
7. [6] "一、深刻认识重大意义" 【格式：黑体, 16pt, 加粗, 蓝色文字】 → section_header
8. [7] "朱继东，中国社会科学院马克思主义研究院研究员，博士生导师，主要从事马克思主义理论研究..." 【格式：仿宋, 14pt, 首行缩进】 → author_note
9. [8] "理论文章：习近平文化思想的方法论意蕴 《马克思主义研究》" 【格式：仿宋, 14pt】 → tag_label
10. [9] "（2024年01月15日  第09版）" 【格式：仿宋, 14pt】 → edition

## 判断原则

1. **格式+内容联合判断**：同样的文本不同格式→不同类型。如"朱继东"居中加粗=author_name，左对齐普通格式=body_text片段
2. **位置很重要**：第一段通常是main_title；文档头部短文本优先考虑作者/来源
3. **正文不会很短**：少于80字的段落一般不是body_text
4. **颜色是强信号**：蓝色→section_header，红色→article_title
5. **括号是关键信号**：括号开头→affiliation或edition
6. **中文姓名2-4字**：单独出现+加粗/居中→author_name

返回JSON: {{"classifications": [{{"index": 0, "type": "main_title", "confidence": "high"}}]}}
confidence: high/medium/low。只返回JSON。"""

    brief_section = f"\n\n{style_brief}\n" if style_brief else ""
    user_prompt = f"请分析以下文档段落：\n\n{text_block}{brief_section}"

    payload = json.dumps({
        "model": AI_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.1,
        "max_tokens": 8192,
        "response_format": {"type": "json_object"},
    }).encode("utf-8")

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {AI_API_KEY}",
    }

    url = f"{AI_BASE_URL.rstrip('/')}/v1/chat/completions"
    req = urllib.request.Request(url, data=payload, headers=headers, method="POST")

    try:
        with urllib.request.urlopen(req, timeout=180) as resp:
            response = json.loads(resp.read().decode("utf-8"))
        content = response["choices"][0]["message"]["content"]
        ai_result = _parse_ai_json(content)
    except Exception as e:
        raise RuntimeError(f"AI classification failed: {e}")

    # ── Map AI results to ContentParagraph objects ──
    classified_map: dict[int, tuple[str, str]] = {}
    for c in ai_result.get("classifications", []):
        idx = c["index"]
        ptype = c.get("type", "other")
        conf_str = c.get("confidence", "low")
        classified_map[idx] = (ptype, conf_str)

    result: list[ContentParagraph] = []
    for item in paragraphs:
        idx = item["index"]
        ptype_str, conf_str = classified_map.get(idx, ("other", "low"))
        try:
            ptype = ParagraphType(ptype_str)
        except ValueError:
            ptype = ParagraphType.other
            conf_str = "low"
        try:
            conf = ConfidenceLevel(conf_str)
        except ValueError:
            conf = ConfidenceLevel.low

        cp = ContentParagraph(
            index=idx,
            text=item.get("text", ""),
            para_type=ptype,
            confidence=conf,
            has_image=item.get("has_image", False),
            image_count=item.get("image_count", 0),
            images=item.get("images", []),
            is_table=item.get("is_table", False),
            table_data=item.get("table_data"),
            runs=item.get("runs", []),
        )
        result.append(cp)

    return result


def _parse_ai_json(content: str) -> dict:
    """Parse JSON from AI response, handling common formatting issues."""
    import re

    # Strip markdown code fences if present
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content)

    # Remove trailing commas before closing brackets/braces (common LLM mistake)
    content = re.sub(r",\s*([}\]])", r"\1", content)

    return json.loads(content)


# ── Post-Processing ───────────────────────────────────────

def _post_process(paragraphs: list[ContentParagraph]) -> list[ContentParagraph]:
    """
    Post-process AI-classified paragraphs.

    Pure-AI mode: classification is the LLM's responsibility. This function
    only enforces structural invariants the model is allowed to violate
    (e.g. every document must have a main_title).
    """
    # Ensure at least one main_title exists; promote the first plausible
    # paragraph if the model produced none.
    has_main_title = any(cp.para_type == ParagraphType.main_title for cp in paragraphs)
    if not has_main_title:
        for cp in paragraphs:
            if cp.text and cp.para_type not in (ParagraphType.image, ParagraphType.other):
                cp.para_type = ParagraphType.main_title
                cp.confidence = ConfidenceLevel.medium
                break

    return paragraphs


def _split_articles(paragraphs: list[ContentParagraph]) -> list[ArticleContent]:
    """
    Split paragraphs into articles.
    An article boundary is detected when:
    - A new main_title appears (after the first)
    - A new article_title appears that looks like a major heading
    """
    articles = []
    current_paras = []
    article_index = 0

    for cp in paragraphs:
        # Start new article on main_title (except first) or clear article_title break
        if cp.para_type == ParagraphType.main_title and current_paras:
            articles.append(ArticleContent(
                article_index=article_index,
                paragraphs=current_paras,
            ))
            article_index += 1
            current_paras = [cp]
        else:
            current_paras.append(cp)

    # Don't forget the last article
    if current_paras:
        articles.append(ArticleContent(
            article_index=article_index,
            paragraphs=current_paras,
        ))

    # If only one article, that's fine
    if not articles:
        articles.append(ArticleContent(
            article_index=0,
            paragraphs=paragraphs,
        ))

    return articles


def _collect_warnings(doc: ContentDocument) -> list[str]:
    """Collect warnings about the document structure."""
    warnings = []

    for article in doc.articles:
        has_title = any(p.para_type == ParagraphType.main_title for p in article.paragraphs)
        has_body = any(p.para_type == ParagraphType.body_text for p in article.paragraphs)
        has_author = any(p.para_type == ParagraphType.author_name for p in article.paragraphs)

        if not has_title:
            warnings.append(f"Article {article.article_index}: 未找到标题，将使用文件名")
        if not has_body:
            warnings.append(f"Article {article.article_index}: 未找到正文内容")
        if not has_author:
            warnings.append(f"Article {article.article_index}: 未识别到作者信息")

    return warnings
