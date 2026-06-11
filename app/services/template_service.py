"""
Template analysis service: extracts style profiles from a template Word document
and saves/loads template_config.json.
"""

from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from app.config import TEMPLATES_DIR
from app.schemas.models import (
    BlockRule,
    PageSettings,
    ParagraphType,
    StyleProfile,
    TemplateConfig,
    TemplateListItem,
)
from app.utils.word_utils import (
    emu_to_pt,
    get_paragraph_format_info,
    get_run_font_info,
    is_affiliation_text,
    is_red_color,
    looks_like_date,
    looks_like_edition,
)


# ── Template Storage ──────────────────────────────────────

def _template_dir(template_id: str) -> Path:
    """Get the directory for a specific template."""
    d = TEMPLATES_DIR / template_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def _config_path(template_id: str) -> Path:
    return _template_dir(template_id) / "template_config.json"


def _docx_path(template_id: str) -> Path:
    return _template_dir(template_id) / "template.docx"


# ── CRUD ──────────────────────────────────────────────────

def list_templates() -> list[TemplateListItem]:
    """List all saved templates."""
    items: list[TemplateListItem] = []
    if not TEMPLATES_DIR.exists():
        return items
    for d in sorted(TEMPLATES_DIR.iterdir(), reverse=True):
        if not d.is_dir():
            continue
        cfg = _config_path(d.name)
        if cfg.exists():
            try:
                data = json.loads(cfg.read_text(encoding="utf-8"))
                items.append(
                    TemplateListItem(
                        template_id=data.get("template_id", d.name),
                        template_name=data.get("template_name", d.name),
                        created_at=data.get("created_at", ""),
                        updated_at=data.get("updated_at", ""),
                    )
                )
            except Exception:
                pass
    return items


def get_template(template_id: str) -> TemplateConfig | None:
    """Load a template config by ID."""
    cp = _config_path(template_id)
    if not cp.exists():
        return None
    try:
        data = json.loads(cp.read_text(encoding="utf-8"))
        return TemplateConfig(**data)
    except Exception:
        return None


def delete_template(template_id: str) -> bool:
    """Delete a template and its files."""
    d = _template_dir(template_id)
    if d.exists():
        shutil.rmtree(d)
        return True
    return False


def save_template(template_data: bytes, original_name: str) -> TemplateConfig:
    """Analyze a template Word file and save its config."""
    import io
    doc = Document(io.BytesIO(template_data))
    return analyze_and_save(doc, original_name, template_data)


def analyze_and_save(doc: Document, original_name: str, raw_bytes: bytes) -> TemplateConfig:
    """Analyze template document and persist."""
    config = analyze_template(doc, original_name)

    # Save files
    tdir = _template_dir(config.template_id)
    tdir.mkdir(parents=True, exist_ok=True)

    # Save template .docx
    _docx_path(config.template_id).write_bytes(raw_bytes)

    # Save config JSON
    _config_path(config.template_id).write_text(
        config.model_dump_json(indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return config


# ── Analysis ──────────────────────────────────────────────

def _normalize_alignment(align_str: str | None) -> str | None:
    """Normalize alignment to clean values: LEFT, CENTER, RIGHT, JUSTIFY."""
    if not align_str:
        return None
    upper = align_str.upper()
    if "CENTER" in upper:
        return "CENTER"
    if "LEFT" in upper:
        return "LEFT"
    if "RIGHT" in upper:
        return "RIGHT"
    if "JUSTIFY" in upper:
        return "JUSTIFY"
    # Try numeric
    try:
        val = int(align_str)
        mapping = {0: "LEFT", 1: "CENTER", 2: "RIGHT", 3: "JUSTIFY"}
        return mapping.get(val)
    except (ValueError, TypeError):
        pass
    return None


def analyze_template(doc: Document, name: str = "") -> TemplateConfig:
    """
    Dynamically analyze a template Word document and extract style profiles.

    Strategy:
    1. Collect ALL formatting patterns with their statistics (count, samples)
    2. Score each pattern against each ParagraphType using heuristics
    3. Assign the BEST pattern per label (no overwrites)
    4. Fill sensible defaults for missing styles
    """
    config = TemplateConfig(template_name=name or "未命名模板")

    # ── Page settings ──
    if doc.sections:
        sec = doc.sections[0]
        config.page_settings = PageSettings(
            width_emu=sec.page_width,
            height_emu=sec.page_height,
            top_margin_emu=sec.top_margin,
            bottom_margin_emu=sec.bottom_margin,
            left_margin_emu=sec.left_margin,
            right_margin_emu=sec.right_margin,
        )

    # ── Collect ALL formatting patterns with statistics ──
    patterns: dict[str, dict] = {}  # key → {style, texts[], count, first_pos}

    for pos, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text or not para.runs:
            continue

        first_run = para.runs[0]
        run_info = get_run_font_info(first_run)
        pf_info = get_paragraph_format_info(para)

        # Normalize alignment
        if pf_info.get("alignment"):
            pf_info["alignment"] = _normalize_alignment(pf_info["alignment"])

        # Build a pattern key for deduplication (include color now)
        pattern_key = _make_pattern_key(run_info, pf_info)

        if pattern_key not in patterns:
            style_dict = {**run_info, **pf_info}
            style_dict["line_spacing"] = pf_info.get("line_spacing", 1.5)
            patterns[pattern_key] = {
                "style": style_dict,
                "texts": [],
                "count": 0,
                "first_pos": pos,
            }
        patterns[pattern_key]["count"] += 1
        patterns[pattern_key]["texts"].append(text[:80])

    # ── Score each pattern against each label ──
    best_for_label: dict[str, dict] = {}
    label_scores: dict[str, dict] = {}  # label → {key: score}

    for key, pat in patterns.items():
        style = pat["style"]
        texts = pat["texts"]
        count = pat["count"]
        pos = pat["first_pos"]
        sample_text = texts[0]
        font_size = style.get("font_size_pt", 12)
        is_bold = style.get("bold", False)
        alignment = style.get("alignment")
        indent_emu = style.get("first_line_indent_emu") or 0
        color = style.get("color")
        is_centered = alignment == "CENTER"
        text_len = len(sample_text)

        scores = {}

        # main_title: first position, large font, centered
        score = 0
        if pos == 0:
            score += 50
        if is_centered:
            score += 20
        if font_size and font_size >= 16:
            score += 20
        if font_size and font_size > 14:
            score += 10
        scores["main_title"] = score

        # article_title: centered, possibly colored (red), medium font, NOT bold body-text
        score = 0
        if is_centered:
            score += 30
        if color and is_red_color(color):
            score += 25
        if not is_bold:
            score += 15  # article titles tend not to be bold
        if 10 <= text_len <= 80:
            score += 15
        if font_size and font_size >= 14:
            score += 10
        # Penalize if it looks like body text
        if indent_emu and indent_emu > 200000:
            score -= 30
        scores["article_title"] = score

        # section_header: bold, not centered, medium length
        score = 0
        if is_bold:
            score += 30
        if not is_centered:
            score += 15
        if 5 <= text_len <= 60:
            score += 15
        if indent_emu and indent_emu > 0:
            score += 10  # section headers sometimes have indent
        if text_len > 200:
            score -= 40  # definitely not a section header
        scores["section_header"] = score

        # body_text: has indent, NOT centered, NOT bold (typically), longer text
        score = 0
        if indent_emu and indent_emu > 100000:
            score += 30
        if not is_centered:
            score += 20
        if not is_bold:
            score += 20
        if text_len > 80:
            score += 20
        if count >= 3:
            score += 15  # body text should be the most frequent pattern
        # Penalize if centered (not body text)
        if is_centered:
            score -= 40
        # Penalize very large indent (like source line indent of 308pt)
        if indent_emu and indent_emu > 1000000:
            score -= 20
        scores["body_text"] = score

        # author_name: centered, short (2-4 chars), not bold
        score = 0
        if is_centered:
            score += 25
        if 2 <= text_len <= 6:
            score += 30
        if not is_bold:
            score += 15
        if indent_emu and indent_emu > 100000:
            score -= 30  # author names don't have indent
        scores["author_name"] = score

        # source: often has large indent or RIGHT alignment
        score = 0
        if indent_emu and indent_emu > 1000000:
            score += 35  # very large indent → source line
        if alignment == "RIGHT":
            score += 30
        if text_len < 30:
            score += 10
        if "来源" in sample_text or "选自" in sample_text:
            score += 30
        scores["source"] = score

        # publish_date: date-like pattern
        score = 0
        if looks_like_date(sample_text):
            score += 50
        scores["publish_date"] = score

        # edition: edition-like pattern
        score = 0
        if looks_like_edition(sample_text):
            score += 50
        scores["edition"] = score

        # editor_note: starts with 编者按
        score = 0
        if sample_text.startswith("编者按") or sample_text.startswith("【编者按】"):
            score += 50
        scores["editor_note"] = score

        # affiliation: author unit, typically in （）
        score = 0
        if is_affiliation_text(sample_text):
            score += 50
        scores["affiliation"] = score

        # tag_label: short centered text, possibly colored, pre-body
        score = 0
        if is_centered and text_len < 30:
            score += 25
        if color:
            score += 15
        if not is_bold:
            score += 10
        if text_len < 8:
            score += 10
        scores["tag_label"] = score

        # image_caption: short text after body content
        score = 0
        if text_len < 40 and not is_centered and not is_bold:
            score += 10
        scores["image_caption"] = score

        # Record scores
        label_scores[key] = scores

    # ── Assign best pattern per label (greedy, each pattern → at most one label) ──
    assigned_patterns: set[str] = set()

    # First pass: assign high-confidence matches (score >= 50)
    # NOTE: source and author_note are NOT assigned from template patterns —
    # they get semantic defaults from _ensure_default_styles() instead.
    # NOTE: affiliation is NOT assigned from template patterns —
    # it gets semantic defaults from _ensure_default_styles() instead
    # (template affiliation lines often have unusual indents that shouldn't
    # propagate to content documents).
    for label in [
        "main_title", "body_text", "article_title", "section_header",
        "author_name", "publish_date", "edition", "editor_note",
        "tag_label", "image_caption",
    ]:
        best_key = None
        best_score = 30  # minimum threshold
        for key, scores in label_scores.items():
            if key in assigned_patterns:
                continue
            s = scores.get(label, 0)
            if s > best_score:
                best_score = s
                best_key = key
        if best_key:
            assigned_patterns.add(best_key)
            best_for_label[label] = patterns[best_key]["style"]

    # Second pass: for labels with very low scores, check unassigned patterns
    # NOTE: author_note and footer are NOT assigned from template patterns —
    # they get semantic defaults from _ensure_default_styles() instead.
    # NOTE: affiliation is also excluded — use semantic defaults
    for label in [
        "author_name", "publish_date", "edition", "editor_note",
        "tag_label", "image_caption", "footer",
    ]:
        if label in best_for_label:
            continue
        best_key = None
        best_score = 15
        for key, scores in label_scores.items():
            if key in assigned_patterns:
                continue
            s = scores.get(label, 0)
            if s > best_score:
                best_score = s
                best_key = key
        if best_key:
            assigned_patterns.add(best_key)
            best_for_label[label] = patterns[best_key]["style"]

    # ── Build style profiles from assigned patterns ──
    for label, style_dict in best_for_label.items():
        profile = StyleProfile(
            font_name=style_dict.get("font_name"),
            font_size_pt=style_dict.get("font_size_pt"),
            font_size_emu=style_dict.get("font_size_emu"),
            font_bold=style_dict.get("bold"),
            font_color=style_dict.get("color"),
            line_spacing=style_dict.get("line_spacing", 1.5),
            space_before_pt=style_dict.get("space_before_pt"),
            space_after_pt=style_dict.get("space_after_pt"),
            first_line_indent_emu=style_dict.get("first_line_indent_emu"),
            alignment=_normalize_alignment(style_dict.get("alignment")),
        )
        config.styles[label] = profile

    # ── Fill defaults for missing styles (smarter: use heuristics, not blind copy) ──
    _ensure_default_styles(config)

    # ── Set block rules ──
    config.block_rules = {
        "main_title": BlockRule.required,
        "body_text": BlockRule.required,
        "article_title": BlockRule.optional,
        "section_header": BlockRule.optional,
        "author_name": BlockRule.optional,
        "author_note": BlockRule.optional,
        "source": BlockRule.optional,
        "publish_date": BlockRule.optional,
        "edition": BlockRule.optional,
        "editor_note": BlockRule.optional,
        "affiliation": BlockRule.optional,
        "tag_label": BlockRule.optional,
        "image": BlockRule.optional,
        "image_caption": BlockRule.optional,
        "footer": BlockRule.skip,
        "other": BlockRule.skip,
    }

    return config


# ── Helpers ───────────────────────────────────────────────

def _make_pattern_key(run_info: dict, pf_info: dict) -> str:
    """Create a unique key for a formatting pattern (includes color)."""
    parts = [
        run_info.get("font_name", ""),
        str(run_info.get("font_size_pt", "")),
        str(run_info.get("bold", "")),
        str(run_info.get("color", "")),
        str(pf_info.get("alignment", "")),
        str(pf_info.get("first_line_indent_emu", "")),
    ]
    return "|".join(parts)


def _infer_style_label(para, run_info: dict, pf_info: dict, position: int) -> str:
    """Infer the style label from formatting characteristics."""
    text = para.text.strip()
    font_size = run_info.get("font_size_pt", 12)
    is_bold = run_info.get("bold", False)
    alignment = pf_info.get("alignment", "")
    indent = pf_info.get("first_line_indent_emu")

    # Position 0 → main_title
    if position == 0:
        return "main_title"

    # Large font + bold → article_title
    if font_size and font_size >= 16 and is_bold:
        return "article_title"

    # Centered text → article_title
    if alignment and "CENTER" in str(alignment):
        return "article_title"

    # Bold but not largest → section_header
    if is_bold:
        return "section_header"

    # Has first-line indent → body_text
    if indent and indent > 100000:
        return "body_text"

    # Short text → could be author or label
    if len(text) <= 6:
        return "author_name"

    # Default → body_text
    return "body_text"


def _ensure_default_styles(config: TemplateConfig) -> None:
    """Ensure all required style types have sensible defaults, inheriting spacing from template."""
    # Get reference styles for fallback
    main_title_style = config.styles.get("main_title")
    body_style = config.styles.get("body_text")
    article_style = config.styles.get("article_title")

    # Base font — use body_text's font if available, otherwise 仿宋 14pt
    base_font = "仿宋"
    base_size = 14.0
    if body_style:
        base_font = body_style.font_name or base_font
        base_size = body_style.font_size_pt or base_size
    elif main_title_style:
        base_font = main_title_style.font_name or base_font

    # Inherit spacing from existing styles if available
    body_space_before = body_style.space_before_pt if body_style else None
    body_space_after = body_style.space_after_pt if body_style else None
    # If body doesn't have spacing, try main_title
    if body_space_before is None and main_title_style:
        body_space_before = main_title_style.space_before_pt
        body_space_after = main_title_style.space_after_pt

    # CENTER-aligned base (inherit from article_title if it's centered)
    center_base = article_style if article_style and article_style.alignment == "CENTER" else None

    # Smart defaults for each type
    defaults = {
        "main_title": StyleProfile(
            font_name=main_title_style.font_name if main_title_style else "宋体",
            font_size_pt=main_title_style.font_size_pt if main_title_style else 16.0,
            font_bold=True,
            line_spacing=1.5,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "article_title": StyleProfile(
            font_name=center_base.font_name if center_base else base_font,
            font_size_pt=center_base.font_size_pt if center_base else base_size,
            font_bold=False,
            line_spacing=1.5,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "section_header": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=True,
            line_spacing=1.5,
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "body_text": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            first_line_indent_emu=355600,  # 28pt ~ 2 chars at 14pt
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "author_name": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "author_note": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            first_line_indent_emu=355600,  # 28pt — body-text indent for bio paragraphs
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "source": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            alignment="CENTER",
            first_line_indent_emu=0,
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "publish_date": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "edition": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "editor_note": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            first_line_indent_emu=355600,
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "affiliation": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=1.5,
            alignment="CENTER",
            first_line_indent_emu=0,
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "tag_label": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size - 2 if base_size else 12,
            font_bold=False,
            line_spacing=1.5,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "image_caption": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size - 4 if base_size else 11,
            font_bold=False,
            line_spacing=1.2,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
        ),
        "image": StyleProfile(
            font_name=base_font,
            font_size_pt=base_size,
            font_bold=False,
            line_spacing=body_style.line_spacing if body_style else 1.5,
            alignment="CENTER",
            space_before_pt=body_space_before,
            space_after_pt=body_space_after,
            first_line_indent_emu=0,
        ),
        "footer": StyleProfile(
            font_name=base_font,
            font_size_pt=10,
            font_bold=False,
            line_spacing=1.0,
            alignment="CENTER",
        ),
    }

    for label, default_profile in defaults.items():
        if label not in config.styles:
            config.styles[label] = default_profile
