"""
Shared Word document utilities used by the parser, AI recognizer, and
template analyzer. Helpers that were only consumed by the removed
rule-based classifier have been pruned.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def is_empty_paragraph(para) -> bool:
    """Check if paragraph has no text and no images."""
    if para.text.strip():
        return False
    for run in para.runs:
        if run._element.findall(qn("w:drawing")):
            return False
        if run._element.findall(qn("wp:inline")):
            return False
        if run._element.findall(qn("wp:anchor")):
            return False
    if para._element.findall(qn("w:drawing")):
        return False
    return True


def has_image(para) -> bool:
    """Check if paragraph contains an image."""
    for run in para.runs:
        if run._element.findall(qn("w:drawing")):
            return True
        if run._element.findall(qn("wp:inline")):
            return True
        if run._element.findall(qn("wp:anchor")):
            return True
    if para._element.findall(qn("w:drawing")):
        return True
    return False


def count_images(para) -> int:
    """Count images in a paragraph."""
    count = 0
    for run in para.runs:
        count += len(run._element.findall(qn("w:drawing")))
        count += len(run._element.findall(qn("wp:inline")))
        count += len(run._element.findall(qn("wp:anchor")))
    count += len(para._element.findall(qn("w:drawing")))
    return count


def get_run_font_info(run) -> dict:
    """Extract font information from a run."""
    info = {}
    if run.font.name:
        info["font_name"] = run.font.name
    if run.font.size:
        info["font_size_pt"] = run.font.size.pt
        info["font_size_emu"] = run.font.size
    if run.font.bold is not None:
        info["bold"] = run.font.bold
    if run.font.italic is not None:
        info["italic"] = run.font.italic
    if run.font.color and run.font.color.rgb:
        info["color"] = str(run.font.color.rgb)
    return info


def get_paragraph_format_info(para) -> dict:
    """Extract paragraph formatting."""
    pf = para.paragraph_format
    info = {}
    if pf.line_spacing:
        info["line_spacing"] = pf.line_spacing
    if pf.space_before:
        info["space_before_pt"] = pf.space_before.pt
    if pf.space_after:
        info["space_after_pt"] = pf.space_after.pt
    if pf.first_line_indent:
        info["first_line_indent_emu"] = pf.first_line_indent
    if pf.alignment is not None:
        info["alignment"] = str(pf.alignment)
    return info


# ── Color heuristics (used by template analyzer + AI format description) ──

def is_blue_color(rgb: str | None) -> bool:
    """Check if color is in the blue range."""
    if rgb is None:
        return False
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return b > r + 40 and b > g + 20


def is_red_color(rgb: str | None) -> bool:
    """Check if color is in the red range."""
    if rgb is None:
        return False
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return r > g + 40 and r > b + 40


# ── Text-pattern heuristics (used by template analyzer) ─────

def is_affiliation_text(text: str) -> bool:
    """Check if text is an affiliation/author unit note."""
    t = text.strip()
    return t.startswith("（") and any(
        kw in t for kw in ("作者为", "执笔人", "作者单位", "作者系")
    )


def looks_like_date(text: str) -> bool:
    """Check if text looks like a publish date (possibly with edition)."""
    t = text.strip()
    return bool(re.match(
        r"^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日"
        r"(\s+第[0-9０-９一二三四五六七八九十]+版)?$",
        t,
    ))


def looks_like_edition(text: str) -> bool:
    """Check if text looks like edition info."""
    return bool(re.search(r"(第[0-9０-９一二三四五六七八九十]+版|版次)", text))


def looks_like_ad(text: str) -> bool:
    """Simple heuristic to detect ad/spam text."""
    t = text.strip()
    ad_keywords = ["点击领取", "免费领", "加微信", "扫码", "关注公众号", "下载APP"]
    return any(kw in t for kw in ad_keywords)
