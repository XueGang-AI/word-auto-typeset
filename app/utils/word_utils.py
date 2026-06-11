"""
Shared Word document utilities using python-docx.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu, Pt


def open_docx(path: Path) -> Document:
    """Safely open a .docx file."""
    return Document(str(path))


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


def get_first_run_color(para) -> str | None:
    """Get the color of the first non-default run."""
    for run in para.runs:
        if run.font.color and run.font.color.rgb:
            return str(run.font.color.rgb)
    return None


def get_first_run_bold(para) -> bool | None:
    """Get bold status of first run."""
    for run in para.runs:
        if run.font.bold is not None:
            return run.font.bold
    return None


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


# ── Text normalization ─────────────────────────────────────

def normalize_metadata_text(text: str) -> str:
    """Normalize metadata-like text: collapse whitespace, strip parens."""
    t = (text or "").replace("\xa0", " ")
    t = re.sub(r"\s+", " ", t).strip()
    return t.strip("()（）").strip()


# ── Person name / author heading ───────────────────────────

def looks_like_person_name(text: str) -> bool:
    """Check if text looks like a Chinese person name or short author heading."""
    t = normalize_metadata_text(text)

    # Exclude common non-name patterns
    non_name_patterns = {
        "本期导读", "编者按", "发言摘登", "内容提要", "文章摘要",
        "全文", "摘要", "关键词", "参考文献", "作者简介",
        "目录", "前言", "后记", "序言", "跋",
    }
    if t in non_name_patterns:
        return False
    # Exclude patterns like "第一章", "第二节", etc.
    if re.match(r"^第[一二三四五六七八九十百千\d]+[章节条段]", t):
        return False

    # Plain Chinese name, e.g. 朱继东 / 李泽中
    if 2 <= len(t) <= 4 and re.fullmatch(r"[一-鿿]+", t):
        return True

    # Short author heading with title, e.g. "朱继东 教授"
    # Exclude standalone job titles that happen to match the pattern
    standalone_titles = {"助理研究员", "副研究员", "博士生导师", "硕士研究生导师"}
    if t in standalone_titles:
        return False

    title_pattern = r"教授|研究员|博士|导师|博士生导师|助理研究员"
    return bool(re.fullmatch(rf"[一-鿿]{{2,4}}\s*({title_pattern})", t))


# ── Affiliation ────────────────────────────────────────────

def is_affiliation_text(text: str) -> bool:
    """Check if text is an affiliation/author unit note."""
    t = text.strip()
    return t.startswith("（") and any(
        kw in t for kw in ("作者为", "执笔人", "作者单位", "作者系")
    )


# ── Author biography ───────────────────────────────────────

def looks_like_author_bio(text: str) -> bool:
    """Check if text is a biographical author/introduction paragraph."""
    t = normalize_metadata_text(text)
    if len(t) < 30:
        return False
    # Must start with a Chinese name followed by Chinese/comma
    if not re.match(r"^[一-鿿]{2,4}[，,]", t):
        return False

    credential_keywords = (
        "研究员", "教授", "博士", "博士后", "博士生导师", "首席专家",
        "助理研究员", "中国社会科学院", "大学", "学院", "研究院",
    )
    activity_keywords = ("主要从事", "兼任", "主持", "发表", "出版", "论著", "课题", "项目")
    credential_count = sum(1 for kw in credential_keywords if kw in t)
    has_activity = any(kw in t for kw in activity_keywords)
    return credential_count >= 2 or (credential_count >= 1 and has_activity)


# ── Recommendation / link-list ─────────────────────────────

def looks_like_recommendation(text: str) -> bool:
    """Check if text looks like a guide/recommendation/link-list item."""
    t = normalize_metadata_text(text)
    if len(t) < 10:
        return False

    # Contains a newspaper/publication title in book-name quotes
    publication_cue = bool(re.search(
        r"《[^》]*(日报|晚报|时报|报|杂志|周刊|月刊|网|出版社)[^》]*》", t
    ))
    recommendation_cue = any(kw in t for kw in ("理论文章：", "重磅好文", "合集", "连发"))
    colon_title = "：" in t and publication_cue
    return (publication_cue and recommendation_cue) or colon_title


# ── Source ─────────────────────────────────────────────────

def looks_like_source(text: str) -> bool:
    """Check if text looks like a source attribution."""
    t = normalize_metadata_text(text)

    # Explicit source prefixes
    if re.match(r"^(来源|来源：|来源\s*\||摘自|原载|原载于|转自)", t):
        return True

    # Standalone publication title in book-name quotes, e.g. 《山西日报》
    if re.fullmatch(r"《[^》]{2,20}(日报|晚报|时报|报|杂志|周刊|月刊|网|出版社)》", t):
        return True

    # Prefixed quoted source, e.g. "原载于《山西日报》"
    if re.match(r"^(来源|摘自|原载|原载于|转自).*《[^》]+》", t):
        return True

    return False


# ── Date ───────────────────────────────────────────────────

def looks_like_date(text: str) -> bool:
    """Check if text looks like a publish date (possibly with edition)."""
    t = normalize_metadata_text(text)
    # Date only: 2026年06月02日
    # Date + edition: 2026年06月02日 第09版
    return bool(re.match(
        r"^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日"
        r"(\s+第[0-9０-９一二三四五六七八九十]+版)?$",
        t,
    ))


# ── Edition ────────────────────────────────────────────────

def looks_like_edition(text: str) -> bool:
    """Check if text looks like edition info."""
    t = normalize_metadata_text(text)
    return bool(re.search(r"(第[0-9０-９一二三四五六七八九十]+版|版次)", t))


# ── Hyperlink / Ad ─────────────────────────────────────────

def is_hyperlink(run) -> bool:
    """Check if run is part of a hyperlink."""
    # Check parents for hyperlink element
    parent = run._element.getparent()
    while parent is not None:
        if parent.tag == qn("w:hyperlink"):
            return True
        parent = parent.getparent()
    return False


def looks_like_ad(text: str) -> bool:
    """Simple heuristic to detect ad/spam text."""
    t = text.strip()
    ad_keywords = ["点击领取", "免费领", "加微信", "扫码", "关注公众号", "下载APP"]
    return any(kw in t for kw in ad_keywords)


# ── Unit conversion ────────────────────────────────────────

def emu_to_pt(emu: int) -> float:
    """Convert EMU to points."""
    return emu / 12700


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU."""
    return int(pt * 12700)
