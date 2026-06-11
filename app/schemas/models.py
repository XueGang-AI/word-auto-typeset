"""
Pydantic models for the auto-typesetting system.
"""

from __future__ import annotations

import uuid
from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


# ── Enums ─────────────────────────────────────────────────

class BlockRule(str, Enum):
    required = "required"
    optional = "optional"
    skip = "skip"


class ParagraphType(str, Enum):
    main_title = "main_title"
    article_title = "article_title"
    section_header = "section_header"
    body_text = "body_text"
    author_name = "author_name"
    author_note = "author_note"
    source = "source"
    publish_date = "publish_date"
    edition = "edition"
    editor_note = "editor_note"
    affiliation = "affiliation"
    tag_label = "tag_label"
    image = "image"
    image_caption = "image_caption"
    footer = "footer"
    other = "other"


class TaskStatus(str, Enum):
    pending = "pending"
    processing = "processing"
    completed = "completed"
    failed = "failed"


class ConfidenceLevel(str, Enum):
    high = "high"
    medium = "medium"
    low = "low"


# ── Style Profile ─────────────────────────────────────────

class StyleProfile(BaseModel):
    """A single style configuration extracted from template."""
    font_name: str | None = None
    font_size_pt: float | None = None
    font_size_emu: int | None = None  # EMU for python-docx
    font_bold: bool | None = None
    font_italic: bool | None = None
    font_color: str | None = None
    line_spacing: float = 1.5
    space_before_pt: float | None = None
    space_after_pt: float | None = None
    first_line_indent_emu: int | None = None
    alignment: str | None = None  # LEFT | CENTER | RIGHT | JUSTIFY
    block_rule: BlockRule = BlockRule.required


class PageSettings(BaseModel):
    width_emu: int | None = None
    height_emu: int | None = None
    top_margin_emu: int | None = None
    bottom_margin_emu: int | None = None
    left_margin_emu: int | None = None
    right_margin_emu: int | None = None


# ── Template ──────────────────────────────────────────────

class TemplateConfig(BaseModel):
    """Full template configuration saved as template_config.json."""
    template_id: str = Field(default_factory=lambda: f"tpl_{uuid.uuid4().hex[:8]}")
    template_name: str = ""
    page_settings: PageSettings = Field(default_factory=PageSettings)
    styles: dict[str, StyleProfile] = Field(default_factory=dict)
    # Block rules control which paragraph types to render
    block_rules: dict[str, BlockRule] = Field(default_factory=dict)
    created_at: str = Field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = Field(default_factory=lambda: datetime.now().isoformat())


class TemplateListItem(BaseModel):
    template_id: str
    template_name: str
    created_at: str
    updated_at: str


# ── Content ───────────────────────────────────────────────

class ContentParagraph(BaseModel):
    """A single paragraph from the content document with AI classification."""
    index: int
    text: str = ""
    para_type: ParagraphType = ParagraphType.other
    confidence: ConfidenceLevel = ConfidenceLevel.low
    # Preserved formatting hints from original doc
    has_image: bool = False
    image_count: int = 0
    # Extracted image data: list of {bytes, content_type, width_emu, height_emu}
    images: list[dict[str, Any]] = Field(default_factory=list)
    # Table content
    is_table: bool = False
    table_data: list[list[str]] | None = None
    # Run-level properties preserved
    runs: list[dict[str, Any]] = Field(default_factory=list)


class ArticleContent(BaseModel):
    """A single article's structured content."""
    article_index: int
    paragraphs: list[ContentParagraph] = Field(default_factory=list)
    metadata: dict[str, str] = Field(default_factory=dict)


class ContentDocument(BaseModel):
    """Full content document with AI-recognized structure."""
    content_id: str = Field(default_factory=lambda: f"cnt_{uuid.uuid4().hex[:8]}")
    original_filename: str = ""
    articles: list[ArticleContent] = Field(default_factory=list)
    # Overall confidence assessment
    has_low_confidence: bool = False
    warnings: list[str] = Field(default_factory=list)


# ── Batch Processing ──────────────────────────────────────

class BatchTask(BaseModel):
    task_id: str = Field(default_factory=lambda: f"task_{uuid.uuid4().hex[:8]}")
    batch_id: str
    original_filename: str
    status: TaskStatus = TaskStatus.pending
    error_message: str | None = None
    output_filename: str | None = None
    content_id: str | None = None
    has_low_confidence: bool = False
    warnings: list[str] = Field(default_factory=list)


class BatchJob(BaseModel):
    batch_id: str = Field(default_factory=lambda: f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    template_id: str
    tasks: list[BatchTask] = Field(default_factory=list)
    status: TaskStatus = TaskStatus.pending
    total_count: int = 0
    success_count: int = 0
    failed_count: int = 0
    low_confidence_count: int = 0
    zip_download_url: str | None = None
    report_download_url: str | None = None
    created_at: str = Field(default_factory=lambda: datetime.now().isoformat())


class BatchProgress(BaseModel):
    batch_id: str
    status: TaskStatus
    total: int
    completed: int
    success: int
    failed: int
    low_confidence: int


# ── Conversion (Word → PDF) ───────────────────────────────

class ConvertRequest(BaseModel):
    filenames: list[str] = Field(default_factory=list)
    target_names: list[str] = Field(default_factory=list)
    overwrite: bool = True


class ConvertResult(BaseModel):
    success_count: int
    failed_count: int
    failures: list[dict[str, str]] = Field(default_factory=list)
    zip_download_url: str | None = None


# ── API Response ──────────────────────────────────────────

class APIResponse(BaseModel):
    success: bool
    message: str = ""
    data: Any = None
