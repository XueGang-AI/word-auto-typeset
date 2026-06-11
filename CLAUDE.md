# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Backend — requires Python 3.11 (installed at /opt/homebrew/bin/python3.11)
/opt/homebrew/bin/python3.11 -m uvicorn app.main:app --host 0.0.0.0 --port 8765

# Dev mode (auto-reload)
/opt/homebrew/bin/python3.11 -m uvicorn app.main:app --host 0.0.0.0 --port 8765 --reload

# One-click start (checks Python, installs deps, builds frontend, starts server)
./start.sh

# Frontend dev server (port 5173, proxies /api to :8765)
cd frontend && npm run dev

# Frontend production build → frontend/dist/ (served by FastAPI as static files)
cd frontend && npm run build

# Type-check frontend
cd frontend && npx vue-tsc --noEmit

# Run pipeline directly (bypass HTTP — prefer this for large files >500KB)
python3.11 -c "
from pathlib import Path
from docx import Document
from app.services.template_service import analyze_template
from app.services.content_parser import parse_content
from app.services.ai_recognizer import recognize_structure
from app.services.renderer import render_document

template_doc = Document('template.docx')
config = analyze_template(template_doc)
paragraphs = parse_content(Path('content.docx'))
content_doc = recognize_structure(paragraphs, 'content.docx')
render_document(content_doc, config, Path('output.docx'), Path('template.docx'))
"

# Legacy CLI (in legacy/ directory)
python legacy/convert_word_to_pdf.py --input-dir ./docs --output-dir ./pdf
python legacy/format_with_template.py --template t.docx --content c.docx
python legacy/word2pdf_web.py
```

## Architecture

### Core Pipeline (the "typesetting" flow)

```
Template .docx  ──► template_service.analyze_template()  ──► TemplateConfig (14 style profiles + page settings + block rules)
Content  .docx  ──► content_parser.parse_content()       ──► list[dict] (paragraphs + runs + extracted image bytes + tables)
                         └──► ai_recognizer.recognize_structure()  ──► ContentDocument (typed paragraphs with confidence scores)
                                └──► renderer.render_document()     ──► output .docx (template styles applied, images embedded)
```

Each stage is independently callable. The pipeline is used by both the single-file endpoint (`/api/typeset/single/download`) and the batch processor (`batch_service.py`).

### Template Analysis (template_service.py)

**Scoring-based style extraction** (not simple label inference). The analyzer:
1. Collects ALL unique formatting patterns (font, size, bold, color, alignment, indent) with statistics (count, text samples, position)
2. Scores each pattern against each ParagraphType using heuristics (e.g., body_text: indent>0 and not centered → high score; article_title: CENTER + red color → high score)
3. Assigns the BEST pattern per label, ensuring no overwrites
4. Fills sensible defaults for missing types (author_name→CENTER, source→RIGHT, etc.) — NOT blind copies of body_text

Key detail: `_make_pattern_key()` includes **color** in the dedup key, so CENTER+red vs CENTER+blue produce distinct style profiles.

### Content Parsing (content_parser.py)

Parses `.docx` body elements (`w:p` and `w:tbl`), extracting:
- Paragraph text, run-level formatting (font, size, bold, color)
- **Image data**: extracts raw image bytes + dimensions (EMU) + content type from `w:drawing`/`wp:inline`/`a:blip` elements via relationship lookups
- Table structure (rows × cells)
- Skips empty paragraphs, ads, hyperlink-only placeholders

Image data flows through `ContentParagraph.images` field → ai_recognizer preserves it → renderer embeds via `run.add_picture()`.

### AI Recognition (ai_recognizer.py)

Pure-AI mode — classification is delegated entirely to the configured
OpenAI-compatible LLM (DeepSeek / OpenAI / others). The service requires
`AI_ENABLED=true` and `AI_API_KEY`; if either is missing it raises
`RuntimeError`. There is no rule-based fallback; classification failures
propagate to the caller so the operator can act on them.

The prompt is a few-shot system message with ten `(text + format) → type`
examples; format metadata (font, size, bold, color, alignment, indent) is
serialized alongside the text. JSON is requested via
`response_format: {type: json_object}`, with markdown-fence and
trailing-comma cleanup in the parser. Post-processing enforces only the
`main_title` invariant.

All types and confidence levels in `app/schemas/models.py`
(`ParagraphType` enum — 17 types, `ConfidenceLevel` — high/medium/low).

### Renderer (renderer.py)

Strategy:
1. Start from template `.docx` copy (preserves headers/footers/section properties), clear existing content
2. Apply page settings (margins, page size) from template config
3. For each paragraph: look up its `para_type` in template styles → apply font, size, bold, color, alignment, line_spacing, indent, space_before/after
4. **Body text** preserves original run-level bold (for content emphasis like book titles)
5. **Images**: embedded via `run.add_picture(BytesIO(img_bytes), width=Emu(w), height=Emu(h))`
6. Tables rendered with `doc.add_table()`, cells inherit body_text font

### Template Storage

Disk-based under `data/templates/{template_id}/`:
- `template.docx` — original uploaded file
- `template_config.json` — extracted `TemplateConfig` (styles, page settings, block rules)

### Batch Processing (batch_service.py)

In-memory job store (`_jobs` dict), `ThreadPoolExecutor` with `MAX_CONCURRENT_TASKS` (default 4). Each file: parse → recognize → render → save log. Results aggregated into `result.zip` + `report.xlsx`. Progress via `GET /api/typeset/batch/{id}/progress`.

A startup hook in `app/main.py` calls `cleanup_dir(OUTPUT_DIR, BATCH_OUTPUT_MAX_AGE_HOURS)` (default 24h) so completed batch output directories don't accumulate indefinitely.

### API Response Pattern

All endpoints return `APIResponse(success=bool, message=str, data=Any)`. File downloads use `Response(content=bytes)` — NOT `FileResponse` — because `FileResponse` references disk paths that may be cleaned up before Starlette's async send completes.

### Static Frontend Serving

FastAPI mounts `frontend/dist/` at `/`. SPA router handles `/templates`, `/typeset`, `/batch`, `/convert`. API routes prefixed `/api/` take precedence.

## Key Design Decisions

- **No database** — templates and batch results stored as files on disk. Batch jobs lost on restart.
- **python-docx requires BytesIO, not raw bytes** — `Document(io.BytesIO(data))` for uploads.
- **HTTP header encoding** — Chinese filenames use RFC 5987: `filename*=UTF-8''url_encoded_name`.
- **System Python is 3.9** (no `X | None` syntax) but project requires **3.10+** — use `/opt/homebrew/bin/python3.11`.
- **Original scripts preserved in `legacy/`** — not duplicated in project root.

## Known Issues

- HTTP upload of files >800KB can cause uvicorn to hang/crash. Use the direct Python pipeline for large files.
- Rule-based classifier may misclassify unformatted short text (e.g., author names without bold/centering). Enable AI mode for better accuracy.
- LibreOffice required for Word→PDF conversion (`/api/convert/word-to-pdf`).
