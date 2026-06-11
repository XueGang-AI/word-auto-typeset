"""
FastAPI application entry point for the Word Auto-Typesetting System.

Provides:
- Template management (upload, analyze, configure)
- Single-file typesetting (content parsing + AI recognition + rendering)
- Batch typesetting (ZIP upload → concurrent processing → ZIP + report.xlsx)
- Word → PDF batch conversion (preserved from original project)
"""

from __future__ import annotations

import logging
import sys
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.config import (
    BATCH_OUTPUT_MAX_AGE_HOURS,
    CORS_ORIGINS,
    HOST,
    OUTPUT_DIR,
    PORT,
    STATIC_DIR,
)
from app.routers import convert, template, typeset
from app.utils.file_utils import cleanup_dir

# Ensure app package is importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan: clean up old batch outputs on startup."""
    try:
        removed = cleanup_dir(OUTPUT_DIR, BATCH_OUTPUT_MAX_AGE_HOURS)
        if removed:
            logger.info(
                "Startup cleanup: removed %d stale batch entries from %s "
                "(older than %dh)",
                removed, OUTPUT_DIR, BATCH_OUTPUT_MAX_AGE_HOURS,
            )
    except Exception as e:
        logger.warning("Startup batch cleanup failed: %s", e)
    yield


app = FastAPI(
    title="Word 自动套版系统",
    description="模板管理 + AI内容识别 + 批量排版 + Word转PDF",
    version="2.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    lifespan=lifespan,
)

# ── CORS ──────────────────────────────────────────────────
app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS.split(",") if CORS_ORIGINS != "*" else ["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Routers ───────────────────────────────────────────────
app.include_router(template.router)
app.include_router(typeset.router)
app.include_router(convert.router)


# ── Health Check ──────────────────────────────────────────
@app.get("/api/health")
async def health_check():
    """Health check endpoint."""
    return {
        "status": "ok",
        "version": "2.0.0",
        "modules": {
            "template_management": True,
            "ai_recognition": True,
            "batch_typesetting": True,
            "word_to_pdf": True,
        },
    }


# ── Serve Frontend (production) ───────────────────────────
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")


# ── CLI Entry ─────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app.main:app",
        host=HOST,
        port=PORT,
        reload=True,
        log_level="info",
    )
