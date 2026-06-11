"""
Application configuration.
"""

import os
from pathlib import Path

# ── Load .env file ────────────────────────────────────────
try:
    from dotenv import load_dotenv
    _env_path = Path(__file__).resolve().parent.parent / ".env"
    if _env_path.exists():
        load_dotenv(_env_path)
except ImportError:
    pass

# ── Paths ────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = DATA_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "batch_output"
STATIC_DIR = BASE_DIR / "frontend" / "dist"

for d in [DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ── Database ──────────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", f"sqlite+aiosqlite:///{DATA_DIR / 'app.db'}")

# ── AI / LLM ─────────────────────────────────────────────
AI_ENABLED = os.environ.get("AI_ENABLED", "false").lower() == "true"
AI_PROVIDER = os.environ.get("AI_PROVIDER", "openai")  # openai | deepseek | qwen
AI_API_KEY = os.environ.get("AI_API_KEY", "")
AI_BASE_URL = os.environ.get("AI_BASE_URL", "https://api.openai.com")
AI_MODEL = os.environ.get("AI_MODEL", "gpt-4o-mini")

# ── Conversion ───────────────────────────────────────────
SOFFICE_BIN = os.environ.get("WORD2PDF_SOFFICE", "")
DOCKER_IMAGE = os.environ.get("WORD2PDF_DOCKER_IMAGE", "linuxserver/libreoffice:latest")

# ── Batch ────────────────────────────────────────────────
MAX_CONCURRENT_TASKS = int(os.environ.get("MAX_CONCURRENT_TASKS", "4"))
MAX_IMAGES_PER_DOC = int(os.environ.get("MAX_IMAGES_PER_DOC", "20"))
# How long batch output directories are kept before being cleaned up at startup.
BATCH_OUTPUT_MAX_AGE_HOURS = int(os.environ.get("BATCH_OUTPUT_MAX_AGE_HOURS", "24"))

# ── Server ───────────────────────────────────────────────
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8765"))
CORS_ORIGINS = os.environ.get("CORS_ORIGINS", "*")
