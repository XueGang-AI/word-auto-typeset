"""
File handling utilities.
"""

import hashlib
import shutil
import zipfile
from pathlib import Path
from typing import BinaryIO


def safe_filename(name: str) -> str:
    """Sanitize a filename for safe storage."""
    invalid = '<>:"/\\|?*'
    table = str.maketrans({ch: "_" for ch in invalid})
    name = name.translate(table).strip().rstrip(" .")
    return name or "untitled"


def file_hash(filepath: Path, algo: str = "md5") -> str:
    """Compute file hash."""
    h = hashlib.new(algo)
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def save_upload(file_data: bytes, dest: Path) -> Path:
    """Save uploaded file bytes to disk."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(file_data)
    return dest


def create_zip(
    files: list[tuple[str, Path]], output_path: Path, extra_files: list[tuple[str, Path]] | None = None
) -> Path:
    """Create a ZIP file from a list of (arcname, path) tuples."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        seen = set()
        for arcname, src in files:
            if not src.exists():
                continue
            # Avoid duplicate names
            base = arcname
            counter = 1
            while base in seen:
                stem, ext = Path(arcname).stem, Path(arcname).suffix
                base = f"{stem}_{counter}{ext}"
                counter += 1
            seen.add(base)
            zf.write(src, arcname=base)
        if extra_files:
            for arcname, src in extra_files:
                if src.exists():
                    zf.write(src, arcname=arcname)
    return output_path


def read_uploaded_files(files: list[tuple[str, bytes, str]], dest_dir: Path) -> list[Path]:
    """Write multiple uploaded files to a directory. Returns list of saved paths."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for original_name, data, _content_type in files:
        safe = safe_filename(original_name)
        path = dest_dir / safe
        path.write_bytes(data)
        saved.append(path)
    return saved


def cleanup_dir(dir_path: Path, max_age_hours: int = 24) -> int:
    """Remove files and subdirectories older than max_age_hours from dir_path.

    Returns the count of entries removed. Missing dir is a no-op (returns 0).
    Used to keep batch output directories from growing unbounded.
    """
    import time
    now = time.time()
    threshold_seconds = max_age_hours * 3600
    removed = 0
    if not dir_path.exists():
        return 0
    for f in dir_path.iterdir():
        try:
            mtime = f.stat().st_mtime
        except FileNotFoundError:
            continue
        if (now - mtime) <= threshold_seconds:
            continue
        if f.is_file() or f.is_symlink():
            try:
                f.unlink()
                removed += 1
            except FileNotFoundError:
                pass
        elif f.is_dir():
            shutil.rmtree(f, ignore_errors=True)
            removed += 1
    return removed
