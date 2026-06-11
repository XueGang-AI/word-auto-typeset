"""
Word → PDF conversion API endpoints.
Preserves the existing Word-to-PDF batch conversion functionality.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import uuid
import zipfile
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from starlette.responses import Response

from app.schemas.models import APIResponse
from app.utils.file_utils import safe_filename

router = APIRouter(prefix="/api/convert", tags=["convert"])


def _find_soffice() -> str | None:
    """Find LibreOffice executable."""
    import os
    env_bin = os.environ.get("WORD2PDF_SOFFICE", "")
    if env_bin and Path(env_bin).exists():
        return env_bin
    for name in ("soffice", "soffice.com", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return None


def _convert_one(soffice_bin: str, src: Path, out_dir: Path) -> Path:
    """Convert a single Word file to PDF using LibreOffice."""
    staging = out_dir / "_staging"
    staging.mkdir(parents=True, exist_ok=True)
    safe_src = staging / f"input_{uuid.uuid4().hex}{src.suffix.lower()}"
    shutil.copy2(src, safe_src)

    profile_dir = (out_dir / f"_profile_{uuid.uuid4().hex}").resolve()
    profile_dir.mkdir(parents=True, exist_ok=True)

    proc = subprocess.run(
        [
            soffice_bin,
            "--headless", "--nologo", "--nolockcheck",
            "--nodefault", "--norestore",
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", str(out_dir),
            str(safe_src),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    expected = out_dir / f"{safe_src.stem}.pdf"
    if expected.exists():
        return expected

    # Fallback: try ODT intermediate
    odt_stage = out_dir / f"{safe_src.stem}.odt"
    if odt_stage.exists():
        odt_stage.unlink()
    subprocess.run(
        [
            soffice_bin,
            "--headless", "--nologo", "--nolockcheck",
            "--nodefault", "--norestore",
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--convert-to", "odt",
            "--outdir", str(out_dir),
            str(safe_src),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if odt_stage.exists():
        subprocess.run(
            [
                soffice_bin,
                "--headless", "--nologo", "--nolockcheck",
                "--nodefault", "--norestore",
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--convert-to", "pdf:writer_pdf_Export",
                "--outdir", str(out_dir),
                str(odt_stage),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        if expected.exists():
            return expected

    # Try finding any PDF
    candidates = list(out_dir.glob("*.pdf"))
    if len(candidates) == 1:
        return candidates[0]
    if candidates:
        return candidates[0]

    raise RuntimeError(f"转换失败: {src.name}\n{proc.stderr.strip() or proc.stdout.strip()}")


@router.post("/word-to-pdf", response_model=APIResponse)
async def convert_word_to_pdf(
    files: list[UploadFile] = File(...),
    target_names: str = Form(""),
    overwrite: bool = Form(True),
):
    """
    Convert Word files to PDF and return as ZIP.

    This preserves the original batch Word→PDF functionality.
    """
    soffice = _find_soffice()
    if not soffice:
        raise HTTPException(
            status_code=503,
            detail="未找到 LibreOffice。请安装 LibreOffice 或设置 WORD2PDF_SOFFICE 环境变量。",
        )

    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个文件")

    # Parse target names
    name_list = [n.strip() for n in target_names.splitlines() if n.strip()]
    if name_list and len(name_list) != len(files):
        raise HTTPException(
            status_code=400,
            detail=f"目标名字数量 ({len(name_list)}) 与文件数量 ({len(files)}) 不一致",
        )

    with tempfile.TemporaryDirectory(prefix="w2p_") as td:
        root = Path(td)
        src_dir = root / "input"
        out_dir = root / "output"
        src_dir.mkdir()
        out_dir.mkdir()

        # Save uploaded files
        saved = []
        for i, f in enumerate(files):
            if not f.filename:
                continue
            data = await f.read()
            safe = safe_filename(f.filename)
            path = src_dir / safe
            path.write_bytes(data)
            saved.append((path, f.filename))

        if not saved:
            raise HTTPException(status_code=400, detail="没有有效文件")

        # Convert each file
        failures = []
        success = 0
        for i, (src_path, orig_name) in enumerate(saved):
            try:
                pdf_path = _convert_one(soffice, src_path, out_dir)

                # Validate
                with open(pdf_path, "rb") as f:
                    if f.read(5) != b"%PDF-":
                        raise RuntimeError("输出文件不是有效 PDF")

                # Rename
                if name_list and i < len(name_list):
                    target = name_list[i]
                    if not target.lower().endswith(".pdf"):
                        target += ".pdf"
                else:
                    target = f"{src_path.stem}.pdf"

                target_path = out_dir / safe_filename(target)
                shutil.move(str(pdf_path), str(target_path))
                success += 1
            except Exception as e:
                failures.append({"filename": orig_name, "error": str(e)})

        if success == 0 and failures:
            raise HTTPException(
                status_code=500,
                detail=f"所有文件转换失败: {failures[0]['error']}",
            )

        # Create ZIP
        zip_path = root / "word2pdf_result.zip"
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for pdf in sorted(out_dir.glob("*.pdf")):
                zf.write(pdf, arcname=pdf.name)

        # Read ZIP into memory before temp dir cleanup
        zip_bytes = zip_path.read_bytes()

    # Return response after temp dir is safely cleaned up
    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=word2pdf_result.zip",
        },
    )
