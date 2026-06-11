"""
Typesetting API endpoints — single file and batch processing.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import OUTPUT_DIR
from app.schemas.models import APIResponse, BatchJob, TemplateListItem
from app.services.batch_service import (
    create_batch_job,
    get_batch_job,
    get_batch_progress,
)
from app.services.content_parser import parse_content
from app.services.ai_recognizer import recognize_structure
from app.services.renderer import render_document
from app.services.template_service import (
    _docx_path,
    get_template,
    list_templates,
)
from app.utils.file_utils import safe_filename

router = APIRouter(prefix="/api/typeset", tags=["typeset"])


# ── Single File Typesetting ───────────────────────────────

@router.post("/single/download")
async def typeset_single_download(
    template_id: str = Form(...),
    file: UploadFile = File(...),
):
    """
    Typeset a single file and return the formatted .docx as a download.
    """
    from starlette.responses import Response

    template = get_template(template_id)
    if not template:
        raise HTTPException(status_code=404, detail="模板不存在")

    if not file.filename:
        raise HTTPException(status_code=400, detail="文件名不能为空")

    content_bytes = await file.read()
    stem = Path(safe_filename(file.filename)).stem

    with tempfile.TemporaryDirectory(prefix="typeset_") as td:
        root = Path(td)
        content_path = root / safe_filename(file.filename)
        content_path.write_bytes(content_bytes)

        paragraphs = parse_content(content_path)
        if not paragraphs:
            raise HTTPException(status_code=400, detail="内容文档为空")

        content_doc = recognize_structure(paragraphs, file.filename)

        output_path = root / f"{stem}_排版后.docx"
        template_docx = _docx_path(template_id)
        render_document(content_doc, template, output_path, template_docx)

        # Read file into memory BEFORE exiting the temp dir context
        output_bytes = output_path.read_bytes()

    # Return bytes response after temp dir cleanup is safe
    from urllib.parse import quote
    safe_filename_out = f"{stem}_排版后.docx"
    encoded_filename = quote(safe_filename_out)
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": (
                f"attachment; filename=\"{encoded_filename}\"; "
                f"filename*=UTF-8''{encoded_filename}"
            ),
        },
    )


# ── Content Analysis Only (Preview) ────────────────────────

@router.post("/analyze", response_model=APIResponse)
async def analyze_content(file: UploadFile = File(...)):
    """
    Analyze a content Word document and return the recognized structure
    without rendering. Useful for preview/debugging.
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="文件名不能为空")

    content_bytes = await file.read()

    with tempfile.TemporaryDirectory(prefix="analyze_") as td:
        root = Path(td)
        content_path = root / safe_filename(file.filename)
        content_path.write_bytes(content_bytes)

        paragraphs = parse_content(content_path)
        content_doc = recognize_structure(paragraphs, file.filename)

        return APIResponse(
            success=True,
            message="分析完成",
            data=content_doc.model_dump(),
        )


# ── Batch Typesetting ─────────────────────────────────────

@router.post("/batch", response_model=APIResponse)
async def typeset_batch(
    template_id: str = Form(...),
    files: list[UploadFile] = File(...),
):
    """
    Start a batch typesetting job.

    Accepts multiple Word files and a template ID.
    Returns a batch_id for progress tracking.
    """
    template = get_template(template_id)
    if not template:
        raise HTTPException(status_code=404, detail="模板不存在")

    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个文件")

    # Read all files
    filenames = []
    file_data_list = []
    for f in files:
        if f.filename:
            filenames.append(f.filename)
            file_data_list.append(await f.read())

    if not filenames:
        raise HTTPException(status_code=400, detail="没有有效文件")

    template_docx = _docx_path(template_id)

    job = create_batch_job(
        filenames=filenames,
        file_data_list=file_data_list,
        template_id=template_id,
        template_config=template,
        template_docx_path=template_docx if template_docx.exists() else None,
    )

    return APIResponse(
        success=True,
        message=f"批量任务已创建，共 {job.total_count} 个文件",
        data={
            "batch_id": job.batch_id,
            "total_count": job.total_count,
            "status": job.status.value,
        },
    )


@router.get("/batch/{batch_id}/progress", response_model=APIResponse)
async def batch_progress(batch_id: str):
    """Get the progress of a batch job."""
    progress = get_batch_progress(batch_id)
    if not progress:
        raise HTTPException(status_code=404, detail="批量任务不存在")
    return APIResponse(
        success=True,
        message="获取进度成功",
        data=progress.model_dump(),
    )


@router.get("/batch/{batch_id}", response_model=APIResponse)
async def batch_detail(batch_id: str):
    """Get the full detail of a batch job."""
    job = get_batch_job(batch_id)
    if not job:
        raise HTTPException(status_code=404, detail="批量任务不存在")
    return APIResponse(
        success=True,
        message="获取任务详情成功",
        data=job.model_dump(),
    )


@router.get("/batch/{batch_id}/download/zip")
async def download_batch_zip(batch_id: str):
    """Download the result ZIP for a completed batch job."""
    zip_path = OUTPUT_DIR / batch_id / "result.zip"
    if not zip_path.exists():
        raise HTTPException(status_code=404, detail="ZIP 文件尚未生成或不存在")
    return FileResponse(
        path=str(zip_path),
        filename=f"{batch_id}_排版结果.zip",
        media_type="application/zip",
    )


@router.get("/batch/{batch_id}/download/report")
async def download_batch_report(batch_id: str):
    """Download the Excel report for a completed batch job."""
    report_path = OUTPUT_DIR / batch_id / "report.xlsx"
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="报告文件尚未生成或不存在")
    return FileResponse(
        path=str(report_path),
        filename=f"{batch_id}_处理报告.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
