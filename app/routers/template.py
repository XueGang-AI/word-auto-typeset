"""
Template management API endpoints.
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile

from app.config import TEMPLATES_DIR
from app.schemas.models import APIResponse, TemplateConfig, TemplateListItem
from app.services.template_service import (
    aggregate_templates,
    analyze_template,
    delete_template,
    get_template,
    list_templates,
    save_template,
)

router = APIRouter(prefix="/api/templates", tags=["templates"])


@router.get("/", response_model=APIResponse)
async def list_all_templates():
    """List all saved templates."""
    items = list_templates()
    return APIResponse(
        success=True,
        message=f"共 {len(items)} 个模板",
        data=[item.model_dump() for item in items],
    )


@router.get("/{template_id}", response_model=APIResponse)
async def get_template_detail(template_id: str):
    """Get a template's full configuration."""
    config = get_template(template_id)
    if not config:
        raise HTTPException(status_code=404, detail="模板不存在")
    return APIResponse(
        success=True,
        message="获取模板成功",
        data=config.model_dump(),
    )


@router.post("/upload", response_model=APIResponse)
async def upload_template(file: UploadFile = File(...)):
    """
    Upload a template Word document for analysis.

    The server will:
    1. Analyze the document's styles
    2. Extract page settings
    3. Generate template_config.json
    4. Save the template for later use
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="文件名不能为空")

    # Validate file type
    ext = Path(file.filename).suffix.lower()
    if ext not in (".docx", ".doc"):
        # Accept .doc but warn that only .docx is fully supported
        if ext != ".docx":
            pass  # We'll try anyway

    try:
        content = await file.read()
        config = save_template(content, file.filename)
        return APIResponse(
            success=True,
            message=f"模板 '{file.filename}' 上传并分析完成",
            data=config.model_dump(),
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"模板分析失败: {e}")


@router.post("/batch-upload", response_model=APIResponse)
async def batch_upload_templates(files: list[UploadFile] = File(...)):
    """
    Upload multiple template Word documents at once.

    Each file is analyzed independently with `analyze_template()` and saved
    as its own template. The response lists the new template IDs and a
    convenience aggregated config (majority vote across all uploaded
    templates) that can be used directly for typesetting.
    """
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个模板文件")

    saved: list[dict] = []
    configs: list[TemplateConfig] = []
    errors: list[dict] = []

    for f in files:
        if not f.filename:
            errors.append({"filename": "", "error": "文件名为空"})
            continue
        try:
            content = await f.read()
            config = save_template(content, f.filename)
            configs.append(config)
            saved.append({
                "filename": f.filename,
                "template_id": config.template_id,
                "template_name": config.template_name,
            })
        except Exception as e:
            errors.append({"filename": f.filename, "error": str(e)})

    aggregated = None
    if configs:
        try:
            agg = aggregate_templates(configs)
            # Persist the aggregated config so it's selectable like any
            # regular template.
            tdir = TEMPLATES_DIR / agg.template_id
            tdir.mkdir(parents=True, exist_ok=True)
            (tdir / "template_config.json").write_text(
                agg.model_dump_json(indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            aggregated = agg.model_dump()
        except Exception as e:
            errors.append({"filename": "(aggregate)", "error": str(e)})

    if not saved and errors:
        raise HTTPException(
            status_code=400,
            detail=f"所有模板上传失败: {errors[0]['error']}",
        )

    return APIResponse(
        success=True,
        message=f"批量上传完成：{len(saved)} 成功，{len(errors)} 失败",
        data={
            "uploaded": saved,
            "failed": errors,
            "aggregated": aggregated,
        },
    )


@router.put("/{template_id}/replace", response_model=APIResponse)
async def replace_template(template_id: str, file: UploadFile = File(...)):
    """
    Replace an existing template with a new Word file.

    The old template config is replaced with the new analysis.
    """
    existing = get_template(template_id)
    if not existing:
        raise HTTPException(status_code=404, detail="模板不存在")

    # Delete old files
    delete_template(template_id)

    try:
        content = await file.read()
        config = save_template(content, file.filename or existing.template_name)
        # Preserve the original template_id
        config.template_id = template_id
        # Re-save with original ID
        from app.services.template_service import _config_path, _docx_path
        _config_path(template_id).write_text(
            config.model_dump_json(indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        return APIResponse(
            success=True,
            message=f"模板 '{template_id}' 已替换",
            data=config.model_dump(),
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"模板替换失败: {e}")


@router.delete("/{template_id}", response_model=APIResponse)
async def remove_template(template_id: str):
    """Delete a template and all associated files."""
    if delete_template(template_id):
        return APIResponse(success=True, message=f"模板 '{template_id}' 已删除")
    raise HTTPException(status_code=404, detail="模板不存在")


@router.get("/{template_id}/config", response_model=APIResponse)
async def get_template_config(template_id: str):
    """Get the template_config.json for a template."""
    config = get_template(template_id)
    if not config:
        raise HTTPException(status_code=404, detail="模板不存在")
    return APIResponse(
        success=True,
        message="获取配置成功",
        data=config.model_dump(),
    )


@router.put("/{template_id}/config", response_model=APIResponse)
async def update_template_config(template_id: str, config: TemplateConfig):
    """
    Update template configuration (block rules, style overrides).

    This allows users to customize which content types are rendered
    and adjust style settings without re-uploading the template.
    """
    existing = get_template(template_id)
    if not existing:
        raise HTTPException(status_code=404, detail="模板不存在")

    from app.services.template_service import _config_path
    _config_path(template_id).write_text(
        config.model_dump_json(indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return APIResponse(
        success=True,
        message="模板配置已更新",
        data=config.model_dump(),
    )
