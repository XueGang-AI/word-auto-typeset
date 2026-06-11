#!/usr/bin/env python3
import argparse
import cgi
import html
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import webbrowser
import zipfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from convert_word_to_pdf import (
    convert_one,
    ensure_valid_pdf,
    natural_key,
    reorder_files_by_list,
    sanitize_pdf_name,
)
from format_with_template import format_document

PAGE_CSS = """
:root{
  --bg:#f5f1ea;
  --card:#fffaf3;
  --ink:#1f2937;
  --muted:#6b7280;
  --line:#e7ddd0;
  --accent:#8b5e34;
  --accent2:#2f6f73;
  --shadow:0 18px 50px rgba(31,41,55,.12);
}
*{box-sizing:border-box}
body{
  margin:0;
  font-family:"Aptos","Noto Sans SC","Microsoft YaHei",sans-serif;
  color:var(--ink);
  background:
    radial-gradient(circle at top left, rgba(139,94,52,.12), transparent 28%),
    radial-gradient(circle at right top, rgba(47,111,115,.12), transparent 30%),
    linear-gradient(180deg, #fff, var(--bg));
  min-height:100vh;
}
.wrap{max-width:1100px;margin:0 auto;padding:40px 20px 56px}
.hero{
  display:grid;gap:18px;grid-template-columns:1.3fr .7fr;align-items:stretch;
  margin-bottom:22px;
}
.brand,.panel,.note,.result{
  background:rgba(255,250,243,.88);
  border:1px solid rgba(231,221,208,.85);
  box-shadow:var(--shadow);
  border-radius:24px;
  backdrop-filter: blur(8px);
}
.brand{padding:28px}
.kicker{color:var(--accent2);font-weight:700;letter-spacing:.08em;text-transform:uppercase;font-size:12px}
h1{margin:10px 0 12px;font-size:42px;line-height:1.05}
.sub{color:var(--muted);font-size:15px;line-height:1.75;max-width:60ch}
.pillrow{display:flex;flex-wrap:wrap;gap:10px;margin-top:18px}
.pill{
  padding:8px 12px;border-radius:999px;background:#fff;border:1px solid var(--line);
  color:#444;font-size:13px
}
.status{
  padding:28px;display:flex;flex-direction:column;justify-content:space-between;gap:16px
}
.statusbox{
  border-radius:20px;padding:18px;background:linear-gradient(135deg, rgba(139,94,52,.08), rgba(47,111,115,.08));
  border:1px solid rgba(139,94,52,.15)
}
.statusbox strong{display:block;margin-bottom:8px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.panel{padding:22px}
label{display:block;font-weight:700;margin:0 0 10px}
input[type=file], textarea, input[type=text]{
  width:100%;border:1px solid var(--line);border-radius:16px;background:#fff;
  padding:14px 15px;font:inherit;color:var(--ink)
}
textarea{min-height:220px;resize:vertical;line-height:1.6}
.small{font-size:13px;color:var(--muted);margin-top:8px;line-height:1.6}
.row{display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin-top:14px}
.check{display:flex;align-items:center;gap:8px;font-size:14px;color:#374151}
button{
  border:none;border-radius:14px;padding:14px 18px;font-weight:800;font-size:15px;color:white;
  background:linear-gradient(135deg, var(--accent), #b07a49);cursor:pointer;box-shadow:0 12px 25px rgba(139,94,52,.22)
}
button:hover{filter:brightness(1.03)}
.note{padding:18px 20px;margin-top:18px;color:#4b5563;line-height:1.75}
.result{padding:18px 20px;margin-top:18px}
.tag{display:inline-block;padding:4px 10px;border-radius:999px;background:rgba(47,111,115,.1);color:var(--accent2);font-size:12px;font-weight:700}
@media (max-width:900px){
  .hero,.grid{grid-template-columns:1fr}
  h1{font-size:34px}
}
"""


def read_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def extract_uploaded_files(form: cgi.FieldStorage, field_name: str = "files") -> list[tuple[str, bytes]]:
    uploaded_files: list[tuple[str, bytes]] = []
    items = form.list or []
    for item in items:
        if getattr(item, "name", None) != field_name:
            continue
        filename = getattr(item, "filename", "")
        if not filename:
            continue
        uploaded_files.append((filename, item.file.read()))
    return uploaded_files


def sanitize_zip_name(name: str) -> str:
    name = Path(name.strip() or "word2pdf_result.zip").name
    base = name[:-4] if name.lower().endswith(".zip") else name
    invalid = '<>:"/\\|?*'
    table = str.maketrans({ch: "_" for ch in invalid})
    base = base.translate(table).strip().rstrip(" .")
    if not base:
        base = "word2pdf_result"
    return f"{base}.zip"


def safe_upload_name(name: str) -> str:
    if isinstance(name, bytes):
        name = name.decode("utf-8", errors="replace")
    name = name.replace("\\", "/").split("/")[-1].strip()
    name = name or "upload"
    return name


def app_base_dir() -> Path:
    # PyInstaller onefile: sys.executable points to the extracted launcher.
    # Dev mode: fall back to script directory.
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_runtime_soffice() -> str | None:
    candidates = []
    env_bin = os.environ.get("WORD2PDF_SOFFICE", "").strip()
    if env_bin:
        candidates.append(env_bin)

    base = app_base_dir()
    candidates += [
        str(base / "runtime" / "libreoffice" / "program" / "soffice"),
        str(base / "runtime" / "libreoffice" / "program" / "soffice.exe"),
        str(base / "runtime" / "LibreOffice" / "program" / "soffice"),
        str(base / "runtime" / "LibreOffice" / "program" / "soffice.exe"),
    ]
    for name in ("soffice.com", "soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            candidates.append(found)

    for cand in candidates:
        if cand and Path(cand).exists():
            return cand
    return None


def docker_available() -> bool:
    return shutil.which("docker") is not None


def backend_description() -> str:
    soffice = find_runtime_soffice()
    if soffice:
        return f"本地 LibreOffice: {soffice}"
    if getattr(sys, "frozen", False):
        return "未检测到内置 LibreOffice（请检查 runtime/libreoffice）"
    if docker_available():
        image = os.environ.get("WORD2PDF_DOCKER_IMAGE", "linuxserver/libreoffice:latest")
        return f"Docker LibreOffice: {image}"
    return "未检测到可用转换引擎"


def convert_batch_with_soffice(src_files: list[Path], raw_out: Path, soffice_bin: str) -> None:
    for src in src_files:
        convert_one(soffice_bin, src, raw_out)


def convert_batch_with_docker(src_dir: Path, raw_out: Path, image: str) -> None:
    cmd = [
        "docker",
        "run",
        "--rm",
        "--entrypoint",
        "/bin/bash",
        "-v",
        f"{src_dir}:/in",
        "-v",
        f"{raw_out}:/out",
        image,
        "-lc",
        'shopt -s nullglob; for f in /in/*; do case "$f" in *.doc|*.DOC|*.docx|*.DOCX) soffice --headless --convert-to pdf --outdir /out "$f" ;; esac; done',
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "Docker 转换失败")


def format_job(template_data: bytes, template_name: str, content_data: bytes, content_name: str) -> bytes:
    """格式对齐任务：模板+内容 → 格式化后的 Word 文档。"""
    with tempfile.TemporaryDirectory(prefix="word2pdf_format_") as td:
        root = Path(td)
        template_path = root / safe_upload_name(template_name)
        content_path = root / safe_upload_name(content_name)
        template_path.write_bytes(template_data)
        content_path.write_bytes(content_data)

        output_path = root / f"{content_path.stem}_格式对齐.docx"

        stats = format_document(template_path, content_path, output_path)

        if not output_path.exists():
            raise RuntimeError("格式对齐失败，未生成输出文件")

        return output_path.read_bytes(), stats


def build_zip(final_dir: Path, zip_path: Path, extra_files: list[tuple[str, Path]] | None = None) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for pdf in sorted(final_dir.glob("*.pdf"), key=lambda p: natural_key(p.name)):
            zf.write(pdf, arcname=pdf.name)
        if extra_files:
            for arcname, src in extra_files:
                if src.exists():
                    zf.write(src, arcname=arcname)


def convert_job(uploaded_files: list[tuple[str, bytes]], target_names: list[str], source_order: list[str] | None, overwrite: bool) -> bytes:
    if len(uploaded_files) != len(target_names):
        raise ValueError(f"文件数与名字数不一致：{len(uploaded_files)} vs {len(target_names)}")

    soffice_bin = find_runtime_soffice()
    image = os.environ.get("WORD2PDF_DOCKER_IMAGE", "linuxserver/libreoffice:latest")

    with tempfile.TemporaryDirectory(prefix="word2pdf_web_") as td:
        root = Path(td)
        src_dir = root / "input"
        raw_dir = root / "raw"
        final_dir = root / "final"
        src_dir.mkdir()
        raw_dir.mkdir()
        final_dir.mkdir()

        saved_paths: list[Path] = []
        seen_names: set[str] = set()
        for original_name, data in uploaded_files:
            safe_name = safe_upload_name(original_name)
            if safe_name in seen_names:
                raise ValueError(f"上传文件名重复：{safe_name}")
            seen_names.add(safe_name)
            p = src_dir / safe_name
            p.write_bytes(data)
            saved_paths.append(p)

        if source_order:
            saved_paths = reorder_files_by_list(saved_paths, source_order)
        else:
            saved_paths = sorted(saved_paths, key=lambda p: natural_key(p.name))

        if len(saved_paths) != len(target_names):
            raise ValueError("源文件数量和目标名字数量不一致")

        failures: list[tuple[str, str, str]] = []
        seen_targets: set[str] = set()
        success_count = 0

        if soffice_bin:
            for idx, src in enumerate(saved_paths):
                raw_target = target_names[idx]
                target_name = sanitize_pdf_name(raw_target)
                if not target_name:
                    failures.append((src.name, raw_target, f"第 {idx + 1} 个目标名无效"))
                    continue
                if target_name in seen_targets and not overwrite:
                    failures.append((src.name, target_name, "目标名字重复（未勾选覆盖）"))
                    continue
                seen_targets.add(target_name)
                try:
                    with tempfile.TemporaryDirectory(prefix="word2pdf_one_") as td_one:
                        one_out = Path(td_one)
                        raw_pdf = convert_one(soffice_bin, src, one_out)
                        ensure_valid_pdf(raw_pdf)
                        shutil.copy2(raw_pdf, final_dir / target_name)
                        success_count += 1
                except Exception as exc:
                    failures.append((src.name, target_name, str(exc)))
        elif getattr(sys, "frozen", False):
            raise RuntimeError("未找到内置 LibreOffice。请确认与 EXE 同级目录存在 runtime/libreoffice/program/soffice.exe")
        elif docker_available():
            convert_batch_with_docker(src_dir, raw_dir, image)
            raw_map = {p.stem.strip(): p for p in raw_dir.glob("*.pdf")}
            if not raw_map:
                raise RuntimeError("没有生成任何 PDF")
            for idx, src in enumerate(saved_paths):
                raw_target = target_names[idx]
                target_name = sanitize_pdf_name(raw_target)
                if not target_name:
                    failures.append((src.name, raw_target, f"第 {idx + 1} 个目标名无效"))
                    continue
                if target_name in seen_targets and not overwrite:
                    failures.append((src.name, target_name, "目标名字重复（未勾选覆盖）"))
                    continue
                seen_targets.add(target_name)
                raw_pdf = raw_map.get(src.stem.strip())
                if raw_pdf is None:
                    matches = [p for p in raw_dir.glob("*.pdf") if p.stem.strip() == src.stem.strip()]
                    if matches:
                        raw_pdf = matches[0]
                if raw_pdf is None:
                    failures.append((src.name, target_name, "未找到对应 PDF 输出"))
                    continue
                try:
                    ensure_valid_pdf(raw_pdf)
                    shutil.copy2(raw_pdf, final_dir / target_name)
                    success_count += 1
                except Exception as exc:
                    failures.append((src.name, target_name, str(exc)))
        else:
            raise RuntimeError("未找到可用的转换引擎：请放入可用的 LibreOffice portable，或安装 Docker。")

        fail_path = root / "失败清单.txt"
        lines = [
            "Word 转 PDF 批量失败清单",
            f"总文件数: {len(saved_paths)}",
            f"成功数: {success_count}",
            f"失败数: {len(failures)}",
            "",
        ]
        if failures:
            for i, (src_name, target_name, reason) in enumerate(failures, start=1):
                lines.append(f"{i}. 源文件: {src_name}")
                lines.append(f"   目标名: {target_name}")
                lines.append(f"   原因: {reason}")
        else:
            lines.append("无失败")
        fail_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        extras: list[tuple[str, Path]] = [("失败清单.txt", fail_path)]

        if success_count == 0 and not failures:
            raise RuntimeError("没有生成任何 PDF")

        zip_path = root / "word2pdf_result.zip"
        build_zip(final_dir, zip_path, extra_files=extras)
        return zip_path.read_bytes()


def render_page(message: str = "", error: str = "", backend: str = "", format_stats: dict | None = None) -> str:
    format_msg = ""
    if format_stats:
        lines = [
            f"格式对齐完成！原段落 {format_stats['total_original']} → 删除空行 {format_stats['empty_removed']}",
            f"主标题 {format_stats.get('main_title', 0)} | 章节标题 {format_stats.get('article_title', 0)} | 层级标题 {format_stats.get('section_header', 0)}",
            f"作者 {format_stats.get('author_name', 0)} | 正文 {format_stats.get('body_text', 0)} | 标签 {format_stats.get('tag_label', 0)} | 单位 {format_stats.get('affiliation', 0)} | 图片 {format_stats.get('image', 0)}",
        ]
        format_msg = "<div class='result' style='border-color:#b5e8c3;background:#f7fff9'><span class='tag' style='background:rgba(28,185,74,.08);color:#1b9928'>格式对齐完成</span><div style='margin-top:10px;white-space:pre-wrap;font-size:14px;line-height:1.8'>" + "<br>".join(html.escape(l) for l in lines) + "</div></div>"

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Word 批量转 PDF</title>
  <style>{PAGE_CSS}</style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <div class="brand">
        <div class="kicker">Local Batch Converter</div>
        <h1>Word 批量转 PDF</h1>
        <div class="sub">两步走：① 上传模板 + 内容文档 → 格式对齐；② 上传 Word + 粘贴命名 → 批量转 PDF 下载 ZIP。</div>
        <div class="pillrow">
          <span class="pill">格式对齐</span>
          <span class="pill">支持 .doc / .docx</span>
          <span class="pill">批量重命名</span>
          <span class="pill">结果打包下载</span>
        </div>
      </div>
      <div class="status">
        <div class="statusbox">
          <strong>当前转换引擎</strong>
          <div>{html.escape(backend or backend_description())}</div>
        </div>
        <div class="note" style="margin:0">
          <strong>流程：</strong>先做格式对齐（得到格式化 Word）→ 再做批量转 PDF 并命名。
        </div>
      </div>
    </div>

    {format_msg}
    {"<div class='result'><span class='tag'>提示</span><div style='margin-top:10px'>" + html.escape(message) + "</div></div>" if message else ""}
    {"<div class='result' style='border-color:#e8b5b5;background:#fff7f7'><span class='tag' style='background:rgba(185,28,28,.08);color:#991b1b'>错误</span><div style='margin-top:10px;white-space:pre-wrap'>" + html.escape(error) + "</div></div>" if error else ""}

    <!-- 格式对齐 -->
    <div class="panel" style="margin-bottom:18px">
      <div class="kicker" style="margin-bottom:4px">步骤一</div>
      <h2 style="margin:0 0 16px;font-size:22px">格式对齐（模板 → 内容）</h2>
      <form action="/format" method="post" enctype="multipart/form-data">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
          <div>
            <label>上传模板文档（定义目标格式）</label>
            <input type="file" name="template" accept=".docx,.doc" required />
            <div class="small">模板决定字体、字号、缩进、行距等格式。可随时替换模板。</div>
          </div>
          <div>
            <label>上传内容文档（待格式化的 Word）</label>
            <input type="file" name="content" accept=".docx,.doc" required />
            <div class="small">内容保持原样，格式对齐到模板。</div>
          </div>
        </div>
        <div class="row" style="margin-top:16px">
          <button type="submit">开始格式对齐</button>
          <span class="small" style="margin:0">输出：格式对齐后的 .docx 文件</span>
        </div>
      </form>
    </div>

    <!-- 批量转 PDF -->
    <div class="panel">
      <div class="kicker" style="margin-bottom:4px">步骤二</div>
      <h2 style="margin:0 0 16px;font-size:22px">批量转 PDF + 命名</h2>
      <form class="grid" action="/convert" method="post" enctype="multipart/form-data">
        <div class="panel" style="box-shadow:none;border:none;padding:0">
          <label>上传 Word 文件（可多选）</label>
          <input type="file" name="files" multiple accept=".doc,.docx,.DOC,.DOCX" />
          <div class="small">选中全部 Word 文件后直接提交。文件会按"源文件顺序"或自然排序依次对应名字列表。</div>

          <div style="height:16px"></div>
          <label>目标名字列表</label>
          <textarea name="names" placeholder="h1wh202605001&#10;h1wh202605002&#10;h1wh202605003" required></textarea>
          <div class="small">一行一个名字，不要带 <code>.pdf</code> 也可以，程序会自动补上。</div>
        </div>

        <div class="panel" style="box-shadow:none;border:none;padding:0">
          <label>源文件顺序（可选，但强烈建议 200 份时填写）</label>
          <textarea name="order" placeholder="1.docx&#10;2.docx&#10;3.docx"></textarea>
          <div class="small">一行一个源文件名或文件主名，用来精确控制与名字列表的对应关系。</div>

          <div style="height:16px"></div>
          <label>下载文件名</label>
          <input type="text" name="zip_name" value="word2pdf_result.zip" />

          <div class="row">
            <label class="check"><input type="checkbox" name="overwrite" checked />允许覆盖同名目标</label>
          </div>

          <div class="row" style="margin-top:18px">
            <button type="submit">开始批量转换</button>
          </div>
        </div>
      </form>
    </div>

    <div class="note">
      使用方式很简单：先格式对齐、再批量转 PDF。若文件顺序不稳定，就把源文件顺序也贴上。
    </div>
  </div>
</body>
</html>"""


class Word2PDFHandler(BaseHTTPRequestHandler):
    server_version = "Word2PDFWeb/1.0"

    def log_message(self, fmt: str, *args) -> None:  # noqa: A003
        sys.stdout.write("%s - - [%s] %s\n" % (self.client_address[0], self.log_date_time_string(), fmt % args))

    def _send_html(self, body: str, status: int = 200) -> None:
        data = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:  # noqa: N802
        if self.path not in {"/", "/index.html"}:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        self._send_html(render_page(backend=self.server.backend_desc))

    def do_POST(self) -> None:  # noqa: N802
        if self.path == "/format":
            self._handle_format()
        elif self.path == "/convert":
            self._handle_convert()
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def _handle_format(self) -> None:
        """处理格式对齐请求：模板 + 内容 → 格式化后的 Word。"""
        ctype, _ = cgi.parse_header(self.headers.get("content-type", ""))
        if ctype != "multipart/form-data":
            self._send_html(render_page(error="请使用网页表单上传文件。", backend=self.server.backend_desc), status=400)
            return

        env = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.headers.get("content-type", ""),
        }
        if "content-length" in self.headers:
            env["CONTENT_LENGTH"] = self.headers["content-length"]

        try:
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=env, keep_blank_values=True)

            template_files = extract_uploaded_files(form, "template")
            if not template_files:
                raise ValueError("请上传模板文档。")
            template_name, template_data = template_files[0]

            content_files = extract_uploaded_files(form, "content")
            if not content_files:
                raise ValueError("请上传内容文档。")
            content_name, content_data = content_files[0]

            payload, stats = format_job(template_data, template_name, content_data, content_name)

            safe_content_name = safe_upload_name(content_name)
            output_name = Path(safe_content_name).stem + "_格式对齐.docx"
            # HTTP header 只支持 ASCII，中文文件名需 URL 编码
            from urllib.parse import quote
            output_name_encoded = quote(output_name)

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f"attachment; filename=\"{output_name_encoded}\"; filename*=UTF-8''{output_name_encoded}")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
        except Exception as exc:
            import traceback
            traceback.print_exc()
            self._send_html(
                render_page(error=f"{exc}", backend=self.server.backend_desc),
                status=400,
            )

    def _handle_convert(self) -> None:
        """处理批量转换请求：Word → PDF ZIP。"""
        ctype, _ = cgi.parse_header(self.headers.get("content-type", ""))
        if ctype != "multipart/form-data":
            self._send_html(render_page(error="请使用网页表单上传文件。", backend=self.server.backend_desc), status=400)
            return

        env = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.headers.get("content-type", ""),
        }
        if "content-length" in self.headers:
            env["CONTENT_LENGTH"] = self.headers["content-length"]

        try:
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=env, keep_blank_values=True)

            names_text = form.getfirst("names", "")
            target_names = read_lines(names_text)
            if not target_names:
                raise ValueError("目标名字列表不能为空。")

            order_text = form.getfirst("order", "")
            source_order = read_lines(order_text) or None
            overwrite = bool(form.getfirst("overwrite"))
            zip_name = sanitize_zip_name(form.getfirst("zip_name", "word2pdf_result.zip"))

            uploaded_files = extract_uploaded_files(form, "files")
            if not uploaded_files:
                raise ValueError("没有读取到任何上传文件。")

            if len(uploaded_files) != len(target_names):
                raise ValueError(f"上传文件数量 {len(uploaded_files)} 与名字数量 {len(target_names)} 不一致。")

            payload = convert_job(uploaded_files, target_names, source_order, overwrite=overwrite)

            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", f'attachment; filename="{zip_name}"')
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
        except Exception as exc:
            self._send_html(
                render_page(
                    error=f"{exc}",
                    backend=self.server.backend_desc,
                ),
                status=400,
            )


def main() -> int:
    parser = argparse.ArgumentParser(description="本地网页批量 Word 转 PDF 工具。")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址")
    parser.add_argument("--port", type=int, default=8765, help="监听端口")
    parser.add_argument("--no-open", action="store_true", help="不自动打开浏览器")
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), Word2PDFHandler)
    server.backend_desc = backend_description()

    url = f"http://{args.host}:{args.port}/"
    print(f"[INFO] 本地网页已启动: {url}")
    print(f"[INFO] 转换引擎: {server.backend_desc}")

    if not args.no_open:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] 已退出")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
