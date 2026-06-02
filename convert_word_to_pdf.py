#!/usr/bin/env python3
import argparse
import csv
import json
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

WORD_EXTS = {".doc", ".docx"}


def natural_key(text: str):
    parts = re.split(r"(\d+)", text.lower())
    key = []
    for part in parts:
        key.append(int(part) if part.isdigit() else part)
    return key


def find_word_files(input_dir: Path, recursive: bool) -> list[Path]:
    if recursive:
        files = [p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in WORD_EXTS]
    else:
        files = [p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in WORD_EXTS]
    return sorted(files, key=lambda p: natural_key(p.name))


def sanitize_pdf_name(name: str) -> str:
    name = name.strip()
    if not name:
        return ""
    invalid = '<>:"/\\|?*'
    table = str.maketrans({ch: "_" for ch in invalid})
    name = name.translate(table).rstrip(" .")
    if not name:
        return ""
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name


def load_mapping(mapping_file: Path) -> dict[str, str]:
    suffix = mapping_file.suffix.lower()
    data: dict[str, str] = {}

    if suffix == ".json":
        with mapping_file.open("r", encoding="utf-8") as f:
            raw = json.load(f)
        if isinstance(raw, dict):
            iterable = raw.items()
        elif isinstance(raw, list):
            iterable = []
            for row in raw:
                if isinstance(row, dict) and "source" in row and "target" in row:
                    iterable.append((row["source"], row["target"]))
        else:
            raise ValueError("JSON 格式不支持，请使用对象或包含 source/target 的数组。")

        for src, dst in iterable:
            if src is None or dst is None:
                continue
            data[str(src).strip()] = str(dst).strip()
        return data

    if suffix in {".csv", ".tsv", ".txt"}:
        delimiter = "\t" if suffix == ".tsv" else ","
        with mapping_file.open("r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(1024)
            f.seek(0)
            sniffed = csv.Sniffer().sniff(sample, delimiters=",\t") if sample else None
            reader = csv.reader(f, delimiter=(sniffed.delimiter if sniffed else delimiter))
            rows = [r for r in reader if r and any(c.strip() for c in r)]

        if not rows:
            return {}

        header = [c.strip().lower() for c in rows[0]]
        has_header = ("source" in header and "target" in header) or ("src" in header and "dst" in header)

        if has_header:
            src_idx = header.index("source") if "source" in header else header.index("src")
            dst_idx = header.index("target") if "target" in header else header.index("dst")
            body = rows[1:]
        else:
            src_idx, dst_idx = 0, 1
            body = rows

        for row in body:
            if len(row) <= max(src_idx, dst_idx):
                continue
            src, dst = row[src_idx].strip(), row[dst_idx].strip()
            if src and dst:
                data[src] = dst
        return data

    raise ValueError("mapping 文件仅支持 .csv/.tsv/.txt/.json")


def load_name_list(name_list_file: Path) -> list[str]:
    names: list[str] = []
    with name_list_file.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                names.append(line)
    return names


def load_files_order(files_order_file: Path) -> list[str]:
    items: list[str] = []
    with files_order_file.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                items.append(line)
    return items


def reorder_files_by_list(files: list[Path], order_items: list[str]) -> list[Path]:
    by_name = {p.name: p for p in files}
    by_stem = {}
    for p in files:
        by_stem.setdefault(p.stem, []).append(p)

    ordered: list[Path] = []
    used: set[Path] = set()

    for raw in order_items:
        item = raw.strip()
        if not item:
            continue
        match = by_name.get(item)
        if match and match not in used:
            ordered.append(match)
            used.add(match)
            continue

        cands = [p for p in by_stem.get(item, []) if p not in used]
        if len(cands) == 1:
            ordered.append(cands[0])
            used.add(cands[0])
            continue
        if len(cands) > 1:
            raise ValueError(f"files-order 中名称不唯一（同 stem 多个文件）: {item}")
        raise ValueError(f"files-order 未匹配到输入文件: {item}")

    if len(ordered) != len(files):
        missing = [p.name for p in files if p not in used]
        raise ValueError(f"files-order 数量与输入文件不一致，仍有未覆盖文件: {missing[:5]}{'...' if len(missing) > 5 else ''}")

    return ordered


def convert_one(soffice_bin: str, src: Path, temp_out: Path) -> Path:
    # Some Windows/LibreOffice combinations fail on complex Unicode names.
    # Copy to a safe ASCII temp file first to stabilize headless conversion.
    staging = temp_out / "_word2pdf_input"
    staging.mkdir(parents=True, exist_ok=True)
    safe_src = staging / f"input_{uuid.uuid4().hex}{src.suffix.lower()}"
    shutil.copy2(src, safe_src)

    generated = temp_out / f"{safe_src.stem}.pdf"
    if generated.exists():
        generated.unlink()

    profile_dir = (temp_out / f"_lo_profile_{uuid.uuid4().hex}").resolve()
    profile_dir.mkdir(parents=True, exist_ok=True)
    user_install_arg = f"-env:UserInstallation={profile_dir.as_uri()}"

    soffice_path = Path(soffice_bin)
    runners: list[str] = [soffice_bin]
    if soffice_path.suffix.lower() in {".exe", ".com"}:
        swriter = soffice_path.with_name("swriter.exe")
        soffice_exe = soffice_path.with_name("soffice.exe")
        if swriter.exists():
            runners.append(str(swriter))
        if soffice_exe.exists() and str(soffice_exe) not in runners:
            runners.append(str(soffice_exe))

    def build_cmd(runner: str, convert_to: str, input_file: Path) -> list[str]:
        return [
            runner,
            "--headless",
            "--nologo",
            "--nolockcheck",
            "--nodefault",
            "--norestore",
            user_install_arg,
            "--convert-to",
            convert_to,
            "--outdir",
            str(temp_out),
            str(input_file),
        ]

    errors: list[str] = []
    attempt = 0
    for runner in runners:
        attempt += 1
        proc = subprocess.run(build_cmd(runner, "pdf:writer_pdf_Export", safe_src), stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if generated.exists():
            return generated
        detail = proc.stderr.strip() or proc.stdout.strip() or f"返回码 {proc.returncode}，无详细输出"
        errors.append(f"尝试{attempt}失败({Path(runner).name} 直转PDF): {detail}")

    # Fallback path: convert to ODT first, then ODT -> PDF.
    # This can bypass crashes from some DOCX edge cases in direct export.
    odt_stage = temp_out / f"{safe_src.stem}.odt"
    if odt_stage.exists():
        odt_stage.unlink()
    for runner in runners:
        attempt += 1
        proc1 = subprocess.run(build_cmd(runner, "odt", safe_src), stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if not odt_stage.exists():
            detail = proc1.stderr.strip() or proc1.stdout.strip() or f"返回码 {proc1.returncode}，无详细输出"
            errors.append(f"尝试{attempt}失败({Path(runner).name} DOCX->ODT): {detail}")
            continue
        if generated.exists():
            generated.unlink()
        proc2 = subprocess.run(build_cmd(runner, "pdf:writer_pdf_Export", odt_stage), stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if generated.exists():
            return generated
        detail = proc2.stderr.strip() or proc2.stdout.strip() or f"返回码 {proc2.returncode}，无详细输出"
        errors.append(f"尝试{attempt}失败({Path(runner).name} ODT->PDF): {detail}")

    candidates = list(temp_out.glob("*.pdf"))
    if len(candidates) == 1:
        return candidates[0]
    raise RuntimeError(f"转换失败: {src.name}\\n" + "\\n".join(errors))


def ensure_valid_pdf(pdf_path: Path) -> None:
    with pdf_path.open("rb") as f:
        header = f.read(5)
    if header != b"%PDF-":
        raise RuntimeError(f"输出文件不是有效 PDF: {pdf_path.name}")


def resolve_target_name(
    src: Path,
    idx: int,
    mode: str,
    mapping: dict[str, str] | None,
    name_list: list[str] | None,
) -> str:
    if mode == "keep":
        return f"{src.stem}.pdf"

    if mode == "map":
        assert mapping is not None
        key_candidates = [src.name, src.stem]
        target = ""
        for k in key_candidates:
            if k in mapping:
                target = mapping[k]
                break
        if not target:
            raise ValueError(f"映射缺失: {src.name} (需要在 mapping 中配置)")
        safe = sanitize_pdf_name(target)
        if not safe:
            raise ValueError(f"映射目标名无效: {src.name} -> {target}")
        return safe

    if mode == "list":
        assert name_list is not None
        if idx >= len(name_list):
            raise ValueError(f"name-list 数量不足: 第 {idx+1} 个文件 {src.name} 没有对应目标名")
        safe = sanitize_pdf_name(name_list[idx])
        if not safe:
            raise ValueError(f"name-list 目标名无效: 第 {idx+1} 行")
        return safe

    raise ValueError(f"未知模式: {mode}")


def main() -> int:
    parser = argparse.ArgumentParser(description="批量 Word 转 PDF，并支持重命名。")
    parser.add_argument("--input-dir", required=True, help="Word 文件目录")
    parser.add_argument("--output-dir", default="./output_pdf", help="PDF 输出目录")
    parser.add_argument("--recursive", action="store_true", help="递归扫描子目录")
    parser.add_argument(
        "--rename-mode",
        choices=["keep", "map", "list"],
        default="keep",
        help="keep=保持原名; map=按映射表; list=按顺序名单",
    )
    parser.add_argument("--mapping-file", help="映射表文件(.csv/.tsv/.txt/.json)，配合 --rename-mode map")
    parser.add_argument("--name-list-file", help="顺序名单文件(每行一个名称)，配合 --rename-mode list")
    parser.add_argument("--files-order-file", help="文件顺序清单(每行一个源文件名或stem)，用于精确控制 list 模式对应关系")
    parser.add_argument("--soffice-bin", default="soffice", help="LibreOffice 可执行文件路径")
    parser.add_argument("--overwrite", action="store_true", help="允许覆盖同名 PDF")
    args = parser.parse_args()

    input_dir = Path(args.input_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"[ERROR] 输入目录不存在: {input_dir}")
        return 2

    soffice_path = shutil.which(args.soffice_bin) if args.soffice_bin == "soffice" else args.soffice_bin
    if not soffice_path:
        print("[ERROR] 未找到 soffice。请先安装 LibreOffice，或通过 --soffice-bin 指定路径。")
        return 2

    files = find_word_files(input_dir, args.recursive)
    if not files:
        print("[WARN] 未找到 .doc/.docx 文件")
        return 0

    mapping = None
    name_list = None
    if args.rename_mode == "map":
        if not args.mapping_file:
            print("[ERROR] rename-mode=map 时必须提供 --mapping-file")
            return 2
        try:
            mapping = load_mapping(Path(args.mapping_file).expanduser().resolve())
        except Exception as e:
            print(f"[ERROR] 读取 mapping 失败: {e}")
            return 2
    elif args.rename_mode == "list":
        if not args.name_list_file:
            print("[ERROR] rename-mode=list 时必须提供 --name-list-file")
            return 2
        try:
            name_list = load_name_list(Path(args.name_list_file).expanduser().resolve())
        except Exception as e:
            print(f"[ERROR] 读取 name-list 失败: {e}")
            return 2

    if args.files_order_file:
        try:
            order_items = load_files_order(Path(args.files_order_file).expanduser().resolve())
            files = reorder_files_by_list(files, order_items)
        except Exception as e:
            print(f"[ERROR] 读取/应用 files-order 失败: {e}")
            return 2

    output_dir.mkdir(parents=True, exist_ok=True)

    ok = 0
    fail = 0

    print(f"[INFO] 输入目录: {input_dir}")
    print(f"[INFO] 输出目录: {output_dir}")
    print(f"[INFO] 文件数量: {len(files)}")
    print(f"[INFO] 重命名模式: {args.rename_mode}")

    for idx, src in enumerate(files):
        try:
            target_name = resolve_target_name(src, idx, args.rename_mode, mapping, name_list)
            target_path = output_dir / target_name

            if target_path.exists() and not args.overwrite:
                raise FileExistsError(f"目标文件已存在: {target_path.name} (可加 --overwrite)")

            with tempfile.TemporaryDirectory(prefix="word2pdf_") as td:
                temp_out = Path(td)
                generated = convert_one(soffice_path, src, temp_out)
                ensure_valid_pdf(generated)
                shutil.move(str(generated), str(target_path))

            ok += 1
            print(f"[OK] {src.name} -> {target_path.name}")

        except Exception as e:
            fail += 1
            print(f"[FAIL] {src.name}: {e}")

    print("-" * 48)
    print(f"[DONE] 成功: {ok}, 失败: {fail}, 总计: {len(files)}")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
