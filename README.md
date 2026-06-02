# Word2PDF 批量转换工具

批量将 `.doc` / `.docx` 转换为 `.pdf`，支持多种重命名模式，自带本地网页界面。

## 功能特性

- **真实格式转换**：调用 LibreOffice 进行转换，不是简单改后缀
- **三种重命名模式**：保留原名 / 映射表命名 / 顺序名单命名
- **源文件顺序控制**：可显式指定文件对应关系，避免错位
- **本地网页版**：浏览器操作，适合发给不会用命令行的朋友
- **跨平台**：Windows / macOS / Linux 均可使用

## 快速开始

### 1. 安装依赖

- Python 3.10+
- [LibreOffice](https://www.libreoffice.org/)（确保 `soffice` 命令可用）

```bash
# macOS
brew install libreoffice

# Ubuntu/Debian
sudo apt install libreoffice

# Windows：下载安装 LibreOffice，或使用便携版
```

如果 `soffice` 不在 PATH 中，用 `--soffice-bin` 指定路径。

### 2. 基本用法

```bash
# 保持原文件名，仅转换格式
python convert_word_to_pdf.py --input-dir ./word文件 --output-dir ./pdf输出

# 递归处理子目录
python convert_word_to_pdf.py --input-dir ./word文件 --output-dir ./pdf输出 --recursive
```

## 命令行详细用法

### 模式一：保留原名（`--rename-mode keep`）

转换后的 PDF 保持原 Word 文件的主文件名：

```bash
python convert_word_to_pdf.py \
  --input-dir ./word文件 \
  --output-dir ./pdf输出 \
  --rename-mode keep
```

输入 `报告.docx` → 输出 `报告.pdf`

### 模式二：映射表命名（`--rename-mode map`）

准备一个 CSV 文件，两列 `source,target`：

```csv
source,target
报告.docx,2024年度报告
合同.docx,项目合同-终版
```

```bash
python convert_word_to_pdf.py \
  --input-dir ./word文件 \
  --output-dir ./pdf输出 \
  --rename-mode map \
  --mapping-file mapping.csv
```

> `source` 列支持带后缀（`报告.docx`）或不带后缀（`报告`），自动匹配。

### 模式三：顺序名单命名（`--rename-mode list`）

准备一个名字列表，每行一个目标名（不带 `.pdf`）：

```txt
合同-001
合同-002
合同-003
```

```bash
python convert_word_to_pdf.py \
  --input-dir ./word文件 \
  --output-dir ./pdf输出 \
  --rename-mode list \
  --name-list-file name_list.txt
```

程序会按文件名自然排序后依次对应名单。

### 指定源文件顺序（避免错位）

当文件数量多时，强烈建议额外提供 `files_order.txt` 明确指定对应关系：

```txt
001.docx
002.doc
003
```

```bash
python convert_word_to_pdf.py \
  --input-dir ./word文件 \
  --output-dir ./pdf输出 \
  --rename-mode list \
  --name-list-file name_list.txt \
  --files-order-file files_order.txt
```

> `name_list.txt` 和 `files_order.txt` 行数必须一致，且源文件必须存在，否则报错停止。

## 完整参数列表

| 参数 | 说明 |
|------|------|
| `--input-dir` | Word 文件所在目录（必填） |
| `--output-dir` | PDF 输出目录（默认同 input-dir） |
| `--recursive` | 递归扫描子目录 |
| `--rename-mode` | 重命名模式：`keep` / `map` / `list` |
| `--mapping-file` | 映射表文件（配合 `map` 模式） |
| `--name-list-file` | 名字列表文件（配合 `list` 模式） |
| `--files-order-file` | 源文件顺序清单（配合 `list` 模式） |
| `--soffice-bin` | LibreOffice 可执行文件路径 |
| `--overwrite` | 覆盖已存在的同名 PDF |

## 本地网页版

提供一个浏览器界面，适合发给不太会用命令行的朋友。

### 启动

```bash
# Windows：双击 start_web.bat
# 或手动运行：
python word2pdf_web.py
```

浏览器会自动打开 `http://localhost:8000`。

### 功能

- 一次上传多个 Word 文件
- 粘贴多行目标名字
- 可选填"源文件顺序"以确保一一对应
- 点击开始，完成后自动下载 ZIP 包

### 打包为 EXE 交付

```bash
# 1. 将便携版 LibreOffice 解压到 runtime/libreoffice/
# 2. 运行打包脚本
build_windows_exe.bat
# 3. 将 dist_release/ 目录发给对方，双击即可使用
```

## 项目结构

```
.
├── convert_word_to_pdf.py      # 命令行主程序
├── word2pdf_web.py             # 本地网页版
├── build_windows_exe.bat       # Windows EXE 打包脚本
├── start_web.bat               # 网页版启动脚本
├── convert_in_container.sh     # Docker 容器转换脚本
├── mapping_example.csv         # 映射表模板
├── name_list_example.txt       # 名字列表模板
├── files_order_example.txt     # 源文件顺序模板
└── README.md
```

## 常见问题

**Q: 转换失败怎么办？**
检查 LibreOffice 是否正确安装，运行 `soffice --version` 确认。

**Q: 文件顺序不对？**
使用 `--files-order-file` 显式指定源文件顺序，而不是依赖自动排序。

**Q: 如何给不会用命令行的朋友用？**
打包成 EXE + 便携 LibreOffice，对方双击 `start_web.bat` 即可。

## License

MIT
