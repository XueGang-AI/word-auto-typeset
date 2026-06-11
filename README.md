# Word 自动套版系统

## 项目概述

基于 FastAPI + Vue3 的 Word 文档自动排版系统，支持：

1. **模板管理** — 上传模板 Word → 自动分析样式（14种样式配置）→ 保存 `template_config.json`
2. **AI 内容识别** — 纯 LLM 路线，调用 OpenAI 兼容 API（DeepSeek/OpenAI/豆包），17种段落类型分类（标题、作者、正文、来源等）。无规则回退，未配置时会拒绝运行。
3. **自动套版渲染** — 模板样式 + AI 识别结果 → 生成排版后的 Word（保留图片、表格）
4. **批量处理** — 多文件并发处理，输出 ZIP + report.xlsx
5. **Word 转 PDF** — 批量 Word → PDF 转换（LibreOffice）

## 项目结构

```
word-auto-typeset/
├── app/                              # FastAPI 后端
│   ├── main.py                       # 应用入口，路由注册，静态文件挂载
│   ├── config.py                     # 配置管理（AI、路径、并发、LibreOffice）
│   ├── routers/
│   │   ├── template.py               # 模板 CRUD API
│   │   ├── typeset.py                # 单文件/批量套版 API
│   │   └── convert.py                # Word → PDF 转换 API
│   ├── services/
│   │   ├── template_service.py       # 模板分析引擎（评分制样式提取）
│   │   ├── content_parser.py         # 内容解析（段落、图片、表格）
│   │   ├── ai_recognizer.py          # AI 结构识别（纯 LLM 路线）
│   │   ├── renderer.py               # Word 渲染引擎（样式套用 + 图片嵌入）
│   │   └── batch_service.py          # 批量任务调度
│   ├── schemas/
│   │   └── models.py                 # Pydantic 数据模型
│   └── utils/
│       ├── word_utils.py             # Word 工具函数
│       └── file_utils.py             # 文件处理工具
├── frontend/                         # Vue3 + Element Plus + Vite
│   └── src/views/
│       ├── TemplateManage.vue        # 模板管理
│       ├── SingleTypeset.vue         # 单文件套版
│       ├── BatchTypeset.vue          # 批量套版
│       └── WordToPDF.vue             # Word 转 PDF
├── legacy/                           # 原始脚本（保留兼容）
├── requirements.txt
└── start.sh                          # 一键启动脚本
```

## 快速开始

### 环境要求

- Python 3.10+（macOS 需 `/opt/homebrew/bin/python3.11`）
- Node.js 18+（前端开发）
- LibreOffice（Word → PDF 功能）

### 一键启动

```bash
./start.sh
```

### 手动启动

```bash
# 安装依赖
pip install -r requirements.txt
cd frontend && npm install && npm run build && cd ..

# 启动服务
python3.11 -m uvicorn app.main:app --host 0.0.0.0 --port 8765
```

### 开发模式

```bash
# 终端 1 - 后端（自动重载）
python3.11 -m uvicorn app.main:app --host 0.0.0.0 --port 8765 --reload

# 终端 2 - 前端（热更新）
cd frontend && npm run dev
```

打开浏览器：
- 前端：`http://localhost:5173`
- API 文档：`http://localhost:8765/api/docs`

## API 端点

### 模板管理 `/api/templates`
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | 列出所有模板 |
| POST | `/upload` | 上传并分析模板 |
| PUT | `/{id}/replace` | 替换模板文件 |
| DELETE | `/{id}` | 删除模板 |
| GET | `/{id}/config` | 查看模板配置 |
| PUT | `/{id}/config` | 更新模板配置 |

### 套版 `/api/typeset`
| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/single/download` | 单文件套版并下载 |
| POST | `/analyze` | 仅分析结构（预览） |
| POST | `/batch` | 创建批量任务 |
| GET | `/batch/{id}/progress` | 查询批量进度 |
| GET | `/batch/{id}/download/zip` | 下载结果 ZIP |
| GET | `/batch/{id}/download/report` | 下载处理报告 |

### 转换 `/api/convert`
| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/word-to-pdf` | Word 批量转 PDF |

## AI 结构识别

### 纯 LLM 路线（必需配置）

所有段落分类由云端 LLM 完成，未配置时 `recognize_structure` 直接抛错，无规则回退。

```bash
export AI_ENABLED=true
export AI_API_KEY=sk-xxx
export AI_BASE_URL=https://api.deepseek.com   # 或 https://api.openai.com/v1 等
export AI_MODEL=deepseek-chat                  # 或 gpt-4o-mini 等
```

系统 Prompt 内嵌 10 条 few-shot 示例（文本 + 字号/粗体/颜色/居中/字体/缩进 → 类型），并使用 `response_format: json_object` 强制 JSON 输出。Post-process 仅保证 `main_title` 不变量。

## 核心流水线

```
模板 .docx → 分析样式 → TemplateConfig（14个样式配置 + 页面设置）
内容 .docx → 解析段落/图片/表格 → AI识别结构 → ContentDocument
                                            └→ 渲染引擎 → 输出 .docx
```

- 模板分析采用**评分制**：每种格式模式对每种段落类型打分，取最高分匹配
- 图片通过 `w:drawing` 元素提取原始字节和尺寸，渲染时用 `add_picture()` 嵌入
- 正文保留原始 `run` 级别的粗体（用于书名、关键词强调）

## 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `AI_ENABLED` | `false` | 启用 AI 识别 |
| `AI_API_KEY` | — | API Key |
| `AI_BASE_URL` | `https://api.openai.com/v1` | API 地址 |
| `AI_MODEL` | `gpt-4o-mini` | 模型名称 |
| `WORD2PDF_SOFFICE` | — | LibreOffice 路径 |
| `MAX_CONCURRENT_TASKS` | `4` | 批量并发数 |
| `BATCH_OUTPUT_MAX_AGE_HOURS` | `24` | 启动时清理超过此时长的批量输出目录 |
| `HOST` / `PORT` | `0.0.0.0` / `8765` | 服务地址 |

## 兼容旧版

原始 CLI 工具保留在 `legacy/` 目录：
```bash
python legacy/word2pdf_web.py
python legacy/convert_word_to_pdf.py --input-dir ./docs --output-dir ./pdf
python legacy/format_with_template.py --template t.docx --content c.docx
```

## License

MIT
