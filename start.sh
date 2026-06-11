#!/bin/bash
# Word 自动套版系统 - 启动脚本
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "========================================="
echo "  Word 自动套版系统 v2.0"
echo "========================================="

# Check Python
PYTHON=""
for py in python3.11 python3.10 python3; do
    if command -v $py &>/dev/null; then
        PYTHON=$py
        break
    fi
done
if [ -z "$PYTHON" ]; then
    echo "[ERROR] 未找到 Python 3.10+"
    exit 1
fi
echo "[INFO] Python: $($PYTHON --version)"

# Check deps
$PYTHON -c "import fastapi" 2>/dev/null || {
    echo "[INFO] 安装 Python 依赖..."
    $PYTHON -m pip install -r requirements.txt
}

# Check frontend build
if [ ! -d "frontend/dist" ]; then
    echo "[INFO] 构建前端..."
    cd frontend
    npm install --cache /tmp/npm-cache 2>/dev/null
    npm run build
    cd ..
fi

echo "[INFO] 启动服务..."
echo "[INFO] 访问地址: http://localhost:8765"
echo "[INFO] API 文档: http://localhost:8765/api/docs"
echo ""

$PYTHON -m uvicorn app.main:app --host 0.0.0.0 --port 8765 "$@"
