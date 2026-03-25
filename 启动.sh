#!/bin/bash
# 发票识别工具 - Mac/Linux 启动脚本

echo "============================================"
echo "        发票识别工具 - 启动中..."
echo "============================================"
echo ""

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# 检查 Python 是否安装
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到 Python3，请先安装 Python 3.9 以上版本。"
    echo "Mac: brew install python3"
    echo "Linux: sudo apt install python3 python3-pip"
    exit 1
fi

echo "[1/2] 检查并安装依赖..."
pip3 install -r "$SCRIPT_DIR/requirements.txt" --quiet

echo ""
echo "[2/2] 启动应用..."
echo "浏览器会自动打开，如果没有请手动访问 http://localhost:8501"
echo "按 Ctrl+C 停止应用。"
echo ""

streamlit run "$SCRIPT_DIR/invoice_ui.py" --server.headless=false
