@echo off
chcp 65001 >nul 2>&1
title 发票识别工具

echo ============================================
echo         发票识别工具 - 启动中...
echo ============================================
echo.

:: 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9 以上版本。
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo [1/2] 检查并安装依赖...
pip install -r "%~dp0requirements.txt" --quiet
if %errorlevel% neq 0 (
    echo [警告] 部分依赖安装失败，尝试继续启动...
)

echo.
echo [2/2] 启动应用...
echo 浏览器会自动打开，如果没有请手动访问 http://localhost:8501
echo 关闭此窗口即可停止应用。
echo.

streamlit run "%~dp0invoice_ui.py" --server.headless=false
pause
