@echo off
chcp 65001 >nul 2>&1
title 发票识别工具 - 打包脚本

echo ============================================
echo   发票识别工具 - 一键打包（免安装版）
echo ============================================
echo.
echo 此脚本将创建一个完整的免安装发行包，
echo 用户无需安装 Python，双击即可运行。
echo.

set "DIST_DIR=%~dp0dist\发票识别工具"
set "PYTHON_VER=3.11.9"
set "PYTHON_ZIP=python-%PYTHON_VER%-embed-amd64.zip"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VER%/%PYTHON_ZIP%"

:: 检查 curl 是否可用
curl --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 需要 curl 来下载文件，Windows 10+ 自带 curl。
    pause
    exit /b 1
)

:: 创建输出目录
echo [1/5] 创建输出目录...
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%DIST_DIR%"
mkdir "%DIST_DIR%\python"

:: 下载嵌入式 Python
echo [2/5] 下载嵌入式 Python %PYTHON_VER% ...
if not exist "%~dp0dist\%PYTHON_ZIP%" (
    curl -L -o "%~dp0dist\%PYTHON_ZIP%" "%PYTHON_URL%"
    if %errorlevel% neq 0 (
        echo [错误] 下载 Python 失败。请检查网络。
        pause
        exit /b 1
    )
)

:: 解压 Python
echo [3/5] 解压 Python...
powershell -Command "Expand-Archive -Path '%~dp0dist\%PYTHON_ZIP%' -DestinationPath '%DIST_DIR%\python' -Force"

:: 启用 pip：修改 python311._pth 文件，取消注释 import site
powershell -Command "(Get-Content '%DIST_DIR%\python\python311._pth') -replace '#import site','import site' | Set-Content '%DIST_DIR%\python\python311._pth'"

:: 安装 pip
echo [4/5] 安装 pip 并安装依赖（需要几分钟）...
curl -L -o "%DIST_DIR%\python\get-pip.py" "https://bootstrap.pypa.io/get-pip.py"
"%DIST_DIR%\python\python.exe" "%DIST_DIR%\python\get-pip.py" --no-warn-script-location >nul 2>&1

:: 安装项目依赖
"%DIST_DIR%\python\python.exe" -m pip install PyMuPDF pandas openpyxl streamlit streamlit-sortables --no-warn-script-location --quiet
if %errorlevel% neq 0 (
    echo [警告] 部分依赖安装可能失败，请检查网络后重试。
)

:: 复制项目文件
echo [5/5] 复制项目文件...
copy "%~dp0invoice_recognizer.py" "%DIST_DIR%\" >nul
copy "%~dp0invoice_ui.py" "%DIST_DIR%\" >nul
copy "%~dp0requirements.txt" "%DIST_DIR%\" >nul

:: 创建启动脚本
(
echo @echo off
echo chcp 65001 ^>nul 2^>^&1
echo title 发票识别工具
echo echo ============================================
echo echo         发票识别工具 - 启动中...
echo echo ============================================
echo echo.
echo echo 浏览器会自动打开，如果没有请手动访问 http://localhost:8501
echo echo 关闭此窗口即可停止应用。
echo echo.
echo "%%~dp0python\python.exe" -m streamlit run "%%~dp0invoice_ui.py" --server.headless=false --browser.gatherUsageStats=false
echo pause
) > "%DIST_DIR%\启动发票识别.bat"

:: 清理临时文件
del "%DIST_DIR%\python\get-pip.py" >nul 2>&1

echo.
echo ============================================
echo   打包完成！
echo   输出目录: %DIST_DIR%
echo ============================================
echo.
echo 将整个"发票识别工具"文件夹复制给别人即可使用，
echo 对方无需安装任何软件，双击"启动发票识别.bat"即可运行。
echo.
pause
