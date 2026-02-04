@echo off
chcp 65001 >nul
echo ============================================================
echo 公文格式化工具 - 快速启动
echo ============================================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.7+
    pause
    exit /b 1
)

echo [1/3] 检查依赖包...
pip show flask >nul 2>&1
if errorlevel 1 (
    echo [安装] 正在安装 Flask...
    pip install flask flask-cors python-docx
)

echo [2/3] 启动 Web 服务器...
echo.
echo ============================================================
echo 访问地址: http://localhost:5000
echo 按 Ctrl+C 停止服务器
echo ============================================================
echo.

python web_server.py

pause
