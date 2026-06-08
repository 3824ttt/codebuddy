@echo off
chcp 65001 >nul 2>&1
title 读者荐读系统

cd /d "%~dp0"

echo ========================================
echo   读者荐读系统 - 启动中...
echo ========================================
echo.

:: 设置 Python 路径
set PYTHON=C:\Users\wang\.workbuddy\binaries\python\versions\3.14.3\python.exe

if not exist "%PYTHON%" (
    echo 错误: Python 未找到
    pause
    exit /b 1
)

:: 检查并安装 Flask
echo 检查依赖...
"%PYTHON%" -c "import flask" 2>nul
if errorlevel 1 (
    echo 正在安装 Flask...
    "%PYTHON%" -m pip install flask -q
)

echo.
echo 系统启动！
echo.
echo    读者端: http://localhost:5000
echo    管理端: http://localhost:5000/admin
echo    管理员密码: admin123
echo.
echo ========================================

"%PYTHON%" "%~dp0backend\app.py"

pause
