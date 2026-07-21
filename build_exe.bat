@echo off
chcp 65001 >nul
echo ========================================
echo  打包 Python 脚本为 Windows EXE
echo ========================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8 或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] 安装 PyInstaller...
pip install pyinstaller -q
if errorlevel 1 (
    echo [错误] PyInstaller 安装失败
    pause
    exit /b 1
)

echo [2/4] 安装依赖包...
pip install pandas openpyxl python-calamine -q
if errorlevel 1 (
    echo [错误] 依赖包安装失败
    pause
    exit /b 1
)

echo [3/4] 打包余额汇总脚本...
pyinstaller --onefile --name "一键合并余额汇总" "Python\余额汇总\一键合并余额汇总.py"
if errorlevel 1 (
    echo [错误] 余额汇总脚本打包失败
    pause
    exit /b 1
)

echo [4/4] 打包消耗汇总脚本...
pyinstaller --onefile --name "一键合并消耗汇总" "Python\消耗汇总\总消耗汇总.py"
if errorlevel 1 (
    echo [错误] 消耗汇总脚本打包失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo  打包完成！
echo ========================================
echo.
echo EXE 文件位置:
echo   - dist\一键合并余额汇总.exe
echo   - dist\一键合并消耗汇总.exe
echo.
echo 可以将这两个 exe 文件复制到任意位置使用
echo.
pause
