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

echo [1/5] 安装 PyInstaller...
pip install pyinstaller -q
if errorlevel 1 (
    echo [错误] PyInstaller 安装失败
    pause
    exit /b 1
)

echo [2/5] 安装依赖包...
pip install pandas openpyxl python-calamine -q
if errorlevel 1 (
    echo [错误] 依赖包安装失败
    pause
    exit /b 1
)

echo [3/5] 打包余额汇总脚本...
pyinstaller --onefile --name "balance-merge" "Python\余额汇总\一键合并余额汇总.py"
if errorlevel 1 (
    echo [错误] 余额汇总脚本打包失败
    pause
    exit /b 1
)

echo [4/5] 打包消耗汇总脚本...
pyinstaller --onefile --name "consumption-merge" "Python\消耗汇总\总消耗汇总.py"
if errorlevel 1 (
    echo [错误] 消耗汇总脚本打包失败
    pause
    exit /b 1
)

echo [5/5] 打包创建日期目录脚本...
pyinstaller --onefile --name "create-directory" "Python\创建日期目录.py"
if errorlevel 1 (
    echo [错误] 创建日期目录脚本打包失败
    pause
    exit /b 1
)

echo.
echo 复制模板文件...
copy "Python\余额汇总\模板\余额导入模板.xlsx" "dist\余额导入模板.xlsx"

echo.
echo ========================================
echo  打包完成！
echo ========================================
echo.
echo EXE 文件位置:
echo   - dist\balance-merge.exe
echo   - dist\consumption-merge.exe
echo   - dist\create-directory.exe
echo   - dist\余额导入模板.xlsx
echo.
echo 可以将这四个文件复制到任意位置使用
echo 日期目录会生成在 exe 所在目录下
echo.
pause
