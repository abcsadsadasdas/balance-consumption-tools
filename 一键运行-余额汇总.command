#!/bin/bash
# 双击此文件即可一键运行余额汇总处理（处理今天的日期目录）

set -u
ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT_PATH="$ROOT_DIR/Python/余额汇总/一键合并余额汇总.py"

echo "=========================================="
echo "  一键运行余额汇总处理"
echo "=========================================="
echo ""

if [ ! -f "$SCRIPT_PATH" ]; then
    echo "找不到脚本：$SCRIPT_PATH"
    echo ""
    read -r -n 1 -s -p "按任意键关闭窗口..."
    echo ""
    exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
    echo "未找到 python3，请先安装 Python 3.9 或更高版本。"
    echo ""
    read -r -n 1 -s -p "按任意键关闭窗口..."
    echo ""
    exit 1
fi

if ! python3 -c "import pandas, openpyxl, python_calamine" >/dev/null 2>&1; then
    echo "缺少依赖，请先执行："
    echo "python3 -m pip install pandas openpyxl python-calamine"
    echo ""
    read -r -n 1 -s -p "按任意键关闭窗口..."
    echo ""
    exit 1
fi

python3 "$SCRIPT_PATH" "$@"
EXIT_CODE=$?

echo ""
if [ "$EXIT_CODE" -eq 0 ]; then
    echo "余额汇总处理成功。"
else
    echo "余额汇总处理失败，退出码：$EXIT_CODE"
fi
read -r -n 1 -s -p "按任意键关闭窗口..."
echo ""
exit "$EXIT_CODE"
