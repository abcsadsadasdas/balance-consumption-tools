#!/bin/bash
# 双击此文件即可一键创建今天的日期目录（原始文件/生成文件/失败文件与日志 骨架）
cd "$(dirname "$0")/Python"
python3 "创建日期目录.py"
echo ""
echo "按任意键关闭窗口..."
read -n 1
