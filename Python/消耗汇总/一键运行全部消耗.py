"""
一键运行全部消耗处理脚本

功能：
  依次运行 7 个平台的消耗处理脚本：
  1. 快手消耗处理（磁力、金教、本地生活）
  2. 金牛消耗处理（华北、广分、澳比）
  3. 抖音千川消耗处理（竞价、品牌、千川、巨懂车、本地推）
  4. 陵致长风消耗处理（巨量、千川）
  5. 广点通ADQ消耗处理
  6. 支付宝消耗处理
  7. B站消耗处理（信息流、商业起飞）

用法：
  python 一键运行全部消耗.py                    # 处理今天的日期目录
  python 一键运行全部消耗.py 2026-07-08         # 处理指定日期目录

依赖: pip install pandas openpyxl
"""

import sys
import os
from pathlib import Path
from datetime import datetime

# 获取 exe 或脚本的真实路径（必须在 import 平台脚本之前设置）
if getattr(sys, 'frozen', False):
    # PyInstaller 打包后的 exe：exe 所在目录就是项目根目录
    PROJECT_ROOT = Path(sys.executable).resolve().parent
    SCRIPT_DIR = PROJECT_ROOT
else:
    # 直接运行 Python 脚本：脚本在 Python/消耗汇总/ 下
    SCRIPT_DIR = Path(__file__).resolve().parent
    PROJECT_ROOT = SCRIPT_DIR.parent.parent

# 设置环境变量，让被导入的平台脚本能找到 exe 目录
os.environ['_EXE_DIR'] = str(SCRIPT_DIR)
os.environ['_PROJECT_ROOT'] = str(PROJECT_ROOT)

# 显式导入所有平台脚本，让 PyInstaller 能检测到依赖
import 快手消耗处理
import 金牛消耗处理
import 抖音千川消耗处理
import 陵致长风消耗处理
import 广点通ADQ_消耗处理
import 支付宝消耗处理
import B站消耗处理


def get_target_date_str():
    """获取目标日期"""
    if len(sys.argv) > 1:
        return sys.argv[1]
    return datetime.now().strftime("%Y-%m-%d")


def main():
    date_str = get_target_date_str()
    
    print("=" * 50)
    print("  一键运行全部消耗处理")
    print("=" * 50)
    print()
    print(f"📅 日期: {date_str}")
    print(f"📂 目录: {PROJECT_ROOT / date_str}")
    print()
    
    # 检查日期目录是否存在
    date_dir = PROJECT_ROOT / date_str
    if not date_dir.is_dir():
        print(f"❌ 未找到日期目录: {date_dir}")
        sys.exit(1)
    
    # 各平台脚本映射
    scripts = [
        (快手消耗处理, "快手消耗处理"),
        (金牛消耗处理, "金牛消耗处理"),
        (抖音千川消耗处理, "抖音千川消耗处理"),
        (陵致长风消耗处理, "陵致长风消耗处理"),
        (广点通ADQ_消耗处理, "广点通ADQ消耗处理"),
        (支付宝消耗处理, "支付宝消耗处理"),
        (B站消耗处理, "B站消耗处理"),
    ]
    
    success_count = 0
    fail_count = 0
    
    for idx, (module, display_name) in enumerate(scripts, 1):
        print(f"【{idx}/{len(scripts)}】{display_name}...")
        try:
            # 设置 sys.argv 为 [脚本名, 日期]
            old_argv = sys.argv.copy()
            sys.argv = [display_name, date_str]
            
            try:
                # 调用 main 函数
                module.main()
                success_count += 1
            finally:
                # 恢复 sys.argv
                sys.argv = old_argv
            
        except Exception as e:
            print(f"  ❌ 运行失败: {e}")
            import traceback
            traceback.print_exc()
            fail_count += 1
        
        print()
    
    print("=" * 50)
    print(f"  全部消耗处理完成！")
    print(f"  ✅ 成功: {success_count} 个平台")
    if fail_count > 0:
        print(f"  ❌ 失败: {fail_count} 个平台")
    print("=" * 50)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        pass
    except Exception as e:
        print(f"\n❌ 程序异常: {e}")
        import traceback
        traceback.print_exc()
    finally:
        input("\n按回车键退出...")
