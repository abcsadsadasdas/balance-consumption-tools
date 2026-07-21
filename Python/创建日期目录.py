"""
创建每日工作目录脚本

功能：
  在项目根目录（py 自动化脚本/）下创建以当天日期命名的文件夹，
  包含余额汇总和消耗汇总相关目录：

  2026-07-07/
  ├── 原始文件/
  │   ├── 磁力- 余额汇总/          （含 余额导入模板.xlsx）
  │   ├── 广点通 ADQ- 消耗/
  │   ├── 快手-磁力- 消耗/
  │   ├── 快手-金教- 消耗/
  │   ├── 快手-本地生活- 消耗/
  │   ├── 抖音-竞价- 消耗/         （侠客行+新美 合并）
  │   ├── 抖音-品牌- 消耗/         （侠客行+新美 合并）
  │   ├── 千川-品牌- 消耗/         （直接上传DSP）
  │   ├── 千川-竞价- 消耗/         （直接上传DSP）
  │   ├── 抖音-本地推- 消耗/       （直接上传DSP）
  │   ├── 抖音-巨懂车- 消耗/       （直接上传DSP）
  │   └── 千川-巨懂车- 消耗/       （直接上传DSP）
  ├── 生成文件/
  │   ├── 磁力- 余额汇总/
  │   ├── 广点通 ADQ- 消耗/
  │   ├── 快手-磁力- 消耗/
  │   ├── 快手-金教- 消耗/
  │   ├── 快手-本地生活- 消耗/
  │   ├── 抖音-竞价- 消耗/
  │   ├── 抖音-品牌- 消耗/
  │   ├── 千川-品牌- 消耗/
  │   ├── 千川-竞价- 消耗/
  │   ├── 抖音-本地推- 消耗/
  │   ├── 抖音-巨懂车- 消耗/
  │   └── 千川-巨懂车- 消耗/
  └── 失败文件与日志/

用法：
  python 创建日期目录.py            # 创建今天的日期目录
  python 创建日期目录.py 2026-07-05  # 创建指定日期的目录
"""

import sys
import shutil
from datetime import date
from pathlib import Path

# 获取 exe 或脚本的真实路径
if getattr(sys, 'frozen', False):
    # PyInstaller 打包后的 exe
    EXE_DIR = Path(sys.executable).resolve().parent
else:
    # 直接运行 Python 脚本
    EXE_DIR = Path(__file__).resolve().parent

# 脚本位于 Python/，项目根目录是其上一级
SCRIPT_DIR = EXE_DIR
PROJECT_ROOT = SCRIPT_DIR.parent

# 模板文件：Python/余额汇总/模板/余额导入模板.xlsx
TEMPLATE_PATH = SCRIPT_DIR / "余额汇总" / "模板" / "余额导入模板.xlsx"

# 余额汇总平台（需要复制余额导入模板）
BALANCE_PLATFORM_FOLDERS = [
    "磁力- 余额汇总",
]

# 消耗平台（不需要模板）
CONSUME_PLATFORM_FOLDERS = [
    "广点通 ADQ- 消耗",
    "快手-磁力- 消耗",
    "快手-金教- 消耗",
    "快手-本地生活- 消耗",
    "抖音-竞价- 消耗",
    "抖音-品牌- 消耗",
    "千川-品牌- 消耗",
    "千川-竞价- 消耗",
    "抖音-本地推- 消耗",
    "抖音-巨懂车- 消耗",
    "千川-巨懂车- 消耗",
    "金牛-华北- 消耗",
    "金牛-广分- 消耗",
    "金牛-澳比- 消耗",
    "支付宝- 消耗",
    "陵致长风-巨量- 消耗",
    "陵致长风-千川- 消耗",
    "B站-信息流- 消耗",
    "B站-商业起飞- 消耗",
]


def get_target_date() -> str:
    """获取目标日期字符串，支持命令行传参，默认今天"""
    if len(sys.argv) > 1:
        date_str = sys.argv[1]
        try:
            datetime_obj = date.fromisoformat(date_str)
        except ValueError:
            print(f"❌ 日期格式不正确: {date_str}（应为 YYYY-MM-DD）")
            sys.exit(1)
        return datetime_obj.isoformat()
    return date.today().isoformat()


def create_daily_structure(date_str: str):
    date_dir = PROJECT_ROOT / date_str

    if date_dir.exists():
        print(f"⚠️  目录已存在: {date_dir}")
        print("   跳过创建，不会覆盖已有文件。")
        return

    print(f"创建日期目录: {date_dir}")

    raw_dir = date_dir / "原始文件"
    output_dir = date_dir / "生成文件"
    log_dir = date_dir / "失败文件与日志"

    # 创建余额汇总平台目录（含模板）
    for platform in BALANCE_PLATFORM_FOLDERS:
        (raw_dir / platform).mkdir(parents=True, exist_ok=True)
        (output_dir / platform).mkdir(parents=True, exist_ok=True)
        print(f"  ✅ 原始文件/{platform}")
        print(f"  ✅ 生成文件/{platform}")

        # 每个余额平台的原始文件目录下放一份余额导入模板
        if TEMPLATE_PATH.is_file():
            shutil.copy2(TEMPLATE_PATH, raw_dir / platform / "余额导入模板.xlsx")
        else:
            print(f"  ⚠️  未找到模板文件，跳过复制: {TEMPLATE_PATH}")

    # 创建消耗平台目录（不需要模板）
    for platform in CONSUME_PLATFORM_FOLDERS:
        (raw_dir / platform).mkdir(parents=True, exist_ok=True)
        (output_dir / platform).mkdir(parents=True, exist_ok=True)
        print(f"  ✅ 原始文件/{platform}")
        print(f"  ✅ 生成文件/{platform}")

    log_dir.mkdir(parents=True, exist_ok=True)
    print(f"  ✅ 失败文件与日志")

    print(f"\n✅ 目录结构创建完成: {date_dir.name}/")


def main():
    date_str = get_target_date()
    create_daily_structure(date_str)


if __name__ == "__main__":
    main()
