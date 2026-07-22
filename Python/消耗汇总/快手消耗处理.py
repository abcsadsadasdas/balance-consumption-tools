"""
快手平台消耗处理脚本（磁力、金教、本地生活）

规则：
1. 合并多个Excel（磁力从ZIP解压，金教/本地生活直接读目录下文件）
2. 只保留指定列：时间、账户ID、总花费、现金总花费、前返总花费、后返总花费、
   信用总花费、框返总花费、激励总花费、活动框返总花费、平台激励总花费
3. 后返总花费 加到 前返总花费，删除后返总花费列
4. 活动框返总花费 加到 框返总花费，删除活动框返总花费列
5. 平台激励总花费 加到 激励总花费，删除平台激励总花费列
6. 删除总花费为0的行

用法：
  python 快手消耗处理.py                    # 处理今天的日期目录下所有快手平台
  python 快手消耗处理.py 2026-07-05         # 处理指定日期目录
  python 快手消耗处理.py 2026-07-05 cili    # 只处理磁力平台

目录结构：
  原始文件放在: {日期}/原始文件/快手-磁力- 消耗/     (ZIP文件)
              {日期}/原始文件/快手-金教- 消耗/      (xlsx文件)
              {日期}/原始文件/快手-本地生活- 消耗/  (xlsx文件)
  生成文件输出: {日期}/生成文件/快手-{平台}- 消耗/

依赖: pip install pandas openpyxl
"""

import sys
import zipfile
import traceback
from pathlib import Path
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd


# 脚本所在目录（Python/消耗汇总/）
import os as _os
_exe_dir = _os.environ.get('_EXE_DIR')
if _exe_dir:
    SCRIPT_DIR = Path(_exe_dir)
    PROJECT_ROOT = Path(_os.environ.get('_PROJECT_ROOT', SCRIPT_DIR))
else:
    SCRIPT_DIR = Path(__file__).parent
    PROJECT_ROOT = SCRIPT_DIR.parent.parent

# 模板文件路径
TEMPLATE_PATH = SCRIPT_DIR / "回填数据导入模板_快手.xlsx"

# 需要保留的列
KEEP_COLUMNS = [
    '时间', '账户ID', '总花费', '现金总花费', '前返总花费',
    '后返总花费', '信用总花费', '框返总花费', '激励总花费',
    '活动框返总花费', '平台激励总花费'
]

# 最终输出列（合并后）
OUTPUT_COLUMNS = [
    '时间', '账户ID', '总花费', '现金总花费', '前返总花费',
    '信用总花费', '框返总花费', '激励总花费'
]

# 平台配置: key -> (目录名, 数据来源类型)
PLATFORMS = {
    'cili': ('快手-磁力- 消耗', 'zip'),
    'jinjiao': ('快手-金教- 消耗', 'files'),
    'bendi': ('快手-本地生活- 消耗', 'files'),
}


def read_from_zip(raw_dir: Path) -> pd.DataFrame:
    """从目录中找到ZIP文件并读取所有Excel"""
    zip_files = list(raw_dir.glob("*.zip"))
    if not zip_files:
        raise RuntimeError(f"在 {raw_dir} 下未找到ZIP文件")

    all_data = []
    for zip_path in zip_files:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for file_info in zf.infolist():
                if file_info.filename.endswith('.xlsx') and not file_info.filename.startswith('__MACOSX'):
                    try:
                        with zf.open(file_info) as f:
                            df = pd.read_excel(BytesIO(f.read()), dtype=str)
                            all_data.append(df)
                            print(f"  📄 {file_info.filename}: {len(df)} 行")
                    except Exception as e:
                        print(f"  ⚠️  读取 {file_info.filename} 失败: {e}")

    if not all_data:
        raise RuntimeError(f"ZIP文件中未读取到有效数据")

    return pd.concat(all_data, ignore_index=True)


def read_from_directory(raw_dir: Path) -> pd.DataFrame:
    """从目录中读取所有Excel并合并"""
    xlsx_files = sorted([f for f in raw_dir.glob("*.xlsx") if not f.name.startswith('~$')])
    if not xlsx_files:
        raise RuntimeError(f"在 {raw_dir} 下未找到xlsx文件")

    all_data = []
    for file_path in xlsx_files:
        try:
            df = pd.read_excel(file_path, dtype=str)
            all_data.append(df)
            print(f"  📄 {file_path.name}: {len(df)} 行")
        except Exception as e:
            print(f"  ⚠️  读取 {file_path.name} 失败: {e}")

    if not all_data:
        raise RuntimeError(f"目录中未读取到有效数据")

    return pd.concat(all_data, ignore_index=True)


def process_kuaishou(df: pd.DataFrame, platform_name: str, output_dir: Path, date_str: str) -> dict:
    """处理快手消耗数据"""

    original_count = len(df)
    print(f"\n📊 {platform_name} 原始数据: {original_count} 行")

    # 检查必需列是否存在
    missing_cols = [col for col in KEEP_COLUMNS if col not in df.columns]
    if missing_cols:
        raise RuntimeError(f"数据缺少必需列: {missing_cols}")

    # 只保留需要的列
    df = df[KEEP_COLUMNS].copy()

    # 转换数值列
    numeric_cols = KEEP_COLUMNS[2:]  # 除了时间和账户ID
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # ========== 规则1: 后返总花费 加到 前返总花费，删除后返列 ==========
    merge_count1 = (df['后返总花费'] != 0).sum()
    df['前返总花费'] = df['前返总花费'] + df['后返总花费']
    df.drop(columns=['后返总花费'], inplace=True)
    print(f"规则1 - 后返合并到前返: {merge_count1} 行有后返数据")

    # ========== 规则2: 活动框返总花费 加到 框返总花费，删除活动框返列 ==========
    merge_count2 = (df['活动框返总花费'] != 0).sum()
    df['框返总花费'] = df['框返总花费'] + df['活动框返总花费']
    df.drop(columns=['活动框返总花费'], inplace=True)
    print(f"规则2 - 活动框返合并到框返: {merge_count2} 行有活动框返数据")

    # ========== 规则3: 平台激励总花费 加到 激励总花费，删除平台激励列 ==========
    merge_count3 = (df['平台激励总花费'] != 0).sum()
    df['激励总花费'] = df['激励总花费'] + df['平台激励总花费']
    df.drop(columns=['平台激励总花费'], inplace=True)
    print(f"规则3 - 平台激励合并到激励: {merge_count3} 行有平台激励数据")

    # ========== 规则4: 删除总花费为0的行 ==========
    before = len(df)
    df = df[df['总花费'] != 0].copy()
    print(f"规则4 - 删除总花费为0: 删除 {before - len(df)} 行, 剩余 {len(df)} 行")

    # 保存处理结果
    output_file = output_dir / f"{platform_name}_消耗处理结果_{date_str}.xlsx"
    if output_file.exists():
        output_file.unlink()
    df.to_excel(output_file, index=False, engine='openpyxl')

    # ========== 规则5: 回填模板 ==========
    yesterday = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")
    template_output = output_dir / f"{platform_name}_回填结果_{date_str}.xlsx"

    if TEMPLATE_PATH.is_file():
        # 读取模板表头（快手模板在"快手"sheet）
        template_header = pd.read_excel(TEMPLATE_PATH, sheet_name='快手', dtype=str, nrows=0)
        template_cols = list(template_header.columns)

        # 数据映射
        col_mapping = {
            '*日期': [yesterday] * len(df),
            '*投放账户ID': df['账户ID'].astype(str).values,
            '*总消耗': df['总花费'].values,
            '信用消耗': df['信用总花费'].values,
            '现金消耗': df['现金总花费'].values,
            '返点消耗': df['前返总花费'].values,
            '框返消耗': df['框返总花费'].values,
            '激励消耗': df['激励总花费'].values,
            '备注': '',
        }

        # 按模板列顺序构造DataFrame
        fill_data = {}
        for col in template_cols:
            if col in col_mapping:
                fill_data[col] = col_mapping[col]
            else:
                fill_data[col] = 0
        fill_df = pd.DataFrame(fill_data)

        if template_output.exists():
            template_output.unlink()
        fill_df.to_excel(template_output, sheet_name='快手', index=False, engine='openpyxl')
        print(f"\n✅ 回填模板已保存: {template_output}")
    else:
        print(f"\n⚠️  未找到模板文件: {TEMPLATE_PATH}，跳过回填")

    print(f"\n✅ {platform_name} 处理完成!")
    print(f"   处理结果: {output_file}")
    print(f"   最终行数: {len(df)}")

    return {
        "原始行数": original_count,
        "最终行数": len(df),
        "输出文件": str(output_file),
        "回填文件": str(template_output) if TEMPLATE_PATH.is_file() else None,
    }


def get_target_date_str():
    """获取目标日期字符串"""
    if len(sys.argv) > 1:
        return sys.argv[1]
    return datetime.now().strftime("%Y-%m-%d")


def main():
    date_str = get_target_date_str()
    date_dir = PROJECT_ROOT / date_str

    if not date_dir.is_dir():
        print(f"❌ 未找到日期目录: {date_dir}")
        print("   请先运行 创建日期目录.py 创建目录结构，再放入原始文件。")
        raise RuntimeError("脚本执行失败")

    # 确定要处理哪些平台
    target_platform = sys.argv[2] if len(sys.argv) > 2 else None
    if target_platform and target_platform not in PLATFORMS:
        print(f"❌ 未知平台: {target_platform}，支持: {list(PLATFORMS.keys())}")
        raise RuntimeError("脚本执行失败")

    platforms_to_process = [target_platform] if target_platform else list(PLATFORMS.keys())

    print(f"📅 日期: {date_str}")
    print(f"📂 目录: {date_dir}")
    print()

    results = []
    for platform_key in platforms_to_process:
        folder_name, source_type = PLATFORMS[platform_key]
        raw_dir = date_dir / "原始文件" / folder_name
        output_dir = date_dir / "生成文件" / folder_name

        if not raw_dir.is_dir():
            print(f"⏭️  跳过 {folder_name}: 原始文件目录不存在")
            continue

        # 检查目录是否有文件
        has_files = any(raw_dir.glob("*.zip")) or any(
            f for f in raw_dir.glob("*.xlsx") if not f.name.startswith('~$')
        )
        if not has_files:
            print(f"⏭️  跳过 {folder_name}: 目录为空")
            continue

        output_dir.mkdir(parents=True, exist_ok=True)

        print(f"{'='*60}")
        print(f"🔄 开始处理 {folder_name}...")

        try:
            # 读取数据
            if source_type == 'zip':
                df = read_from_zip(raw_dir)
            else:
                df = read_from_directory(raw_dir)

            # 处理数据
            result = process_kuaishou(df, folder_name, output_dir, date_str)
            results.append((folder_name, result, None))
        except Exception as e:
            print(f"❌ {folder_name} 处理失败: {e}")
            results.append((folder_name, None, str(e)))

        print()

    # 汇总
    print("=" * 60)
    print("📋 处理汇总:")
    for name, result, error in results:
        if result:
            print(f"  ✅ {name}: {result['原始行数']} → {result['最终行数']} 行")
        else:
            print(f"  ❌ {name}: {error}")

    # 写入失败日志
    failures = [(name, err) for name, _, err in results if err]
    if failures:
        log_dir = date_dir / "失败文件与日志"
        log_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = log_dir / f"快手消耗处理日志_{timestamp}.txt"
        lines = [f"快手消耗处理日志 - {date_str}", f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "=" * 60, ""]
        for name, err in failures:
            lines.append(f"❌ {name} - 失败")
            lines.append(f"  原因: {err}")
            lines.append("")
        log_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n⚠️  失败日志已保存: {log_path}")


if __name__ == "__main__":
    main()
