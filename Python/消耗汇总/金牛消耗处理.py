"""
金牛消耗处理脚本（华北、广分、澳比）

处理平台：
  - 金牛-华北（规则处理+回填）
  - 金牛-广分（规则处理+回填）
  - 金牛-澳比（简化处理+回填，含非有效消耗计算）

金牛-华北/金牛-广分 处理规则：
1. 删除不需要的列
2. 后返总花费 加到 前返总花费，删除后返列
3. 活动框返总花费 加到 框返总花费，删除活动框返列
4. 平台激励总花费 加到 激励总花费，删除平台激励列
5. 删除总花费为0的行
6. 回填模板（日期为昨天）

金牛-澳比 处理规则：
1. 删除不需要的列
2. 只保留：时间、账户ID、总花费、现金总花费、信用总花费
3. 回填模板，额外计算：非有效消耗 = 总消耗 - 现金消耗 - 信用消耗

用法：
  python 金牛消耗处理.py                    # 处理今天的日期目录
  python 金牛消耗处理.py 2026-07-07         # 处理指定日期目录

依赖: pip install pandas openpyxl
"""

import sys
import traceback
from pathlib import Path
from datetime import datetime, timedelta

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
TEMPLATE_KUAISHOU = SCRIPT_DIR / "回填数据导入模板_快手.xlsx"
TEMPLATE_OTHER = SCRIPT_DIR / "回填数据导入模板_其他.xlsx"

# 需要处理的平台
PLATFORMS = {
    'huabei': '金牛-华北- 消耗',
    'guangfen': '金牛-广分- 消耗',
    'aobi': '金牛-澳比- 消耗',
}

# 需要删除的列
DELETE_COLUMNS = [
    '账户名称', '产品名', '企业名称', '账户类型', '快手ID',
    '快手昵称', '端口ID', '端口名称', '一级行业', '二级行业',
    '总账户可用余额', '总账户预占余额'
]


def read_all_files(raw_dir: Path) -> pd.DataFrame:
    """读取目录下所有Excel/CSV文件并合并"""
    all_data = []

    # xlsx文件
    for f in sorted(raw_dir.glob("*.xlsx")):
        if f.name.startswith('~$'):
            continue
        df = pd.read_excel(f, dtype=str)
        all_data.append(df)
        print(f"  📄 {f.name}: {len(df)} 行")

    # csv文件
    for f in sorted(raw_dir.glob("*.csv")):
        df = pd.read_csv(f, dtype=str)
        all_data.append(df)
        print(f"  📄 {f.name}: {len(df)} 行")

    if not all_data:
        raise RuntimeError(f"在 {raw_dir} 下未找到数据文件")

    return pd.concat(all_data, ignore_index=True)


def process_jinniu_standard(df: pd.DataFrame, platform_name: str, output_dir: Path, date_str: str) -> dict:
    """处理金牛-华北/金牛-广分（标准规则）"""

    original_count = len(df)
    print(f"\n📊 {platform_name} 原始数据: {original_count} 行")

    # 删除不需要的列
    cols_to_drop = [col for col in DELETE_COLUMNS if col in df.columns]
    df = df.drop(columns=cols_to_drop)
    print(f"规则1 - 删除无用列: 删除 {len(cols_to_drop)} 列")

    # 转换数值列
    numeric_cols = ['总花费', '现金总花费', '前返总花费', '后返总花费',
                    '信用总花费', '框返总花费', '激励总花费', '活动框返总花费', '平台激励总花费']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 规则2: 后返→前返
    if '后返总花费' in df.columns and '前返总花费' in df.columns:
        merge_count = (df['后返总花费'] != 0).sum()
        df['前返总花费'] = df['前返总花费'] + df['后返总花费']
        df.drop(columns=['后返总花费'], inplace=True)
        print(f"规则2 - 后返合并到前返: {merge_count} 行有后返数据")

    # 规则3: 活动框返→框返
    if '活动框返总花费' in df.columns and '框返总花费' in df.columns:
        merge_count = (df['活动框返总花费'] != 0).sum()
        df['框返总花费'] = df['框返总花费'] + df['活动框返总花费']
        df.drop(columns=['活动框返总花费'], inplace=True)
        print(f"规则3 - 活动框返合并到框返: {merge_count} 行有活动框返数据")

    # 规则4: 平台激励→激励
    if '平台激励总花费' in df.columns and '激励总花费' in df.columns:
        merge_count = (df['平台激励总花费'] != 0).sum()
        df['激励总花费'] = df['激励总花费'] + df['平台激励总花费']
        df.drop(columns=['平台激励总花费'], inplace=True)
        print(f"规则4 - 平台激励合并到激励: {merge_count} 行有平台激励数据")

    # 规则5: 删除总花费为0
    before = len(df)
    df = df[df['总花费'] != 0].copy()
    print(f"规则5 - 删除总花费为0: 删除 {before - len(df)} 行, 剩余 {len(df)} 行")

    # 保存处理结果
    output_file = output_dir / f"{platform_name}_处理结果_{date_str}.xlsx"
    if output_file.exists():
        output_file.unlink()
    df.to_excel(output_file, index=False, engine='openpyxl')

    # 规则6: 回填模板
    yesterday = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")
    template_output = output_dir / f"{platform_name}_回填结果_{date_str}.xlsx"

    if TEMPLATE_KUAISHOU.is_file():
        # 读取模板表头（快手模板在"快手"sheet）
        template_header = pd.read_excel(TEMPLATE_KUAISHOU, sheet_name='快手', dtype=str, nrows=0)
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
        print(f"\n⚠️  未找到模板文件: {TEMPLATE_KUAISHOU}，跳过回填")

    print(f"✅ {platform_name} 处理完成! 最终行数: {len(df)}")
    return {"原始行数": original_count, "最终行数": len(df), "输出文件": str(output_file)}


def process_jinniu_aobi(df: pd.DataFrame, platform_name: str, output_dir: Path, date_str: str) -> dict:
    """处理金牛-澳比（简化规则+非有效消耗）"""

    original_count = len(df)
    print(f"\n📊 {platform_name} 原始数据: {original_count} 行")

    # 规则1: 删除不需要的列
    cols_to_drop = [col for col in DELETE_COLUMNS if col in df.columns]
    df = df.drop(columns=cols_to_drop)
    print(f"规则1 - 删除无用列: 删除 {len(cols_to_drop)} 列")

    # 规则2: 只保留 时间、账户ID、总花费、现金总花费、信用总花费
    keep_cols = ['时间', '账户ID', '总花费', '现金总花费', '信用总花费']
    missing = [col for col in keep_cols if col not in df.columns]
    if missing:
        raise RuntimeError(f"数据缺少必需列: {missing}")
    df = df[keep_cols].copy()

    # 转换数值列
    for col in ['总花费', '现金总花费', '信用总花费']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 删除总花费为0
    before = len(df)
    df = df[df['总花费'] != 0].copy()
    print(f"   删除总花费为0: 删除 {before - len(df)} 行, 剩余 {len(df)} 行")

    # 保存处理结果
    output_file = output_dir / f"{platform_name}_处理结果_{date_str}.xlsx"
    if output_file.exists():
        output_file.unlink()
    df.to_excel(output_file, index=False, engine='openpyxl')

    # 规则3: 回填模板（含非有效消耗计算）- 使用其他模板
    yesterday = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")
    template_output = output_dir / f"{platform_name}_回填结果_{date_str}.xlsx"

    # 非有效消耗 = 总消耗 - 现金消耗 - 信用消耗
    non_effective = df['总花费'] - df['现金总花费'] - df['信用总花费']

    if TEMPLATE_OTHER.is_file():
        # 读取模板表头（其他模板在"其他"sheet）
        template_header = pd.read_excel(TEMPLATE_OTHER, sheet_name='其他', dtype=str, nrows=0)
        template_cols = list(template_header.columns)

        col_mapping = {
            '*日期': [yesterday] * len(df),
            '*投放账户ID': df['账户ID'].astype(str).values,
            '*总消耗': df['总花费'].values,
            '现金消耗': df['现金总花费'].values,
            '信用消耗': df['信用总花费'].values,
            '非有效消耗': non_effective.values,
            '备注': '',
        }

        fill_data = {}
        for col in template_cols:
            if col in col_mapping:
                fill_data[col] = col_mapping[col]
            else:
                fill_data[col] = 0
        fill_df = pd.DataFrame(fill_data)
    else:
        fill_df = pd.DataFrame({
            '*日期': [yesterday] * len(df),
            '*投放账户ID': df['账户ID'].astype(str).values,
            '*总消耗': df['总花费'].values,
            '现金消耗': df['现金总花费'].values,
            '信用消耗': df['信用总花费'].values,
            '非有效消耗': non_effective.values,
            '备注': '',
        })

    if template_output.exists():
        template_output.unlink()
    fill_df.to_excel(template_output, sheet_name='其他', index=False, engine='openpyxl')
    print(f"\n✅ 回填模板已保存: {template_output}")

    print(f"✅ {platform_name} 处理完成! 最终行数: {len(df)}")
    return {"原始行数": original_count, "最终行数": len(df), "输出文件": str(output_file)}


def get_target_date_str():
    if len(sys.argv) > 1:
        return sys.argv[1]
    return datetime.now().strftime("%Y-%m-%d")


def main():
    date_str = get_target_date_str()
    date_dir = PROJECT_ROOT / date_str

    if not date_dir.is_dir():
        print(f"❌ 未找到日期目录: {date_dir}")
        sys.exit(1)

    print(f"📅 日期: {date_str}")
    print(f"📂 目录: {date_dir}")
    print()

    results = []

    for key, folder_name in PLATFORMS.items():
        raw_dir = date_dir / "原始文件" / folder_name
        output_dir = date_dir / "生成文件" / folder_name

        if not raw_dir.is_dir():
            print(f"⏭️  跳过 {folder_name}: 目录不存在")
            continue

        has_files = any(raw_dir.glob("*.xlsx")) or any(raw_dir.glob("*.csv"))
        if not has_files:
            print(f"⏭️  跳过 {folder_name}: 目录为空")
            continue

        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"{'='*60}")
        print(f"🔄 开始处理 {folder_name}...")

        try:
            df = read_all_files(raw_dir)
            if key == 'aobi':
                result = process_jinniu_aobi(df, folder_name, output_dir, date_str)
            else:
                result = process_jinniu_standard(df, folder_name, output_dir, date_str)
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
        log_path = log_dir / f"金牛消耗处理日志_{timestamp}.txt"
        lines = [f"金牛消耗处理日志 - {date_str}", f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "=" * 60, ""]
        for name, err in failures:
            lines.append(f"❌ {name} - 失败")
            lines.append(f"  原因: {err}")
            lines.append("")
        log_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n⚠️  失败日志已保存: {log_path}")


if __name__ == "__main__":
    main()
