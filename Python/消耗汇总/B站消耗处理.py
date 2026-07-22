"""
B站消耗处理脚本（信息流、商业起飞）

处理规则：
1. 筛选列：投放账户ID(账号id)、总消耗(整体实收总消耗)、现金消耗、返货消耗
2. 计算非有效消耗 = 总消耗 - 现金消耗 - 返货消耗
3. 输出：投放账户ID、总消耗、现金消耗、非有效消耗、返货消耗
4. 回填模板（待模板提供）

用法：
  python B站消耗处理.py                    # 处理今天的日期目录
  python B站消耗处理.py 2026-07-08         # 处理指定日期目录

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

# 处理的平台
PLATFORMS = {
    'xinxiliu': 'B站-信息流- 消耗',
    'qifei': 'B站-商业起飞- 消耗',
}


def read_all_xlsx(raw_dir: Path) -> pd.DataFrame:
    """读取目录下所有Excel文件并合并"""
    all_data = []
    for f in sorted(raw_dir.glob("*.xlsx")):
        if f.name.startswith('~$'):
            continue
        df = pd.read_excel(f, dtype=str)
        all_data.append(df)
        print(f"  📄 {f.name}: {len(df)} 行")

    if not all_data:
        raise RuntimeError(f"在 {raw_dir} 下未找到xlsx文件")
    return pd.concat(all_data, ignore_index=True)


def process_bilibili(df: pd.DataFrame, platform_name: str, output_dir: Path, date_str: str) -> dict:
    """处理B站消耗数据"""

    original_count = len(df)
    print(f"\n📊 {platform_name} 原始数据: {original_count} 行")

    # 检查必需列
    required = ['账号id', '整体实收总消耗', '现金消耗', '返货消耗']
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise RuntimeError(f"数据缺少必需列: {missing}")

    # 转换数值列
    for col in ['整体实收总消耗', '现金消耗', '返货消耗']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 计算非有效消耗 = 总消耗 - 现金消耗 - 返货消耗
    df['非有效消耗'] = df['整体实收总消耗'] - df['现金消耗'] - df['返货消耗']

    # 删除总消耗为0的行
    before = len(df)
    df = df[df['整体实收总消耗'] != 0].copy()
    if before != len(df):
        print(f"   删除总消耗为0: 删除 {before - len(df)} 行")

    # 构造输出
    result = pd.DataFrame({
        '投放账户ID': df['账号id'].astype(str).values,
        '总消耗': df['整体实收总消耗'].values,
        '现金消耗': df['现金消耗'].values,
        '非有效消耗': df['非有效消耗'].values,
        '返货消耗': df['返货消耗'].values,
    })

    # 保存处理结果
    output_file = output_dir / f"{platform_name}_处理结果_{date_str}.xlsx"
    if output_file.exists():
        output_file.unlink()
    result.to_excel(output_file, index=False, engine='openpyxl')

    # 回填结果（带日期）- 使用B站信息流模板
    yesterday = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")
    template_output = output_dir / f"{platform_name}_回填结果_{date_str}.xlsx"

    TEMPLATE_BILIBILI = SCRIPT_DIR / "回填数据导入模板_B站-信息流.xlsx"

    if TEMPLATE_BILIBILI.is_file():
        # 读取模板表头（B站信息流模板在"其他"sheet）
        template_header = pd.read_excel(TEMPLATE_BILIBILI, sheet_name='其他', dtype=str, nrows=0)
        template_cols = list(template_header.columns)

        col_mapping = {
            '*日期': [yesterday] * len(result),
            '*投放账户ID': result['投放账户ID'].values,
            '*总消耗': result['总消耗'].values,
            '现金消耗': result['现金消耗'].values,
            '信用消耗': 0,
            '非有效消耗': result['非有效消耗'].values,
            '返货消耗': result['返货消耗'].values,
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
            '*日期': [yesterday] * len(result),
            '*投放账户ID': result['投放账户ID'].values,
            '*总消耗': result['总消耗'].values,
            '现金消耗': result['现金消耗'].values,
            '信用消耗': 0,
            '非有效消耗': result['非有效消耗'].values,
            '返货消耗': result['返货消耗'].values,
            '备注': '',
        })

    if template_output.exists():
        template_output.unlink()
    fill_df.to_excel(template_output, sheet_name='B站-信息流', index=False, engine='openpyxl')
    print(f"\n✅ 回填结果已保存: {template_output}")

    print(f"✅ {platform_name} 处理完成! 最终行数: {len(result)}")
    return {"原始行数": original_count, "最终行数": len(result), "输出文件": str(output_file)}


def get_target_date_str():
    if len(sys.argv) > 1:
        return sys.argv[1]
    return datetime.now().strftime("%Y-%m-%d")


def main():
    date_str = get_target_date_str()
    date_dir = PROJECT_ROOT / date_str

    if not date_dir.is_dir():
        print(f"❌ 未找到日期目录: {date_dir}")
        raise RuntimeError("脚本执行失败")

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

        has_files = any(f for f in raw_dir.glob("*.xlsx") if not f.name.startswith('~$'))
        if not has_files:
            print(f"⏭️  跳过 {folder_name}: 目录为空")
            continue

        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"{'='*60}")
        print(f"🔄 开始处理 {folder_name}...")

        try:
            df = read_all_xlsx(raw_dir)
            result = process_bilibili(df, folder_name, output_dir, date_str)
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
        log_path = log_dir / f"B站消耗处理日志_{timestamp}.txt"
        lines = [f"B站消耗处理日志 - {date_str}", f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "=" * 60, ""]
        for name, err in failures:
            lines.append(f"❌ {name} - 失败")
            lines.append(f"  原因: {err}")
            lines.append("")
        log_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n⚠️  失败日志已保存: {log_path}")


if __name__ == "__main__":
    main()
