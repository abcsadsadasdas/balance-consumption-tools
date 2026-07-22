"""
支付宝消耗处理脚本

处理规则：
1. 筛选"支付宝账号"列，为空或"-"的，把"商家pid"数据复制到"支付宝账号"
2. 回填模板

用法：
  python 支付宝消耗处理.py                    # 处理今天的日期目录
  python 支付宝消耗处理.py 2026-07-07         # 处理指定日期目录

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

PLATFORM_NAME = "支付宝- 消耗"


def read_all_files(raw_dir: Path) -> pd.DataFrame:
    """读取目录下所有Excel/CSV文件并合并"""
    all_data = []

    for f in sorted(raw_dir.glob("*.csv")):
        df = pd.read_csv(f, dtype=str)
        all_data.append(df)
        print(f"  📄 {f.name}: {len(df)} 行")

    for f in sorted(raw_dir.glob("*.xlsx")):
        if f.name.startswith('~$'):
            continue
        df = pd.read_excel(f, dtype=str)
        all_data.append(df)
        print(f"  📄 {f.name}: {len(df)} 行")

    if not all_data:
        raise RuntimeError(f"在 {raw_dir} 下未找到数据文件")

    return pd.concat(all_data, ignore_index=True)


def process_zhifubao(df: pd.DataFrame, output_dir: Path, date_str: str) -> dict:
    """处理支付宝消耗数据"""

    original_count = len(df)
    print(f"\n📊 支付宝 原始数据: {original_count} 行")

    # 检查必需列
    required = ['支付宝账号', '商家pid']
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise RuntimeError(f"数据缺少必需列: {missing}")

    # 规则1: 支付宝账号为空或"-"的，用商家pid填充
    mask = df['支付宝账号'].isna() | (df['支付宝账号'].str.strip() == '') | (df['支付宝账号'].str.strip() == '-')
    fill_count = mask.sum()
    df.loc[mask, '支付宝账号'] = df.loc[mask, '商家pid']
    print(f"规则1 - 填充支付宝账号: {fill_count} 行从商家pid补充")

    # 保存处理结果
    output_file = output_dir / f"支付宝_处理结果_{date_str}.xlsx"
    if output_file.exists():
        output_file.unlink()
    df.to_excel(output_file, index=False, engine='openpyxl')

    # 规则2: 回填模板（支付宝模板格式）
    yesterday = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")
    template_output = output_dir / f"支付宝_回填结果_{date_str}.xlsx"

    TEMPLATE_ZFB = SCRIPT_DIR / "回填数据导入模板_支付宝.xlsx"

    # 转换数值列 - 支付宝的消耗列名为"扣款金额（元）"
    consume_col = None
    for col_name in ['扣款金额（元）', '消费', '消耗', '总消耗', '花费']:
        if col_name in df.columns:
            consume_col = col_name
            break

    # 现金和授信列
    cash_col = None
    for col_name in ['现金（元）', '现金消耗', '现金']:
        if col_name in df.columns:
            cash_col = col_name
            break

    credit_col = None
    for col_name in ['授信（元）', '信用消耗', '授信']:
        if col_name in df.columns:
            credit_col = col_name
            break

    if consume_col:
        # 处理千分位格式（如 1,435.90）
        df[consume_col] = df[consume_col].astype(str).str.replace(',', '', regex=False)
        df[consume_col] = pd.to_numeric(df[consume_col], errors='coerce').fillna(0)
    if cash_col:
        df[cash_col] = df[cash_col].astype(str).str.replace(',', '', regex=False)
        df[cash_col] = pd.to_numeric(df[cash_col], errors='coerce').fillna(0)
    if credit_col:
        df[credit_col] = df[credit_col].astype(str).str.replace(',', '', regex=False)
        df[credit_col] = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)

    if TEMPLATE_ZFB.is_file() and consume_col:
        # 读取模板表头（支付宝模板在"其他"sheet）
        template_header = pd.read_excel(TEMPLATE_ZFB, sheet_name='其他', dtype=str, nrows=0)
        template_cols = list(template_header.columns)

        # 计算非有效消耗 = 总消耗 - 现金 - 授信
        cash_values = df[cash_col].values if cash_col else 0
        credit_values = df[credit_col].values if credit_col else 0
        non_effective = df[consume_col].values - (cash_values if cash_col else 0) - (credit_values if credit_col else 0)

        col_mapping = {
            '*日期': [yesterday] * len(df),
            '*支付宝账户': df['支付宝账号'].astype(str).str.strip().values,
            '*总消耗': df[consume_col].values,
            '现金消耗': cash_values if cash_col else 0,
            '信用消耗': credit_values if credit_col else 0,
            '非有效消耗': non_effective,
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
        # fallback: 直接输出处理后的数据
        fill_df = df.copy()

    if template_output.exists():
        template_output.unlink()
    fill_df.to_excel(template_output, sheet_name='其他', index=False, engine='openpyxl')
    print(f"\n✅ 回填结果已保存: {template_output}")

    print(f"✅ 支付宝处理完成! 最终行数: {len(df)}")
    return {"原始行数": original_count, "最终行数": len(df), "填充行数": fill_count, "输出文件": str(output_file)}


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

    raw_dir = date_dir / "原始文件" / PLATFORM_NAME
    output_dir = date_dir / "生成文件" / PLATFORM_NAME

    if not raw_dir.is_dir():
        print(f"❌ 未找到原始文件目录: {raw_dir}")
        raise RuntimeError("脚本执行失败")

    has_files = any(raw_dir.glob("*.csv")) or any(raw_dir.glob("*.xlsx"))
    if not has_files:
        print(f"❌ 目录为空: {raw_dir}")
        raise RuntimeError("脚本执行失败")

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"📅 日期: {date_str}")
    print(f"🔄 开始处理支付宝消耗...")

    try:
        df = read_all_files(raw_dir)
        result = process_zhifubao(df, output_dir, date_str)
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"❌ 支付宝处理失败: {e}")
        log_dir = date_dir / "失败文件与日志"
        log_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = log_dir / f"支付宝消耗处理日志_{timestamp}.txt"
        lines = [f"支付宝消耗处理日志 - {date_str}", f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "=" * 60, "",
                 f"❌ 处理失败", f"  原因: {e}", "", "详细堆栈:", error_msg]
        log_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"⚠️  失败日志已保存: {log_path}")
        raise RuntimeError("脚本执行失败")


if __name__ == "__main__":
    main()
