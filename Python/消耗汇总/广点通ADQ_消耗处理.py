"""
广点通ADQ - 客户财务日结报表处理脚本

规则：
1. 将"账户总消耗"填入"总支出"列，删除总支出为0的数据
2. 删除资金账户名称为"合约信用账户"的数据
3. 资金账户名称为"赠送账户"且资金类型为"赠送金"→资金类型改为"侠客行"
4. 资金账户名称为"分成收入快周转"/"IAP分成收入快周转"且资金类型为"现金"→资金类型改为"快周转"
5. 数据透视：行=账号ID，列=资金类型，值=总支出(元) 求和
6. 回填导入模板

用法:
  python 广点通ADQ_消耗处理.py             # 处理今天的日期目录
  python 广点通ADQ_消耗处理.py 2026-07-05  # 处理指定日期目录

目录结构：
  原始文件放在: {日期}/原始文件/广点通 ADQ- 消耗/
  生成文件输出: {日期}/生成文件/广点通 ADQ- 消耗/

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

PLATFORM_NAME = "广点通 ADQ- 消耗"

# 模板文件路径
TEMPLATE_PATH = SCRIPT_DIR / "回填数据导入模板_广点通.xlsx"


def process_adq_report(input_path: Path, output_path: Path, date_str: str) -> dict:
    """处理广点通ADQ客户财务日结报表"""

    # 读取数据
    df = pd.read_excel(input_path, dtype=str)
    original_count = len(df)
    print(f"📄 读取数据: {original_count} 行")

    # 转换数值列
    numeric_cols = ['共享钱包消耗(元)', '账户总消耗(元)', '总支出(元)']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # ========== 规则1: 处理总支出列 ==========
    # 共享钱包消耗≠0的行，原始总支出为0，需要用账户总消耗填入
    # 共享钱包消耗=0的行，保留原始总支出（本身已有值）
    # 简化：凡是总支出=0的行，用账户总消耗补充
    mask_zero_spend = df['总支出(元)'] == 0
    df.loc[mask_zero_spend, '总支出(元)'] = df.loc[mask_zero_spend, '账户总消耗(元)']
    print(f"规则1 - 总支出为0的行({mask_zero_spend.sum()}行)，用账户总消耗填入总支出: 完成")
    print(f"   总支出原本有值的行({(~mask_zero_spend).sum()}行)保持不变")

    # ========== 规则2: 删除总支出为0的数据 ==========
    before = len(df)
    df = df[df['总支出(元)'] != 0].copy()
    print(f"规则2 - 删除总支出为0: 删除 {before - len(df)} 行, 剩余 {len(df)} 行")

    # ========== 规则3: 删除资金账户名称为"合约信用账户"的数据 ==========
    before = len(df)
    df = df[df['资金账户名称'] != '合约信用账户'].copy()
    print(f"规则3 - 删除合约信用账户: 删除 {before - len(df)} 行, 剩余 {len(df)} 行")

    # ========== 规则4: 赠送账户+赠送金 → 资金类型改为侠客行 ==========
    mask5 = (df['资金账户名称'] == '赠送账户') & (df['资金类型'] == '赠送金')
    count5 = mask5.sum()
    df.loc[mask5, '资金类型'] = '侠客行'
    print(f"规则4 - 赠送金改为侠客行: 修改 {count5} 行")

    # ========== 规则5: 分成收入快周转/IAP分成收入快周转+现金 → 资金类型改为快周转 ==========
    mask6 = (df['资金账户名称'].isin(['分成收入快周转', 'IAP分成收入快周转'])) & (df['资金类型'] == '现金')
    count6 = mask6.sum()
    df.loc[mask6, '资金类型'] = '快周转'
    print(f"规则5 - 现金改为快周转: 修改 {count6} 行")

    # 保存处理后的基础文件（用于后续透视）
    if output_path.exists():
        output_path.unlink()
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"\n✅ 基础文件已保存: {output_path}")
    print(f"   最终行数: {len(df)}")

    # 输出资金类型分布
    print(f"\n📊 当前资金类型分布:")
    print(df['资金类型'].value_counts().to_string())

    # ========== 规则6: 数据透视 ==========
    pivot_df = pd.pivot_table(
        df,
        index='账号ID',
        columns='资金类型',
        values='总支出(元)',
        aggfunc='sum',
        fill_value=0
    )
    pivot_df = pivot_df.reset_index()
    pivot_df.columns.name = None  # 清除列名层级

    # 确保所有资金类型列都存在，不存在则填0
    all_types = ['现金', '信用金', '赠送金', '侠客行', '虚拟金', '快周转']
    for t in all_types:
        if t not in pivot_df.columns:
            pivot_df[t] = 0

    pivot_output = output_path.parent / f"广点通ADQ_透视结果.xlsx"
    if pivot_output.exists():
        pivot_output.unlink()
    pivot_df.to_excel(pivot_output, index=False, engine='openpyxl')
    print(f"\n✅ 透视文件已保存: {pivot_output}")
    print(f"   账号数: {len(pivot_df)}")
    print(f"   资金类型列: {[col for col in pivot_df.columns if col != '账号ID']}")

    # ========== 规则7: 回填模板 ==========
    yesterday = (datetime.strptime(date_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d") if date_str else ""
    template_output = output_path.parent / f"广点通ADQ_回填结果.xlsx"

    if TEMPLATE_PATH.is_file():
        # 读取模板表头（广点通模板在"广点通"sheet）
        template_header = pd.read_excel(TEMPLATE_PATH, sheet_name='广点通', dtype=str, nrows=0)
        template_cols = list(template_header.columns)

        # 计算总消耗（所有资金类型之和）
        type_cols = [col for col in pivot_df.columns if col != '账号ID']
        pivot_df['总消耗'] = pivot_df[type_cols].sum(axis=1)

        # 数据映射：模板列名 → 数据来源
        col_mapping = {
            '*日期': [yesterday] * len(pivot_df),
            '*投放账户ID': pivot_df['账号ID'].astype(str).values,
            '*总消耗': pivot_df['总消耗'].values,
            '现金消耗': pivot_df['现金'].values if '现金' in pivot_df.columns else 0,
            '信用金消耗': pivot_df['信用金'].values if '信用金' in pivot_df.columns else 0,
            '赠送金消耗(客户返点)': pivot_df['赠送金'].values if '赠送金' in pivot_df.columns else 0,
            '赠送金消耗(侠客行返点)': pivot_df['侠客行'].values if '侠客行' in pivot_df.columns else 0,
            '补偿虚拟金消耗': pivot_df['虚拟金'].values if '虚拟金' in pivot_df.columns else 0,
            '快周转消耗': pivot_df['快周转'].values if '快周转' in pivot_df.columns else 0,
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
        fill_df.to_excel(template_output, sheet_name='广点通', index=False, engine='openpyxl')
        print(f"✅ 回填模板已保存: {template_output}")
    else:
        print(f"⚠️  未找到模板文件: {TEMPLATE_PATH}，跳过回填")

    return {
        "原始行数": original_count,
        "最终行数": len(df),
        "基础文件": str(output_path),
        "透视文件": str(pivot_output),
        "回填文件": str(template_output) if TEMPLATE_PATH.is_file() else None,
        "账号数": len(pivot_df),
    }


def get_target_date_str():
    """获取目标日期字符串，支持命令行传参，默认今天"""
    if len(sys.argv) > 1:
        return sys.argv[1]
    return datetime.now().strftime("%Y-%m-%d")


def main():
    date_str = get_target_date_str()
    date_dir = PROJECT_ROOT / date_str

    if not date_dir.is_dir():
        print(f"❌ 未找到日期目录: {date_dir}")
        print("   请先运行 创建日期目录.py 创建目录结构，再放入原始文件。")
        sys.exit(1)

    raw_dir = date_dir / "原始文件" / PLATFORM_NAME
    output_dir = date_dir / "生成文件" / PLATFORM_NAME

    if not raw_dir.is_dir():
        print(f"❌ 未找到原始文件目录: {raw_dir}")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    # 找到原始文件（支持多个xlsx）
    xlsx_files = [f for f in raw_dir.glob("*.xlsx") if not f.name.startswith('~$') and not f.name.startswith('.~')]
    if not xlsx_files:
        print(f"❌ 在 {raw_dir} 下未找到xlsx文件")
        sys.exit(1)

    # 如果有多个文件，合并读取
    if len(xlsx_files) == 1:
        input_file = xlsx_files[0]
    else:
        print(f"📂 发现 {len(xlsx_files)} 个文件，合并读取...")
        all_dfs = []
        for f in xlsx_files:
            df = pd.read_excel(f, dtype=str)
            all_dfs.append(df)
            print(f"  📄 {f.name}: {len(df)} 行")
        merged_df = pd.concat(all_dfs, ignore_index=True)
        # 保存合并文件作为输入
        input_file = output_dir / "合并原始数据.xlsx"
        merged_df.to_excel(input_file, index=False, engine='openpyxl')

    output_file = output_dir / f"广点通ADQ_处理后基础文件_{date_str}.xlsx"

    print(f"🔄 开始处理广点通ADQ消耗数据...")
    print(f"   日期: {date_str}")
    print(f"   输入: {input_file}")
    print()

    try:
        result = process_adq_report(input_file, output_file, date_str)
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"❌ 广点通ADQ处理失败: {e}")
        log_dir = date_dir / "失败文件与日志"
        log_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = log_dir / f"广点通ADQ消耗处理日志_{timestamp}.txt"
        lines = [f"广点通ADQ消耗处理日志 - {date_str}", f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "=" * 60, "",
                 f"❌ 处理失败", f"  原因: {e}", "", "详细堆栈:", error_msg]
        log_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"⚠️  失败日志已保存: {log_path}")
        sys.exit(1)


if __name__ == "__main__":
    main()
