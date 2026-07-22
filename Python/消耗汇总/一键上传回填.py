"""
DSP系统 - 一键上传回填结果

功能：
  1. 登录DSP系统
  2. 遍历当天生成文件中所有回填结果文件
  3. 逐个上传到对应的媒体

用法:
  python 一键上传回填.py              # 上传今天的回填结果
  python 一键上传回填.py 2026-07-08   # 上传指定日期的回填结果

依赖: 无额外依赖（使用系统curl）
"""

import sys
import json
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime

# 脚本所在目录（Python/消耗汇总/）
SCRIPT_DIR = Path(__file__).parent
# 项目根目录（py 自动化脚本/）
PROJECT_ROOT = SCRIPT_DIR.parent.parent

# DSP系统配置（测试环境）
DSP_BASE_URL = "http://fz.slave1.dypxxm.com"
DSP_LOGIN_URL = f"{DSP_BASE_URL}/api/login/doLoginForVue"
DSP_UPLOAD_URL = f"{DSP_BASE_URL}/api/out/outside/placingAcc/importPlacingAccExpend"

# 登录凭证（测试环境）
LOGIN_DATA = {
    "userName": "sunpengliang@nmgtime.com",
    "loginPwd": "412846",
    "id": "000"
}

# cookie文件路径
COOKIE_FILE = Path(tempfile.gettempdir()) / "dsp_cookies.txt"

# 通用请求头
COMMON_HEADERS = [
    "-H", "accept: application/json, text/plain, */*",
    "-H", "content-type: application/json; charset=UTF-8",
    "-H", f"origin: {DSP_BASE_URL}",
    "-H", "user-agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
]

# ============================================================
# 平台 → 媒体映射配置
# key: 生成文件目录名
# value: (媒体名称, mediaId, mediaType)
# ============================================================
PLATFORM_MEDIA_MAP = {
    # 快手系列
    "快手-磁力- 消耗": ("快手-磁力", "6DCBF78511D8BD7DE050007F010034A6", "ks"),
    "快手-本地生活- 消耗": ("快手-本地生活", "127442cff98fcc2e5a941c90bf8830b9", "other"),
    "快手-金教- 消耗": ("快手-磁力（金教）", "f4266dcfde5b62ccfe5cc02ed15ccb9d", "ks"),
    # 金牛系列
    "金牛-华北- 消耗": ("快手-磁力金牛-华北", "C6FB25E1F7F42E94E050A8C06A1027A6", "ks"),
    "金牛-广分- 消耗": ("快手-磁力金牛-广分", "e8c6f40e1363ff680d1d440fcda0a858", "ks"),  # 测试环境无此媒体
    "金牛-澳比- 消耗": ("磁力-金牛（澳比）", "825edfc334dbf62b0e766fa723a89ea0", "other"),  # 测试环境无此媒体
    # 抖音/巨量系列
    "抖音-竞价- 消耗": ("抖音-巨量", "7B2AF195E8243606E05064ACFD154E37", "tt"),
    "抖音-品牌- 消耗": ("抖音-品牌", "A6AC801DCC0D1FA8E050A8C05410E266", "tt"),
    "抖音-本地推- 消耗": ("抖音-本地推", "E562EE47D50561A9E05017AC8EEC3B27", "tt"),
    "抖音-巨懂车- 消耗": ("抖音-巨懂车", "48e496469deb24fba51db82e354eea3b", "other"),
    # 千川系列
    "千川-竞价- 消耗": ("抖音-千川", "BF81A081F0283E53E050A8C06A100768", "tt"),
    "千川-品牌- 消耗": ("千川-品牌", "DD76D6AE5136B41EE05017AC8EEC222E", "tt"),
    "千川-巨懂车- 消耗": ("抖音-千川-巨懂车", "5c3587f653b08dcff17941fba756c029", "tt"),
    # 陵致长风系列
    "陵致长风-巨量- 消耗": ("抖音-巨量（陵致长风）", "98bc5a4a4d26e651d5f7220d27250b6c", "tt"),
    "陵致长风-千川- 消耗": ("抖音-千川（陵致长风）", "6d2c309f323b4c05372317f89ace4847", "tt"),
    # 广点通
    "广点通 ADQ- 消耗": ("广点通-ADQ", "7516F461BBA84C9EE05064ACFD153D74", "gdt"),
    # 支付宝
    "支付宝- 消耗": ("支付宝-信息流", "2e5db4cac0c5fcb38ac28eea20d57467", "other"),
    # B站
    "B站-信息流- 消耗": ("B站-信息流", "4b82b8b8c5405611645a3787ae5acd08", "other"),
    "B站-商业起飞- 消耗": ("B站-商业起飞", "f8986ed90a00d9078543b290184fc620", "other"),
}


def login() -> bool:
    """登录DSP系统"""
    cmd = [
        "curl", "-s", "-k",
        DSP_LOGIN_URL,
        *COMMON_HEADERS,
        "-c", str(COOKIE_FILE),
        "--data-raw", json.dumps(LOGIN_DATA),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        print(f"❌ 登录请求失败: {proc.stderr}")
        return False
    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        print(f"❌ 登录响应解析失败: {proc.stdout[:200]}")
        return False

    if result.get("status") == "success":
        name = result.get("loginUserInfo", {}).get("realName", "")
        print(f"✅ 登录成功 - {name}")
        return True
    else:
        print(f"❌ 登录失败: {result.get('message', '未知错误')}")
        return False


def upload_file(file_path: Path, media_id: str, media_type: str) -> dict:
    """上传回填文件到DSP"""
    cmd = [
        "curl", "-s", "-k",
        DSP_UPLOAD_URL,
        "-H", "accept: application/json, text/plain, */*",
        "-H", f"origin: {DSP_BASE_URL}",
        "-H", "user-agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "-b", str(COOKIE_FILE),
        "-F", f"mediaId={media_id}",
        "-F", f"file=@{file_path};type=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "-F", f"mediaType={media_type}",
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        return {"status": "error", "message": f"curl失败: {proc.stderr}"}
    try:
        return json.loads(proc.stdout)
    except json.JSONDecodeError:
        return {"status": "error", "message": f"响应解析失败: {proc.stdout[:200]}"}


def find_backfill_file(output_dir: Path):
    """在输出目录中找到回填结果文件"""
    patterns = ["*回填结果*.xlsx", "*回填*.xlsx"]
    for pattern in patterns:
        files = list(output_dir.glob(pattern))
        if files:
            return files[0]
    return None


def get_target_date_str():
    if len(sys.argv) > 1:
        return sys.argv[1]
    return datetime.now().strftime("%Y-%m-%d")


def main():
    date_str = get_target_date_str()
    date_dir = PROJECT_ROOT / date_str
    gen_dir = date_dir / "生成文件"

    if not gen_dir.is_dir():
        print(f"❌ 未找到生成文件目录: {gen_dir}")
        sys.exit(1)

    print(f"📅 日期: {date_str}")
    print(f"📂 目录: {gen_dir}")
    print()

    # 登录
    print("🔐 登录DSP系统...")
    if not login():
        sys.exit(1)
    print()

    # 遍历所有平台，找到回填文件并上传
    success_count = 0
    fail_count = 0
    failures = []  # 记录失败详情

    for folder_name, (media_name, media_id, media_type) in PLATFORM_MEDIA_MAP.items():
        output_dir = gen_dir / folder_name
        if not output_dir.is_dir():
            continue

        backfill_file = find_backfill_file(output_dir)
        if not backfill_file:
            continue

        print(f"📤 上传: {media_name}")
        print(f"   文件: {backfill_file.name}")

        result = upload_file(backfill_file, media_id, media_type)

        if result.get("status") == "success":
            success_count += 1
            msg = result.get("message", result.get("msg", ""))
            print(f"   ✅ 成功 {msg}")
        else:
            fail_count += 1
            msg = result.get("message", result.get("msg", str(result)))
            print(f"   ❌ 失败: {msg}")
            failures.append({
                "媒体": media_name,
                "文件": str(backfill_file),
                "错误": msg,
            })
        print()

    # 汇总
    print("=" * 50)
    print(f"📋 上传汇总: ✅ 成功 {success_count} | ❌ 失败 {fail_count}")

    # 写入失败日志
    if failures:
        log_dir = date_dir / "失败文件与日志"
        log_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = log_dir / f"上传回填失败日志_{timestamp}.txt"
        lines = [
            f"上传回填失败日志 - {date_str}",
            f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "=" * 60,
            f"成功: {success_count} | 失败: {fail_count}",
            "",
        ]
        for f in failures:
            lines.append(f"❌ {f['媒体']}")
            lines.append(f"   文件: {f['文件']}")
            lines.append(f"   错误: {f['错误']}")
            lines.append("")
        log_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n⚠️  失败日志已保存: {log_path}")


if __name__ == "__main__":
    main()
