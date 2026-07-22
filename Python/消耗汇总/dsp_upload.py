"""
DSP系统自动上传脚本

功能：
  1. 登录DSP系统，获取SESSION cookie
  2. 查询媒体列表
  3. 后续：上传回填文件

依赖: 无额外依赖（使用curl）
"""

import json
import subprocess
import tempfile
from pathlib import Path


# DSP系统配置
DSP_BASE_URL = "https://dsp.nmgtime.com"
DSP_LOGIN_URL = f"{DSP_BASE_URL}/api/login/doLoginForVue"
DSP_SEARCH_MEDIA_URL = f"{DSP_BASE_URL}/api/out/outside/common/searchMediaList"

# 登录凭证
LOGIN_DATA = {
    "userName": "sunpengliang@nmgtime.com",
    "loginPwd": "013012",
    "id": "000"
}

# 通用请求头
COMMON_HEADERS = [
    "-H", "accept: application/json, text/plain, */*",
    "-H", "content-type: application/json; charset=UTF-8",
    "-H", "origin: https://dsp.nmgtime.com",
    "-H", "user-agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
]

# cookie文件路径（临时）
COOKIE_FILE = Path(tempfile.gettempdir()) / "dsp_cookies.txt"


def login() -> bool:
    """登录DSP系统，保存cookie到文件"""
    cmd = [
        "curl", "-s", "-k",
        DSP_LOGIN_URL,
        *COMMON_HEADERS,
        "-c", str(COOKIE_FILE),  # 保存cookie到文件
        "--data-raw", json.dumps(LOGIN_DATA),
    ]

    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        print(f"❌ curl执行失败: {proc.stderr}")
        return False

    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        print(f"❌ 响应解析失败: {proc.stdout[:200]}")
        return False

    if result.get("status") == "success":
        print(f"✅ 登录成功 - {result.get('loginUserInfo', {}).get('realName', '')}")
        return True
    else:
        print(f"❌ 登录失败: {result.get('message', '未知错误')}")
        return False


def search_media_list(media_type: str = "ks", is_overseas: str = "") -> list:
    """查询媒体列表
    
    Args:
        media_type: 媒体类型，如 ks(快手), tt(头条), gdt(广点通) 等
        is_overseas: 是否海外，空字符串表示不限
    
    Returns:
        媒体列表
    """
    data = {"mediaType": media_type, "isOverseas": is_overseas}

    cmd = [
        "curl", "-s", "-k",
        DSP_SEARCH_MEDIA_URL,
        *COMMON_HEADERS,
        "-b", str(COOKIE_FILE),  # 使用保存的cookie
        "--data-raw", json.dumps(data),
    ]

    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        print(f"❌ curl执行失败: {proc.stderr}")
        return []

    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        print(f"❌ 响应解析失败: {proc.stdout[:200]}")
        return []

    if result.get("status") == "success":
        media_list = result.get("listData", result.get("data", []))
        print(f"✅ 查询媒体列表成功 (mediaType={media_type}): {len(media_list)} 条")
        return media_list
    else:
        print(f"❌ 查询失败: {result}")
        return []


if __name__ == "__main__":
    # 1. 登录
    if not login():
        exit(1)

    # 2. 查询快手媒体列表
    print()
    media_list = search_media_list("ks")
    if media_list:
        print(f"返回数据示例: {json.dumps(media_list[:2], ensure_ascii=False, indent=2)}")
