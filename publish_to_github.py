# -*- coding: utf-8 -*-
"""
GitHub 一键发布工具 (基于 GitHub REST API)
无需安装 Git 客户端，直接一键安全发布到 GitHub 远程仓库
特点：严格仅上传当前 readexcel 文件夹内的代码与文档，绝不触碰任何上层图片或个人数据
"""

import os
import sys
import base64
import json
import urllib.request
import urllib.error

# 发布的文件清单 (严格白名单，绝不夹带任何隐私或多余文件)
PUBLISH_FILES = [
    ".gitignore",
    "LICENSE",
    "README.md",
    "requirements.txt",
    "excel_reader.py",
    "config.json",
    "使用说明.md",
    "启动朗读助手.bat",
    "assets/ui_overview.png"
]

def make_request(url, method="GET", data=None, token=""):
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "Excel-Speech-Auditor-Uploader",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"

    req_data = json.dumps(data).encode("utf-8") if data else None
    req = urllib.request.Request(url, data=req_data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode("utf-8")
            return json.loads(body) if body else {}
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="ignore")
        return {"error": e.code, "message": err_msg}

def main():
    print("=" * 60)
    print("       Excel 智能朗读核对助手 - GitHub 一键发布工具")
    print("=" * 60)
    print("【安全声明】本脚本仅发布当前文件夹内的开源代码与文档，")
    print("            严格杜绝上传任何个人表格、图片或数据文件。\n")

    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 1. 获取 GitHub Token
    token = input("请输入您的 GitHub Personal Access Token (PAT): ").strip()
    if not token:
        print("❌ Token 不能为空！请提供具备 repo 权限的 GitHub Token。")
        return

    # 2. 验证用户身份
    print("\n🔍 正在验证 GitHub 账号凭据...")
    user_info = make_request("https://api.github.com/user", token=token)
    if "error" in user_info:
        print(f"❌ 认证失败 (状态码 {user_info['error']}): {user_info.get('message')}")
        return

    username = user_info.get("login")
    print(f"✅ 认证成功！当前登录用户: {username}")

    # 3. 确认仓库名称与公开性
    repo_name = input("\n请输入仓库名称 (默认: excel-speech-auditor): ").strip()
    if not repo_name:
        repo_name = "excel-speech-auditor"

    is_private_input = input("是否创建为私有仓库 (Private)? [y/N] (默认公开 Public): ").strip().lower()
    is_private = True if is_private_input in ["y", "yes"] else False

    # 4. 创建或检查仓库
    print(f"\n🚀 正在检查/创建远程仓库: {username}/{repo_name} ...")
    create_payload = {
        "name": repo_name,
        "description": "Excel 智能核对朗读助手：基于微软神经网络高拟人语音与 COM 精准定位的表格人工核验工具",
        "private": is_private,
        "auto_init": False
    }
    repo_res = make_request("https://api.github.com/user/repos", method="POST", data=create_payload, token=token)
    if "error" in repo_res and repo_res["error"] != 422: # 422 表示仓库已存在，可以继续推送
        print(f"❌ 创建仓库失败: {repo_res.get('message')}")
        return

    repo_url = f"https://github.com/{username}/{repo_name}"
    print(f"✅ 远程仓库已就绪: {repo_url}")

    # 5. 上传文件清单中的纯净代码
    print("\n📦 正在将脱敏开源代码上传至 main 分支...")
    for filename in PUBLISH_FILES:
        file_path = os.path.join(current_dir, filename)
        if not os.path.exists(file_path):
            print(f"⚠️ 跳过未找到的文件: {filename}")
            continue

        with open(file_path, "rb") as f:
            content_bytes = f.read()
        content_b64 = base64.b64encode(content_bytes).decode("utf-8")

        # 检查远程是否已存在该文件 (用于获取 sha)
        file_api_url = f"https://api.github.com/repos/{username}/{repo_name}/contents/{filename}"
        get_res = make_request(file_api_url, token=token)
        sha = get_res.get("sha") if "sha" in get_res else None

        put_payload = {
            "message": f"feat: add {filename}",
            "content": content_b64,
            "branch": "main"
        }
        if sha:
            put_payload["sha"] = sha

        put_res = make_request(file_api_url, method="PUT", data=put_payload, token=token)
        if "error" in put_res:
            print(f"❌ 上传 {filename} 失败: {put_res.get('message')}")
        else:
            print(f"   ✓ 已成功同步: {filename}")

    print("\n" + "=" * 60)
    print("🎉 恭喜！项目已成功发布至 GitHub！")
    print(f"🔗 访问地址: {repo_url}")
    print("=" * 60)

if __name__ == "__main__":
    main()
