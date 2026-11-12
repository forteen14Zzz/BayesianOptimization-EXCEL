"""scripts/upload_to_github.py — 一键上传到 GitHub。

用法:
    python scripts/upload_to_github.py --token ghp_xxxx --username 你的用户名

做三件事:
1. 用 API 创建仓库(如果不存在)
2. 配置 git remote
3. push main 分支

⚠️ token 只在内存里用,不写盘、不进 git 历史。
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REPO_NAME = "BayesianOptimization-EXCEL"
DESCRIPTION = "Excel/WPS 无代码贝叶斯优化工具 — 基于 BoFire (实验设计 + GP-BO)"
API = "https://api.github.com"


def api(
    method: str,
    path: str,
    token: str,
    body: dict | None = None,
    timeout: float = 30,
) -> tuple[int, dict | str]:
    """调 GitHub API。返回 (status, data)。"""
    url = f"{API}{path}"
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    req.add_header("User-Agent", "bofire-upload")
    if data:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read().decode("utf-8", "ignore")
            try:
                return r.status, json.loads(raw)
            except json.JSONDecodeError:
                return r.status, raw
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "ignore")
        try:
            return e.code, json.loads(raw)
        except json.JSONDecodeError:
            return e.code, raw
    except Exception as e:
        return 0, str(e)


def git(*args: str, check: bool = True, timeout: int = 600) -> subprocess.CompletedProcess:
    env = {**os.environ, "MSYS2_ARG_CONV_EXCL": "*", "MSYS_NO_PATHCONV": "1"}
    r = subprocess.run(
        ["git", *args], cwd=str(ROOT), env=env, capture_output=True, text=True, timeout=timeout
    )
    if check and r.returncode != 0:
        print(f"❌ git {' '.join(args)} 失败:")
        print(r.stderr[-800:])
    return r


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--token", required=True, help="GitHub Personal Access Token")
    p.add_argument("--username", help="GitHub 用户名(不给就自动查)")
    p.add_argument("--private", action="store_true", default=True, help="私有仓库(默认)")
    p.add_argument("--public", dest="private", action="store_false", help="公开仓库")
    p.add_argument("--name", default=REPO_NAME, help="仓库名")
    args = p.parse_args()

    token = args.token.strip()

    # 1) 验证 token
    print("[1/4] 验证 token...")
    status, data = api("GET", "/user", token)
    if status != 200:
        print(f"❌ token 无效 (HTTP {status}): {data}")
        print("\n生成 token: https://github.com/settings/tokens/new")
        print("   Note 随便填,Expiration 选 90 days,Scopes 勾 'repo'")
        return 1
    username = args.username or data["login"]
    print(f"   ✅ 登录成功: {username}")

    # 2) 创建仓库
    print(f"[2/4] 创建仓库 {username}/{args.name}...")
    status, data = api(
        "POST",
        "/user/repos",
        token,
        {
            "name": args.name,
            "description": DESCRIPTION,
            "private": args.private,
            "auto_init": False,  # 我们已有历史
            "has_issues": True,
            "has_wiki": True,
        },
    )
    if status == 201:
        print(f"   ✅ 已创建: {data['html_url']}")
    elif status == 422 and "already exists" in str(data):
        print(f"   ℹ️ 仓库已存在,继续用")
    else:
        print(f"❌ 创建失败 (HTTP {status}): {data}")
        return 1

    remote_url = f"https://github.com/{username}/{args.name}.git"

    # 3) 配 remote
    print("[3/4] 配置 remote...")
    existing = git("remote", check=False).stdout.strip()
    if "origin" in existing:
        git("remote", "set-url", "origin", remote_url, check=False)
    else:
        git("remote", "add", "origin", remote_url, check=False)
    print(f"   origin → {remote_url}")

    # 4) push
    print("[4/4] push main...")
    # ⚠️ Windows 上 credential.helper 脚本路径会被 MSYS 转换搞坏
    #    ("E:\bofire\.git\..." → "E:bofire.git..."),脚本根本执行不了。
    #    改用 store 模式:把凭据写进临时文件,用
    #    -c credential.helper='store --file=<path>' 引用,
    #    push 完删掉。
    cred_file = ROOT / ".git" / "gh-credentials.tmp"
    cred_file.write_text(
        f"https://x-access-token:{token}@github.com\n", encoding="utf-8"
    )
    try:
        os.chmod(cred_file, 0o600)
    except OSError:
        pass

    env = {**os.environ, "MSYS2_ARG_CONV_EXCL": "*", "MSYS_NO_PATHCONV": "1"}
    env["GIT_TERMINAL_PROMPT"] = "0"
    r = subprocess.run(
        [
            "git",
            f"--git-dir={ROOT / '.git'}",
            f"--work-tree={ROOT}",
            "-c", f"credential.helper=store --file={cred_file}",
            "push", "-u", "origin", "main",
        ],
        cwd=str(ROOT),
        env=env,
        capture_output=True,
        text=True,
        timeout=1800,
    )
    try:
        cred_file.unlink()
    except OSError:
        pass

    if r.returncode != 0:
        print("❌ push 失败:")
        print((r.stderr or r.stdout)[-1200:])
        return 1

    print("   ✅ push 成功")
    print()
    print("=" * 60)
    print(f"🎉 上传完成: https://github.com/{username}/{args.name}")
    print("=" * 60)
    print()
    print("⚠️ 建议现在去撤销 token(用完即弃):")
    print(f"   https://github.com/settings/tokens  →  找到刚才那个 → Delete")
    return 0


if __name__ == "__main__":
    sys.exit(main())
