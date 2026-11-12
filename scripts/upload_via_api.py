"""scripts/upload_via_api.py — 用 GitHub REST API 上传(绕过 git clone 协议)。

⚠️ 为什么不用 `git push`:
   这台机器上 github.com:443 连不上(被墙/路由问题),
   但 api.github.com 和 codeload.github.com 能通。
   所以改用 GitHub REST API 的 Git Database API:
   blobs → tree → commit → update ref

用法:
    python scripts/upload_via_api.py --token ghp_xxxx
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
API = "https://api.github.com"


class GitHub:
    def __init__(self, token: str, repo: str = ""):
        self.token = token
        self.repo = repo
        self.user = ""

    def _req(self, method: str, path: str, body=None, raw: bool = False):
        url = f"{API}{path}"
        data = json.dumps(body).encode() if body is not None else None
        req = urllib.request.Request(url, data=data, method=method)
        req.add_header("Authorization", f"Bearer {self.token}")
        req.add_header("Accept", "application/vnd.github+json")
        req.add_header("X-GitHub-Api-Version", "2022-11-28")
        req.add_header("User-Agent", "bofire-upload")
        if data:
            req.add_header("Content-Type", "application/json")
        for attempt in range(4):
            try:
                with urllib.request.urlopen(req, timeout=60) as r:
                    payload = r.read()
                    if raw:
                        return r.status, payload
                    return r.status, json.loads(payload.decode("utf-8"))
            except urllib.error.HTTPError as e:
                payload = e.read().decode("utf-8", "ignore")
                if e.code == 422 and attempt < 3:
                    time.sleep(1.5)
                    continue
                return e.code, payload
            except Exception:
                if attempt < 3:
                    time.sleep(2)
                    continue
                return 0, "network error"

    def get(self, path: str):
        return self._req("GET", path)

    def post(self, path: str, body: dict):
        return self._req("POST", path, body)

    def patch(self, path: str, body: dict):
        return self._req("PATCH", path, body)

    def put(self, path: str, body: dict):
        return self._req("PUT", path, body)

    def create_blob(self, content: bytes) -> str:
        """上传一个 blob,返回 sha。"""
        enc = base64.b64encode(content).decode()
        status, data = self.post(
            f"/repos/{self.repo}/git/blobs", {"content": enc, "encoding": "base64"}
        )
        if status != 201:
            raise RuntimeError(f"create_blob 失败: {status} {data}")
        return data["sha"]


def git(*args: str) -> str:
    env = {**os.environ, "MSYS2_ARG_CONV_EXCL": "*", "MSYS_NO_PATHCONV": "1"}
    r = subprocess.run(
        ["git", *args], cwd=str(ROOT), env=env, capture_output=True, text=True, timeout=300
    )
    if r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} 失败: {r.stderr[-300:]}")
    return r.stdout


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--token", required=True)
    p.add_argument("--repo", default="BayesianOptimization-EXCEL")
    p.add_argument("--branch", default="main")
    args = p.parse_args()

    gh = GitHub(args.token.strip())

    # 1) 验证
    print("[1/5] 验证 token...")
    status, data = gh.get("/user")
    if status != 200:
        print(f"❌ token 无效: {data}")
        return 1
    gh.user = data["login"]
    gh.repo = f"{gh.user}/{args.repo}"
    print(f"   ✅ {gh.user}")

    # 2) 仓库信息
    print("[2/5] 检查仓库...")
    status, repo = gh.get(f"/repos/{gh.repo}")
    if status != 200:
        print(f"❌ 仓库不可访问: {repo}")
        return 1
    print(f"   ✅ {gh.repo} ({repo['visibility']})")

    # 3) 收集要上传的文件(git ls-files)
    print("[3/5] 收集文件...")
    files = [f for f in git("ls-files").splitlines() if f.strip()]
    print(f"   {len(files)} 个文件")

    # 4) 上传 blob
    print("[4/5] 上传文件(通过 API)...")
    tree_entries = []
    for i, rel in enumerate(files, 1):
        path = ROOT / rel
        if not path.exists():
            print(f"   ⚠️ 跳过(不存在): {rel}")
            continue
        blob = gh.create_blob(path.read_bytes())
        mode = "100755" if os.access(path, os.X_OK) else "100644"
        tree_entries.append({"path": rel, "mode": mode, "type": "blob", "sha": blob})
        if i % 10 == 0 or i == len(files):
            print(f"   {i}/{len(files)}")

    # 5) 建 tree → commit → 更新 ref
    print("[5/5] 提交到远端...")
    # 现有 ref(如果是空仓库会 404)
    status, ref = gh.get(f"/repos/{gh.repo}/git/ref/heads/{args.branch}")
    parent = None
    if status == 200:
        parent = ref["object"]["sha"]
        print(f"   已有分支 {args.branch} @ {parent[:8]}")
    else:
        print(f"   新建分支 {args.branch}")

    status, tree = gh.post(
        f"/repos/{gh.repo}/git/trees",
        {"base_tree": parent, "tree": tree_entries} if parent else {"tree": tree_entries},
    )
    if status != 201:
        print(f"❌ create_tree 失败: {status} {tree}")
        return 1
    tree_sha = tree["sha"]

    # commit message
    msg = "上传 BoFire Lab Portal — Excel/WPS 无代码贝叶斯优化工具\n\n- bofire 0.5.0 + botorch,官方 Convergence Criteria\n- WPS/Excel COM 自动保存/关闭/重开\n- 73 个测试,验收脚本 5/5\n"
    body = {"message": msg, "tree": tree_sha}
    if parent:
        body["parents"] = [parent]
    status, commit = gh.post(f"/repos/{gh.repo}/git/commits", body)
    if status != 201:
        print(f"❌ create_commit 失败: {status} {commit}")
        return 1
    commit_sha = commit["sha"]

    if parent:
        status, data = gh.patch(
            f"/repos/{gh.repo}/git/refs/heads/{args.branch}", {"sha": commit_sha, "force": True}
        )
    else:
        status, data = gh.post(
            f"/repos/{gh.repo}/git/refs", {"ref": f"refs/heads/{args.branch}", "sha": commit_sha}
        )
    if status not in (200, 201):
        print(f"❌ 更新 ref 失败: {status} {data}")
        return 1

    print()
    print("=" * 62)
    print(f"🎉 上传完成!")
    print(f"   https://github.com/{gh.repo}")
    print("=" * 62)
    return 0


if __name__ == "__main__":
    sys.exit(main())
