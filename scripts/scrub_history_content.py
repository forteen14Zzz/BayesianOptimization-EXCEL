"""彻底清除 git 历史里的敏感**内容**(不只是删文件)。

问题:删文件不够 — 老版本的 `init_new_project.bat` 里写死了
本机用户名路径,这些内容留在历史里。

方案:fast-export → 逐个 blob 替换敏感字符串 → fast-import。
blob 长度会变,所以要**重算 data N**。

替换规则(保守,只动明显是本机路径/邮箱的):
    C:/Users/<用户名>/...  →  %USERPROFILE%/...
    E:/bofire/...          →  <REPO>/...
    <用户名>               →  user
    各种邮箱               →  someone@example.com
"""
from __future__ import annotations

import os
import re
import subprocess
import shutil
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# (旧, 新) —— 顺序重要,长的先替换
REPLACEMENTS: list[tuple[bytes, bytes]] = [
    (rb"C:\\Users\\lenovo\\Miniforge3\\python.exe", rb"%LOCALAPPDATA%\Miniforge3\python.exe"),
    (rb"C:\\Users\\lenovo", rb"%USERPROFILE%"),
    (rb"C:/Users/lenovo", rb"%USERPROFILE%"),
    (rb"E:\\bofire\\", rb"<REPO>\\"),
    (rb"E:/bofire/", rb"<REPO>/"),
    (b"455249875@qq.com", b"someone@example.com"),
    (b"729379195@qq.com", b"someone@example.com"),
    (b"gzuhgwz@163.com", b"someone@example.com"),
    (b"dsm@agentmail.to", b"someone@example.com"),
    (b"lenovo", b"user"),
]


def git(args: list[str], binary: bool = False, **kw):
    env = dict(os.environ)
    env["MSYS2_ARG_CONV_EXCL"] = "*"
    env["MSYS_NO_PATHCONV"] = "1"
    if binary:
        kw.setdefault("encoding", None)
        kw.setdefault("errors", None)
    else:
        kw.setdefault("text", True)
        kw.setdefault("encoding", "utf-8")
        kw.setdefault("errors", "replace")
    return subprocess.run(["git", *args], cwd=str(ROOT), env=env, capture_output=True, **kw)


def scrub(data: bytes) -> tuple[bytes, int]:
    """替换所有敏感串,返回 (新数据, 替换次数)。"""
    count = 0
    for old, new in REPLACEMENTS:
        n = data.count(old)
        if n:
            data = data.replace(old, new)
            count += n
    return data, count


def main() -> int:
    print("重建历史(内容脱敏)...")

    # 备份
    bak = Path(tempfile.gettempdir()) / f"git_scrub_{int(time.time())}"
    shutil.copytree(ROOT / ".git", bak)
    print(f"✅ 备份 .git → {bak}")

    exp = git(
        ["fast-export", "--all", "--signed-tags=strip", "--tag-of-filtered-object=drop"],
        binary=True,
    )
    if exp.returncode != 0:
        print("❌ fast-export 失败:", exp.stderr[-300:])
        return 1

    data = exp.stdout
    print(f"  {len(data)} 字节")

    # 逐块扫描:blob / mark / data N / <内容>
    out = bytearray()
    i = 0
    n = len(data)
    changed = 0
    while i < n:
        nl = data.find(b"\n", i)
        if nl == -1:
            out += data[i:]
            break
        line = data[i : nl + 1]
        stripped = line.rstrip(b"\n")

        if stripped == b"blob":
            # blob / mark :N / data SIZE / <SIZE 字节>
            out += line
            i = nl + 1

            # mark 行
            nl2 = data.find(b"\n", i)
            out += data[i : nl2 + 1]
            i = nl2 + 1

            # data SIZE 行
            nl3 = data.find(b"\n", i)
            size = int(data[i:nl3].split()[1])
            body = data[nl3 + 1 : nl3 + 1 + size]
            new_body, cnt = scrub(body)
            if cnt:
                changed += 1
                # ⚠️⚠️ 关键:替换后长度变了,**必须重写 data N**,
                #    否则 fast-import 会按旧长度读,流结构错位,
                #    报 "unsupported command: mit refs/heads/main"
                #    (把 blob 内容的后半截当成了命令)
                out += b"data " + str(len(new_body)).encode() + b"\n"
            else:
                out += data[i : nl3 + 1]
            out += new_body + b"\n"
            i = nl3 + 1 + size + 1
            continue

        out += line
        i = nl + 1

    print(f"  脱敏了 {changed} 个 blob")
    stream = bytes(out)

    # 去掉 reset 行
    stream = re.sub(rb"^reset refs/heads/[^\n]*\n", b"", stream, flags=re.M)

    # 重建
    git(["checkout", "--orphan", "rebuilt_main"])
    git(["reset", "--hard"])
    refs = git(["for-each-ref", "--format=%(refname)"]).stdout.split()
    for ref in refs:
        git(["update-ref", "-d", ref])
    git(["reflog", "expire", "--expire=now", "--all"])

    imp = subprocess.run(
        ["git", "fast-import", "--quiet", "--force"],
        cwd=str(ROOT),
        env={**os.environ, "MSYS2_ARG_CONV_EXCL": "*", "MSYS_NO_PATHCONV": "1"},
        input=stream,
        capture_output=True,
        timeout=1800,
    )
    if imp.returncode != 0:
        print("❌ fast-import 失败:", imp.stderr[-800:].decode("utf-8", "ignore"))
        print(f"   备份: {bak}")
        return 1

    git(["branch", "-m", "rebuilt_main", "main"])
    git(["checkout", "-f", "HEAD", "--", "."])
    git(["reflog", "expire", "--expire=now", "--all"])
    git(["gc", "--prune=now", "--aggressive"])

    print(f"\n✅ 完成。备份: {bak}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
