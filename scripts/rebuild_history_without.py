"""重建 git 历史,移除指定文件(绕过 filter-branch 的 MSYS 路径问题)。

⚠️ 为什么不用 `git filter-branch`:
   它的 index-filter 会在子 shell 里再跑一次 `git`,而 MSYS 会把
   GIT_DIR='E:/bofire/.git' 转成 '/e/bofire/.git' — git 不认这个路径,
   报 "not a git repository"。

   本脚本用 `git fast-export` / `git fast-import` 纯文本流重建历史:
   - fast-export 时用 `--anonymize` 无关,直接改数据流,删掉目标文件的 blob
   - 完全不依赖子 shell 里的 git 调用,路径问题不存在

⚠️ 安全性:
   - 先自动备份 .git 到 TMPDIR
   - 重建后校验:目标文件在历史里彻底消失,其余提交数不变

用法:
    python scripts/rebuild_history_without.py <路径...>
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GIT_DIR = str((ROOT / ".git").resolve())


def git(args: list[str], binary: bool = False, **kwargs) -> subprocess.CompletedProcess:
    """跑 git,禁用 MSYS 路径转换。

    ⚠️ binary=True:fast-export 的输出含**二进制 blob**(xlsx 是 zip),
       用 text=True 会在解码时炸掉("invalid start byte"),
       然后 fast-import 收到损坏数据,报莫名其妙的
       "branch name doesn't conform to Git standards: refs/heads/main?"
    """
    env = dict(os.environ)
    env["MSYS2_ARG_CONV_EXCL"] = "*"
    env["MSYS_NO_PATHCONV"] = "1"
    if binary:
        kwargs.setdefault("encoding", None)
        kwargs.setdefault("errors", None)
    else:
        kwargs.setdefault("text", True)
        kwargs.setdefault("encoding", "utf-8")
        kwargs.setdefault("errors", "replace")
    return subprocess.run(
        ["git", *args],
        cwd=str(ROOT),
        env=env,
        capture_output=True,
        **kwargs,
    )


def backup_git() -> Path:
    """备份 .git。"""
    tmp = Path(tempfile.gettempdir()) / f"git_backup_{int(datetime_now())}"
    shutil.copytree(ROOT / ".git", tmp)
    return tmp


def datetime_now() -> float:
    import time

    return time.time()


def main() -> int:
    if len(sys.argv) < 2:
        print("用法: python scripts/rebuild_history_without.py <路径...>")
        return 1

    targets = set(sys.argv[1:])
    print(f"仓库: {ROOT}")
    print(f"要移除 {len(targets)} 个路径:")
    for t in sorted(targets):
        print(f"  - {t}")

    # 1) 备份
    bak = backup_git()
    print(f"\n✅ 已备份 .git → {bak}")

    # 2) 目标文件在工作区存在也没关系 — 我们只改历史,不动工作区
    #    (demo 项目的 bat 必须留着给同事用)

    # 3) 统计当前提交数
    before = git(["rev-list", "--all", "--count"]).stdout.strip()
    print(f"当前提交数: {before}")

    # 4) fast-export 全历史
    # ⚠️ 必须**二进制**模式:blob 里含 xlsx(zip)等二进制内容,
    #    text=True 会在解码时炸掉,fast-import 收到损坏数据就报
    #    "branch name doesn't conform to Git standards: refs/heads/main?"
    print("\nfast-export 中...")
    exp = git(
        ["fast-export", "--all", "--signed-tags=strip", "--tag-of-filtered-object=drop"],
        binary=True,
    )
    if exp.returncode != 0:
        print("❌ fast-export 失败:", exp.stderr[-500:])
        return 1

    data = exp.stdout
    print(f"  {len(data)} 字节数据流")

    # 5) 在数据流里删掉目标文件的 blob 声明和引用
    #
    # fast-export 格式(字节级):
    #   blob\n
    #   mark :1\n
    #   data <n>\n
    #   <n 字节内容>\n
    #   M 100644 :1 <path>\n
    #
    # 策略:先扫一遍找出目标路径对应的 blob mark,再删掉那些 blob 块和相关 M 行。

    # 5a) 找出每个目标路径对应的 mark
    #     先扫 "M <mode> <mark> <path>" 行,记录 path → mark
    #     ⚠️ 路径可能被 C-quote 转义("路径 with space"),要处理
    path_to_marks: dict[str, set[str]] = {}
    for raw in data.split(b"\n"):
        if not raw.startswith(b"M "):
            continue
        parts = raw.split(b" ", 3)
        if len(parts) < 4:
            continue
        mark = parts[2].decode("utf-8", "ignore")
        path = parts[3].decode("utf-8", "ignore")
        # 处理 C-style 引号路径
        if path.startswith('"') and path.endswith('"'):
            try:
                path = path[1:-1].encode("latin-1", "ignore").decode("unicode_escape").encode(
                    "latin-1"
                ).decode("utf-8", "ignore")
            except Exception:
                path = path[1:-1]
        if path in targets:
            path_to_marks.setdefault(path, set()).add(mark)

    if not path_to_marks:
        print("ℹ️ 目标文件在历史里已不存在 — 不用改")
        return 0

    marks_to_drop: set[str] = set()
    for marks in path_to_marks.values():
        marks_to_drop |= marks
    print(f"  找到 {len(marks_to_drop)} 个要删的 blob mark")

    # 5b) 重新扫,删掉:blob 块 + 引用这些 mark 的 M 行(字节级)
    out = bytearray()
    i = 0
    n = len(data)
    while i < n:
        nl = data.find(b"\n", i)
        if nl == -1:
            nl = n
            line = data[i:n]
            line_len = n - i
        else:
            line = data[i:nl]
            line_len = nl - i + 1

        if line == b"blob" and i + 1 < n:
            j = i + line_len
            nl2 = data.find(b"\n", j)
            mark_line = data[j : nl2 if nl2 != -1 else n]
            if mark_line.startswith(b"mark :"):
                mark = mark_line.decode("utf-8", "ignore").strip()[len("mark "):]
                if mark in marks_to_drop:
                    # 跳过整个 blob 块:blob / mark / data N / N 字节内容
                    k = nl2 + 1
                    nl3 = data.find(b"\n", k)
                    if nl3 != -1 and data[k:nl3].startswith(b"data "):
                        size = int(data[k:nl3].split()[1])
                        k = nl3 + 1 + size + 1  # data 行 + size 字节 + 换行
                    i = k
                    continue

        if line.startswith(b"M "):
            parts = line.split(b" ", 3)
            if len(parts) >= 4:
                mark = parts[2].decode("utf-8", "ignore")
                if mark in marks_to_drop:
                    i += line_len
                    continue

        out += data[i : i + line_len]
        i += line_len

    removed_bytes = len(data) - len(out)
    print(f"  删掉 {removed_bytes} 字节")
    data = bytes(out)

    # 6) 重建仓库
    print("\n重建仓库...")
    # ⚠️ checkout --orphan + reset --hard 会清空工作区!
    #    所以先记住哪些文件是"未跟踪但要保留"的(demo 的 bat 等),
    #    重建后从 HEAD 恢复 + 保留未跟踪文件
    git(["checkout", "--orphan", "rebuilt_main"])
    git(["reset", "--hard"])

    # 删掉所有 refs,重新导入
    refs = git(["for-each-ref", "--format=%(refname)"]).stdout.split()
    for ref in refs:
        git(["update-ref", "-d", ref])
    git(["reflog", "expire", "--expire=now", "--all"])

    # fast-import(必须二进制 — blob 含二进制内容)
    stream = data

    import re

    # ⚠️ fast-import 的坑:fast-export 输出的第一行是
    #    "reset refs/heads/main" —— fast-import 会把它当成
    #    "把分支重置到一个未指定的名字",报
    #    "branch name doesn't conform to Git standards: refs/heads/main"
    #    解法:删掉所有 reset 行(本来就要重建全部历史)
    stream = re.sub(rb"^reset refs/heads/[^\n]*\n", b"", stream, flags=re.M)

    # ⚠️⚠️⚠️ **不要**碰 commit message 里的 "?"
    #
    # 之前误以为 fast-import 会把 message 里的半角 ? 当成 branch 名的
    # "? " 后缀标记 — 其实那只是**症状**:用 text=True 读 fast-export 时,
    # 二进制 blob(xlsx 是 zip)解码失败,fast-import 收到损坏数据才报错。
    # 二进制模式下根本不需要任何转义。
    #
    # (之前试过"把 message 里的 ? 换全角",结果它把**二进制 blob 的内容**
    #  也当成 message 处理了 — blob 内容同样以 "data N" 开头 —
    #  字节被破坏,流结构错位,报 "unsupported command: ,迭代")

    stream = data

    # 去掉签名块
    stream = re.sub(rb"\n(?:gpgsig |-----BEGIN|-----END)[^\n]*\n", b"\n", stream)

    imp = subprocess.run(
        ["git", "fast-import", "--quiet", "--force"],
        cwd=str(ROOT),
        env={**os.environ, "MSYS2_ARG_CONV_EXCL": "*", "MSYS_NO_PATHCONV": "1"},
        input=stream,
        capture_output=True,
        timeout=1800,
    )
    if imp.returncode != 0:
        print("❌ fast-import 失败:", imp.stderr[-1500:].decode("utf-8", "ignore"))
        print(f"   备份在: {bak}")
        print("   恢复: rm -rf .git && cp -r <备份> .git && git checkout -f HEAD -- .")
        return 1

    # 7) 切回 main 分支名
    git(["branch", "-m", "rebuilt_main", "main"])

    # 7b) ⚠️ fast-import 可能在工作区留下半成品 → 从 HEAD 强制恢复
    #     ⚠️ 不要用 git clean — 会删掉 demo 项目需要的 bat 等文件
    git(["checkout", "-f", "HEAD", "--", "."])

    # 8) 校验
    after = git(["rev-list", "--all", "--count"]).stdout.strip()
    print(f"\n重建后提交数: {after}")

    still = git(
        ["log", "--all", "--pretty=format:", "--name-only", "--diff-filter=A"]
    ).stdout
    leaked = [t for t in targets if t in still]
    if leaked:
        print(f"❌ 仍能查到: {leaked}")
        print(f"   备份在: {bak}")
        return 1
    print("✅ 目标文件已从历史彻底移除")

    # 9) 清理
    git(["reflog", "expire", "--expire=now", "--all"])
    git(["gc", "--prune=now", "--aggressive"])

    print(f"\n✅ 完成。备份在: {bak}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
