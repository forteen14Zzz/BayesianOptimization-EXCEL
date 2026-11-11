"""最终验收:模拟别人 clone 下来用。

检查:
1. git 仓库里有所有必需文件
2. 没有硬编码路径
3. 没有敏感信息
4. 从零建项目 + 跑通一次 validate
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
failures = []
notes = []


def ok(msg):
    print(f"  ✅ {msg}")


def bad(msg):
    print(f"  ❌ {msg}")
    failures.append(msg)


def check_files():
    print("\n[1] 必需文件存在")
    required = [
        "README.md",
        "LICENSE",
        ".gitignore",
        "agent.md",
        "init_new_project.bat",
        "install/install.bat",
        "install/requirements.txt",
        "templates/project_template.xlsx",
        "portal/run.py",
        "portal/watcher.py",
        "portal/core/pipeline.py",
        "portal/core/excel_control.py",
        "portal/core/convergence.py",
        "portal/core/convergence_plot.py",
        "scripts/init_new_project.py",
        "scripts/bat_templates.py",
        "scripts/verify_install.py",
        "tests/test_pipeline.py",
        "docs/FAQ.md",
        "docs/PRODUCT_DESIGN.md",
        "docs/IT_装机手册.md",
        "docs/同事使用手册.md",
        "新手使用指南.md",
        "projects/Suzuki_2026Q4_demo/Suzuki_2026Q4_demo.xlsx",
        "projects/Arylation_BMS_demo/Arylation_BMS_demo.xlsx",
    ]
    for rel in required:
        p = ROOT / rel
        if p.exists():
            ok(rel)
        else:
            bad(f"缺 {rel}")

    # 演示项目要有 3 个 bat
    for demo in ["Suzuki_2026Q4_demo", "Arylation_BMS_demo"]:
        for bat in ["跑_BO_点这里.bat", "刷新看板.bat", "生成记录表头.bat"]:
            p = ROOT / "projects" / demo / bat
            if not p.exists():
                bad(f"缺 projects/{demo}/{bat}")


def check_no_hardcoded():
    print("\n[2] 无硬编码路径 / 敏感信息")
    patterns = {
        "本机用户名": ["lenovo", "C:\\\\Users\\\\lenovo"],
        "项目绝对路径": ["E:\\\\bofire", "/c/Users/lenovo"],
        "邮箱": ["@qq.com", "@163.com", "@gmail.com"],
        "密钥": ["api_key", "apikey", "password=", "secret_key"],
    }
    exts = {".py", ".md", ".bat", ".txt", ".json"}
    skip_dirs = {".git", ".pytest_cache", "__pycache__", ".cache"}
    # ⚠️ 排除本脚本 + 历史清理工具 — 它们自己就包含这些搜索关键词
    skip_names = {
        Path(__file__).name,
        "acceptance_check.py",
        "_scan_history.py",
        "scrub_history_content.py",
        "rebuild_history_without.py",
        "strip_sensitive_history.py",
    }
    for label, pats in patterns.items():
        hits = []
        for f in ROOT.rglob("*"):
            if not f.is_file() or f.suffix not in exts:
                continue
            if f.name in skip_names:
                continue
            if any(part in skip_dirs for part in f.parts):
                continue
            try:
                text = f.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                continue
            for pat in pats:
                if pat.lower() in text.lower():
                    hits.append(f"{f.relative_to(ROOT)}: {pat}")
        if hits:
            bad(f"{label} → {hits[:3]}")
        else:
            ok(f"无{label}")


def check_bat_no_arg():
    print("\n[3] 项目 bat 不带参数(表格软件会丢参数)")
    for demo in ["Suzuki_2026Q4_demo", "Arylation_BMS_demo"]:
        for bat in ["跑_BO_点这里.bat", "刷新看板.bat", "生成记录表头.bat"]:
            p = ROOT / "projects" / demo / bat
            if not p.exists():
                continue
            text = p.read_text(encoding="utf-8", errors="ignore")
            # bat 内容里不能有 "xxx.bat sync" 这种带参数的调用
            if ".bat sync" in text or ".bat refresh" in text:
                bad(f"{demo}/{bat} 含参数调用")
    ok("全部独立 bat,无参数")


def check_git_clean():
    print("\n[4] git 状态")
    r = subprocess.run(
        ["git", "status", "--porcelain"], cwd=ROOT, capture_output=True, text=True
    )
    if r.stdout.strip():
        bad(f"有未提交改动:\n{r.stdout[:300]}")
    else:
        ok("working tree 干净")

    r = subprocess.run(
        ["git", "ls-files"], cwd=ROOT, capture_output=True, text=True
    )
    n = len(r.stdout.strip().splitlines())
    ok(f"跟踪 {n} 个文件")

    # 不该跟踪的东西
    tracked = r.stdout
    for bad_pat in ["LOG.txt", ".cache/", "__pycache__", ".pytest_cache", ".bak"]:
        if bad_pat in tracked:
            bad(f"跟踪了不该跟踪的:{bad_pat}")
        else:
            ok(f"未跟踪 {bad_pat}")


def check_end_to_end():
    print("\n[5] 从零建项目跑通")
    name = "ACC_验收"
    proj = ROOT / "projects" / name
    if proj.exists():
        subprocess.run(["rm", "-rf", str(proj)], shell=True)

    r = subprocess.run(
        [sys.executable, "scripts/init_new_project.py", name],
        cwd=ROOT, capture_output=True, text=True, timeout=180,
    )
    if r.returncode != 0:
        bad(f"init_new_project 失败: {r.stderr[-300:]}")
        return
    ok("项目创建成功")

    xlsx = proj / f"{name}.xlsx"
    if not xlsx.exists():
        bad("xlsx 没生成")
        return
    ok("xlsx 生成")

    for bat in ["跑_BO_点这里.bat", "刷新看板.bat", "生成记录表头.bat"]:
        if (proj / bat).exists():
            ok(f"bat 生成:{bat}")
        else:
            bad(f"缺 bat:{bat}")

    # 跑一次 validate(需要 PYTHONPATH 指向项目根)
    env = dict(os.environ)
    env["PYTHONPATH"] = str(ROOT)
    r = subprocess.run(
        [sys.executable, "portal/run.py", "--project", f"projects/{name}", "--phase", "validate"],
        cwd=ROOT, capture_output=True, text=True, timeout=180, env=env,
    )
    if '"ok": true' in r.stdout:
        ok("validate 跑通")
    else:
        bad(f"validate 失败: {r.stdout[-300:]} {r.stderr[-300:]}")

    # 清理
    subprocess.run(["rm", "-rf", str(proj)], shell=True)
    ok("已清理验收项目")


if __name__ == "__main__":
    print("=" * 60)
    print("BoFire Lab Portal — 上传 GitHub 前验收")
    print("=" * 60)
    check_files()
    check_no_hardcoded()
    check_bat_no_arg()
    check_git_clean()
    check_end_to_end()

    print("\n" + "=" * 60)
    if failures:
        print(f"❌ {len(failures)} 项未通过:")
        for f in failures:
            print(f"   - {f}")
        sys.exit(1)
    print("✅ 全部通过 — 可以上传 GitHub")
