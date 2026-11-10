"""scripts/verify_install.py — M2 验证:模拟 IT 装机流程。

不在真创建 venv(避免重复装包 5 分钟),而是用一个"虚拟 venv"
—— 直接在当前 Python 装 requirements.txt 中的"轻量"子集,验证:

1. 模板生成 → ✓(scripts/build_project_template.py)
2. Excel IO + pytest 全通过 → ✓(调用 pytest)
3. 端到端 demo 跑 3 轮 → ✓(scripts/run_demo.py)
4. install.bat / setup_autostart.bat 的 syntax check
5. start_watcher.bat 内容检查

输出 M2 验收报告。
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LIGHTWEIGHT_DEPS = [
    "openpyxl>=3.1",
    "pydantic>=2.5",
    "watchdog>=3.0",
    # 不装 bofire/torch/botorch — 那 125MB 下载太慢
]


def section(title: str):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print("=" * 70)


def ok(msg: str):
    print(f"  ✅ {msg}")


def fail(msg: str):
    print(f"  ❌ {msg}")
    sys.exit(1)


def step(name: str, fn):
    """每个 step:打印名 + 跑 + 标记 OK/FAIL。"""
    print(f"\n--- {name} ---")
    try:
        result = fn()
        if result is False:
            fail(f"{name} 返回 False")
        ok(name)
    except Exception as e:
        fail(f"{name}: {e}")


def test_template_gen():
    """模板生成。"""
    template = ROOT / "templates" / "project_template.xlsx"
    if template.exists():
        return True
    subprocess.run(["python", str(ROOT / "scripts" / "build_project_template.py")], check=True)
    return template.exists()


def test_pytest():
    """全部测试。"""
    rc = subprocess.run(
        ["python", "-m", "pytest", str(ROOT / "tests"), "-q"],
        cwd=str(ROOT),
    ).returncode
    return rc == 0


def test_demo():
    """3 轮 end-to-end demo。"""
    rc = subprocess.run(
        ["python", str(ROOT / "scripts" / "run_demo.py")],
        cwd=str(ROOT),
    ).returncode
    return rc == 0


def test_bat_syntax():
    """验证 install.bat 的 cmd 语法(干跑 echo-only)。"""
    # 在临时目录复制 bat 跑 dry-run — 但 cmd 不支持真 dry-run
    # 用 Python 模拟关键 check
    # 1. install.bat 必须有 @echo off
    bat = (ROOT / "install" / "install.bat").read_text(encoding="utf-8")
    if "@echo off" not in bat:
        fail("install.bat 缺 @echo off")
    if "exit /b 1" not in bat:
        fail("install.bat 缺 error 退出码")
    # 2. setup_autostart.bat 必须有 powershell
    autostart = (ROOT / "install" / "setup_autostart.bat").read_text(encoding="utf-8")
    if "powershell" not in autostart.lower():
        fail("setup_autostart.bat 缺 PowerShell 块")
    if "CreateShortcut" not in autostart:
        fail("setup_autostart.bat 缺 CreateShortcut")
    return True


def test_start_watcher_content():
    """start_watcher.bat 应该自动生成 — 验证生成逻辑存在。"""
    autostart = (ROOT / "install" / "setup_autostart.bat").read_text(encoding="utf-8")
    # 验证: 如果 start_watcher.bat 不存在,setup_autostart.bat 应创建它
    if "start_watcher.bat" not in autostart:
        fail("setup_autostart.bat 应能创建 start_watcher.bat")
    if "portal\\watcher.py" not in autostart:
        fail("setup_autostart.bat 应调 portal/watcher.py")
    return True


def test_requirements_locked():
    """requirements.txt 锁版本。"""
    reqs = (ROOT / "install" / "requirements.txt").read_text(encoding="utf-8")
    if "bofire[optimization]==0.5.0" not in reqs:
        fail("requirements.txt 未锁 bofire[optimization]==0.5.0")
    if "botorch" not in reqs:
        fail("requirements.txt 未列 botorch")
    # ⚠️ numpy<2.5.0 是 bofire 0.5.0 的硬要求,不能漏
    if "numpy" not in reqs or "<2.5.0" not in reqs:
        fail("requirements.txt 未约束 numpy<2.5.0(bofire 0.5.0 硬要求)")
    return True


def main():
    print("\n" + "=" * 70)
    print("  M2 验收:模拟 IT 装机流程")
    print("=" * 70)

    section("Step 1:依赖检查(轻量子集)")
    for d in LIGHTWEIGHT_DEPS:
        try:
            __import__(d.split(">=")[0].split("==")[0])
            print(f"  ✓ {d}")
        except ImportError:
            print(f"  ⚠️ {d} 未装(可能跳过)")

    section("Step 2:逐项验收")
    step("模板生成", test_template_gen)
    step("pytest 全通过", test_pytest)
    step("3 轮 end-to-end demo", test_demo)
    step("install.bat / setup_autostart.bat 语法", test_bat_syntax)
    step("start_watcher.bat 生成逻辑", test_start_watcher_content)
    step("requirements.txt 锁版本", test_requirements_locked)

    section("M2 验收通过 ✅")
    print("""
下一步建议:
  1. 把整个 <REPO>\\ 目录打包给 IT
  2. IT 在新机器:
     - 解压 → D:\\BOFIRELabPortal\\
     - 双击 install\\install.bat(等 5-10 分钟)
     - 选 y 配开机启动
  3. 验证:打开 projects\\Suzuki_2026Q4_demo\\Suzuki_2026Q4_demo.xlsx
     应该默认落在 🚀 开始 sheet, 看到 5 条初始数据
""")


if __name__ == "__main__":
    main()