"""scripts/run_demo.py — M1 演示:跑 3 轮 BO,生成 PNG 看推荐变化。

不在 BoFire 不可用时也跑(走降级路径),验证整个项目从架构到演示
都通。
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
PROJECT_DIR = ROOT / "projects" / "Suzuki_2026Q4_demo"


def run(cmd: list[str], **kw) -> int:
    print(f"$ {' '.join(cmd)}")
    # 把 ROOT 加到 PYTHONPATH 让 subprocess 能找到 portal 包
    env = kw.pop("env", None) or __import__("os").environ.copy()
    env["PYTHONPATH"] = str(ROOT) + __import__("os").pathsep + env.get("PYTHONPATH", "")
    return subprocess.run(cmd, env=env, **kw).returncode


def fresh_demo():
    """重置 demo 项目到初始状态。"""
    if (PROJECT_DIR / ".cache").exists():
        shutil.rmtree(PROJECT_DIR / ".cache")
    # 把 demo xlsx 重新生成
    run(["python", str(ROOT / "scripts" / "build_demo_project.py")], check=True)


def simulate_round(round_num: int):
    """模拟同事:从 📊 看板 复制推荐条件到 🧪 记录,填随机 Yield。"""
    import random
    xlsx = next(PROJECT_DIR.glob("*.xlsx"))
    wb = load_workbook(str(xlsx))
    ws = wb["📊 看板"]

    # 读 row 6-8 的推荐
    candidates = []
    for r in range(6, 9):
        tid = ws.cell(row=r, column=1).value
        if tid is None:
            continue
        temp = ws.cell(row=r, column=2).value
        cat_amt = ws.cell(row=r, column=3).value
        time_h = ws.cell(row=r, column=4).value
        solvent = ws.cell(row=r, column=5).value
        candidates.append({
            "trial_id": tid, "温度": temp, "催化剂用量": cat_amt,
            "反应时间": time_h, "溶剂": solvent,
        })

    # 写回 🧪 记录(找到第一个空行)
    ws_rec = wb["🧪 记录"]
    next_row = 5
    while ws_rec.cell(row=next_row, column=1).value and not str(ws_rec.cell(row=next_row, column=1).value).startswith(("💡", "📌")):
        next_row += 1
    # 跳过提示行
    while next_row <= ws_rec.max_row:
        v = ws_rec.cell(row=next_row, column=1).value
        if v is None:
            break
        next_row += 1

    for c in candidates:
        # 模拟:Yield 跟温度有相关性 — 75 度最佳
        base = 80 - abs(c["温度"] - 75) * 0.5
        yield_v = base + random.uniform(-3, 3)
        ws_rec.cell(row=next_row, column=1, value=c["trial_id"])
        ws_rec.cell(row=next_row, column=2, value=f"2026-09-{15+round_num}")
        ws_rec.cell(row=next_row, column=3, value=c["温度"])
        ws_rec.cell(row=next_row, column=4, value=c["催化剂用量"])
        ws_rec.cell(row=next_row, column=5, value=c["反应时间"])
        ws_rec.cell(row=next_row, column=6, value=c["溶剂"])
        ws_rec.cell(row=next_row, column=7, value=round(yield_v, 1))
        ws_rec.cell(row=next_row, column=8, value=1)
        ws_rec.cell(row=next_row, column=9, value="measured")
        ws_rec.cell(row=next_row, column=10, value=f"第{round_num}轮推荐")
        ws_rec.cell(row=next_row, column=11, value=f"2026-09-{15+round_num} 14:30")
        next_row += 1

    wb.save(str(xlsx))
    wb.close()


def main():
    fresh_demo()
    print("\n=== M1 演示:跑 3 轮 BO ===\n")

    for r in range(1, 4):
        print(f"\n--- 第 {r} 轮 ---")
        # 1. 跑 BO
        rc = run([
            "python", str(ROOT / "portal" / "run.py"),
            "--project", str(PROJECT_DIR),
            "--phase", "ask",
        ])
        if rc != 0:
            print(f"❌ BO 失败 rc={rc}")
            sys.exit(1)
        # 2. 同步到 📊 看板
        rc = run([
            "python", str(ROOT / "portal" / "run.py"),
            "--project", str(PROJECT_DIR),
            "--phase", "refresh",
        ])
        # 3. 模拟同事填实验结果(下一轮)
        if r < 3:
            simulate_round(r)

    print("\n=== ✅ M1 演示完成 ===\n")
    print(f"📁 打开: {PROJECT_DIR / 'Suzuki_2026Q4_demo.xlsx'}")
    print("   看 📊 看板 sheet 看到 3 轮推荐历史")
    print(f"   看 LOG.txt 看到每次 BO 的运行日志")
    print(f"   看 .cache/ 看到每一轮的缓存文件")


if __name__ == "__main__":
    main()