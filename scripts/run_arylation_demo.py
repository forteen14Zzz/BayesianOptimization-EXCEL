"""scripts/run_arylation_demo.py — 真实 CSV + 真 GP-BO, 跑 5 轮 + 真 yield 反馈。

数据源: E:\\EvML-master\\EvML-master\\arylation_app\\All_CH_arylation_Experiments_Jay_12052019.csv
(Princeton Shields + BMS — 1728 条真实 Pd-催化直接芳基化实验)

每轮 ask 后,从 cache 拿推荐条件,去原始 CSV 找 base/ligand/solvent 完全匹配的
那条实验的 yield 作为"实测",写回主 xlsx 的 🧪 记录 sheet。然后 GP 看到新数据
继续 ask,真正迭代。

输出: 5 轮的 yield 收敛图,以及最高 yield 的实验组合。
"""
from __future__ import annotations

import csv
import os
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
PROJECT = ROOT / "projects" / "Arylation_BMS_demo"
CSV_PATH = Path(r"E:\EvML-master\EvML-master\arylation_app\All_CH_arylation_Experiments_Jay_12052019.csv")
N_ROUNDS = 5
BATCH_SIZE = 3


def load_csv() -> list[dict]:
    with open(CSV_PATH, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def run_ask() -> dict:
    """跑一轮 ask,返回 cache 信息。"""
    env = os.environ.copy()
    env["PYTHONPATH"] = "." + os.pathsep + env.get("PYTHONPATH", "")
    rc = subprocess.run(
        ["python", "portal/run.py", "--project", str(PROJECT), "--phase", "ask"],
        env=env, capture_output=True, text=True, cwd=str(ROOT),
    )
    import json
    if rc.returncode != 0:
        print(rc.stdout[-500:])
        print(rc.stderr[-500:])
        sys.exit(1)
    return json.loads(rc.stdout)


def run_refresh() -> None:
    """跑 refresh,把推荐条件写到 📊 看板。"""
    env = os.environ.copy()
    env["PYTHONPATH"] = "." + os.pathsep + env.get("PYTHONPATH", "")
    subprocess.run(
        ["python", "portal/run.py", "--project", str(PROJECT), "--phase", "refresh"],
        env=env, capture_output=True, text=True, cwd=str(ROOT),
    )


def find_match_in_csv(all_rows: list[dict], base: str, ligand: str, solvent: str) -> dict | None:
    """从原始 CSV 找 base+ligand+solvent 完全匹配的实验。"""
    matches = [r for r in all_rows if r["base"] == base and r["ligand"] == ligand and r["solvent"] == solvent]
    # 如果有多条匹配,取平均 yield(同一组合多次实验)
    if matches:
        avg_yield = sum(float(m["yield"]) for m in matches) / len(matches)
        return {
            "yield": avg_yield,
            "n_experiments": len(matches),
            "T": matches[0]["temperature"],
            "C": matches[0]["concentration"],
        }
    return None


def write_cache_to_records(all_rows: list[dict]) -> int:
    """把最新 cache 写到 🧪 记录 sheet, 用 CSV 的 yield 作为实测。返回新写行数。"""
    caches = sorted((PROJECT / ".cache").glob("recommend_*.xlsx"))
    if not caches:
        return 0
    latest = caches[-1]
    wb_r = load_workbook(latest, read_only=True)
    ws_r = wb_r.active
    cache_rows = [tuple(r) for r in ws_r.iter_rows(values_only=True)][1:]
    wb_r.close()

    wb = load_workbook(PROJECT / "Arylation_BMS_demo.xlsx")
    ws_rec = wb["🧪 记录"]
    # 找第一个空行
    next_row = 5
    while ws_rec.cell(row=next_row, column=1).value and str(ws_rec.cell(row=next_row, column=1).value).startswith("T"):
        next_row += 1

    n_written = 0
    for c in cache_rows:
        tid = c[0]
        base = c[3]
        ligand = c[4]
        solvent = c[5]
        T = float(c[6])
        C = float(c[7])
        match = find_match_in_csv(all_rows, base, ligand, solvent)
        if match is None:
            continue
        # 用 CSV 的 yield 当本轮"实测"
        ws_rec.cell(row=next_row, column=1, value=tid)
        ws_rec.cell(row=next_row, column=2, value="2026-09-30")
        ws_rec.cell(row=next_row, column=3, value=base)
        ws_rec.cell(row=next_row, column=4, value=ligand)
        ws_rec.cell(row=next_row, column=5, value=solvent)
        ws_rec.cell(row=next_row, column=6, value=C)
        ws_rec.cell(row=next_row, column=7, value=T)
        ws_rec.cell(row=next_row, column=8, value=match["yield"])
        ws_rec.cell(row=next_row, column=9, value=1)
        ws_rec.cell(row=next_row, column=10, value="measured")
        ws_rec.cell(row=next_row, column=11, value=f"BO 推荐(n={match['n_experiments']} 匹配)")
        ws_rec.cell(row=next_row, column=12, value="2026-09-30 14:30")
        next_row += 1
        n_written += 1
    wb.save(PROJECT / "Arylation_BMS_demo.xlsx")
    wb.close()
    return n_written


def reset_demo():
    """重置 demo 项目到初始 30 条数据状态。"""
    subprocess.run(["python", str(ROOT / "scripts" / "build_arylation_demo.py")], cwd=str(ROOT))


def print_round_summary(r: int, all_rows: list[dict]):
    """打印每轮推荐的结果。"""
    caches = sorted((PROJECT / ".cache").glob("recommend_*.xlsx"))
    if not caches:
        return
    latest = caches[-1]
    wb = load_workbook(latest, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))[1:]
    wb.close()
    print(f"\n第 {r} 轮推荐 ({os.path.basename(latest)}):")
    for c in rows:
        tid = c[0]
        base = c[3]
        ligand = c[4]
        solvent = c[5]
        T = float(c[6])
        C = float(c[7])
        match = find_match_in_csv(all_rows, base, ligand, solvent)
        if match:
            print(f"  {tid}: {base:8s} + {ligand:18s} + {solvent:10s} @ T={T:5.1f}°C C={C:.3f}M → yield={match['yield']:5.1f}% (n={match['n_experiments']})")
        else:
            print(f"  {tid}: {base} + {ligand} + {solvent} @ T={T:.1f}°C C={C:.3f}M → (无匹配)")


def print_final_stats(all_rows: list[dict]):
    """打印总览: 30 条种子 + N*3 条 BO 推荐的 yield 统计。"""
    wb = load_workbook(PROJECT / "Arylation_BMS_demo.xlsx", read_only=True, data_only=True)
    ws = wb["🧪 记录"]
    seed = []
    bo = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r and r[0] and str(r[0]).startswith("T"):
            y = float(r[7]) if r[7] is not None else 0
            source = str(r[10]) if r[10] else ""
            if "CSV" in source or "种子" in source:
                seed.append((r[0], y, r[2], r[3], r[4]))
            else:
                bo.append((r[0], y, r[2], r[3], r[4]))
    wb.close()

    print(f"\n{'='*70}")
    print(f"📊 总结")
    print(f"{'='*70}")
    print(f"CSV 种子: {len(seed)} 条")
    if seed:
        ys = [y for _, y, *_ in seed]
        print(f"  yield 范围 [{min(ys):.1f}, {max(ys):.1f}], 平均 {sum(ys)/len(ys):.2f}")
    print(f"BO 推荐: {len(bo)} 条")
    if bo:
        ys = [y for _, y, *_ in bo]
        print(f"  yield 范围 [{min(ys):.1f}, {max(ys):.1f}], 平均 {sum(ys)/len(ys):.2f}")
        if seed:
            seed_avg = sum(y for _, y, *_ in seed) / len(seed)
            bo_avg = sum(y for _, y, *_ in bo) / len(bo)
            print(f"\n  提升: BO 推荐 yield 平均比 CSV 种子 {'高' if bo_avg > seed_avg else '低'} {(bo_avg - seed_avg):.2f} 个百分点")

    print(f"\n🏅 最高 yield 实验:")
    all_data = seed + bo
    all_data.sort(key=lambda x: -x[1])
    for tid, y, base, ligand, solvent in all_data[:10]:
        tag = "BO" if tid in [t for t, *_ in bo] else "CSV"
        print(f"  {tid} [{tag}] yield={y:5.1f}%  {base}+{ligand}+{solvent}")


def main():
    print("=" * 70)
    print("Arylation Demo — 真实 1728 条反应 + 真 BoFire GP-BO")
    print("=" * 70)
    print("\n重置 demo 项目...")
    all_rows = load_csv()
    reset_demo()
    print(f"  ✓ {len(all_rows)} 条 CSV 实验可用")

    for r in range(1, N_ROUNDS + 1):
        print(f"\n--- 第 {r}/{N_ROUNDS} 轮 ---")
        result = run_ask()
        run_refresh()
        n = write_cache_to_records(all_rows)
        print(f"  ask: {result.get('elapsed_s', 0):.1f}s, 写入 {n} 条新数据到 🧪 记录")
        print_round_summary(r, all_rows)

    print_final_stats(all_rows)
    print(f"\n📁 打开: {PROJECT / 'Arylation_BMS_demo.xlsx'}")
    print(f"   看 📊 看板 sheet 看到 {N_ROUNDS} 轮推荐历史")


if __name__ == "__main__":
    main()