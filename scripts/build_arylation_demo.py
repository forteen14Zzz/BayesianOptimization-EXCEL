"""scripts/build_arylation_demo.py — 把真实 CSV → BoFire Lab Portal demo 项目。

数据源: E:\\EvML-master\\EvML-master\\arylation_app\\All_CH_arylation_Experiments_Jay_12052019.csv
(Princeton Shields + BMS 反应优化游戏 — 1728 条真实 Pd-催化直接芳基化实验)

因子:
  - base: KOAc, KOPiv, CsOAc, CsOPiv (4)
  - solvent: BuCN, BuOAc, DMAc, p-Xylene (4)
  - ligand: 12 种
  - concentration: 0.057, 0.1, 0.153 (3) — 当 continuous
目标: yield (maximize, 0-100)

这个 demo:
  - 留 5 条初始数据(种子随机)
  - 用 BoFire 真 SOBO 跑 N 轮,验证架构 + GP-BO 在真数据上的合理性
"""
from __future__ import annotations

import csv
import random
import shutil
from pathlib import Path

from openpyxl import load_workbook

CSV_PATH = Path(r"E:\EvML-master\EvML-master\arylation_app\All_CH_arylation_Experiments_Jay_12052019.csv")
ROOT = Path(__file__).resolve().parent.parent
TARGET_DIR = ROOT / "projects" / "Arylation_BMS_demo"
TARGET_XLSX = TARGET_DIR / "Arylation_BMS_demo.xlsx"


def load_csv() -> list[dict[str, str]]:
    """读所有反应实验。"""
    with open(CSV_PATH, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def seed_records(rows: list[dict], n: int = 30) -> list[dict]:
    """选 yield 最高的 n 条 + 一些随机 — 模拟"早期化学家有直觉"。

    我们选 n=30 混合:
      - yield > 50 的前 20 条(让 GP 看到"好东西")
      - 10 条随机(让 GP 看到多样性)
    """
    random.seed(42)
    high_yield = sorted([r for r in rows if float(r["yield"]) > 50], key=lambda r: -float(r["yield"]))[:20]
    random_sample = random.sample([r for r in rows if float(r["yield"]) <= 50], 10)
    return high_yield + random_sample


def build_template() -> None:
    """复制 templates/project_template.xlsx 到目标。"""
    template = ROOT / "templates" / "project_template.xlsx"
    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    if TARGET_XLSX.exists():
        TARGET_XLSX.unlink()
    shutil.copy2(template, TARGET_XLSX)


def fill_config(wb) -> None:
    """写 ⚙️ 配置 sheet: 5 个因子 + Yield 目标。"""
    ws = wb["⚙️ 配置"]
    # Factors row 6 起(模板预填了 4 行,我们覆盖前 5 行)
    factors = [
        # key, type, trainable, low, high, categories, unit
        ("base",       "categorical", "TRUE", "", "", "KOAc;KOPiv;CsOAc;CsOPiv", ""),
        ("ligand",     "categorical", "TRUE", "", "", "BrettPhos;CgMe-PPh;GorlosPhos HBF4;JackiePhos;P(fur)3;PCy3 HBF4;PPh2Me;PPh3;PPhMe2;PPhtBu2;X-Phos;tBPh-CPhos", ""),
        ("solvent",    "categorical", "TRUE", "", "", "BuCN;BuOAc;DMAc;p-Xylene", ""),
        ("temperature","continuous",  "TRUE", 80, 130, "", "°C"),
        ("concentration", "continuous", "TRUE", 0.05, 0.2, "", "M"),
    ]
    for i, (k, t, tr, lo, hi, cats, un) in enumerate(factors):
        r = 6 + i
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=t)
        ws.cell(row=r, column=3, value=tr)
        if lo:
            ws.cell(row=r, column=4, value=float(lo))
        if hi:
            ws.cell(row=r, column=5, value=float(hi))
        if cats:
            ws.cell(row=r, column=6, value=cats)
        if un:
            ws.cell(row=r, column=7, value=un)

    # Outputs row 14 起 — Yield maximize
    ws.cell(row=14, column=1, value="yield")
    ws.cell(row=14, column=2, value="maximize")
    ws.cell(row=14, column=4, value=0)
    ws.cell(row=14, column=5, value=100)
    ws.cell(row=14, column=6, value="%")


def fill_records(wb, records: list[dict]) -> None:
    """写 🧪 记录 sheet,初始 N 条数据。

    模板 row 5-12 是输入区(8 行),我们需要更多行 — 解除合并 + 增到 N 行。
    """
    ws = wb["🧪 记录"]
    # 解除 row 5+ 的所有合并(模板可能合并了"💡 提示"行)
    merged_ranges = list(ws.merged_cells.ranges)
    for rng in merged_ranges:
        # 保留 row 4(表头)的合并,解除 row 5+
        if rng.min_row >= 5:
            ws.unmerge_cells(str(rng))

    headers = ["trial_id", "batch_id", "base", "ligand", "solvent", "concentration", "temperature", "yield", "valid", "status", "来源", "录入时间"]
    for j, h in enumerate(headers):
        ws.cell(row=4, column=j + 1, value=h)

    n = len(records)
    for i, rec in enumerate(records):
        r = 5 + i
        ws.cell(row=r, column=1, value=f"T{i+1:04d}")
        ws.cell(row=r, column=2, value="2026-09-29")
        ws.cell(row=r, column=3, value=rec["base"])
        ws.cell(row=r, column=4, value=rec["ligand"])
        ws.cell(row=r, column=5, value=rec["solvent"])
        ws.cell(row=r, column=6, value=float(rec["concentration"]))
        ws.cell(row=r, column=7, value=float(rec["temperature"]))
        ws.cell(row=r, column=8, value=float(rec["yield"]))
        ws.cell(row=r, column=9, value=1)
        ws.cell(row=r, column=10, value="measured")
        ws.cell(row=r, column=11, value="CSV 种子数据")
        ws.cell(row=r, column=12, value="2026-09-29 10:00")


def fill_status(wb) -> None:
    """写 _status sheet 的项目名。"""
    ws = wb["_status"]
    ws["B5"] = "Arylation_BMS_demo"
    ws["B6"] = str(TARGET_DIR)


def main():
    if not CSV_PATH.exists():
        print(f"❌ CSV 不存在: {CSV_PATH}")
        return
    print(f"📂 读 CSV: {CSV_PATH}")
    rows = load_csv()
    print(f"   {len(rows)} 条实验")

    print("🔧 构建 demo 项目模板...")
    build_template()
    wb = load_workbook(str(TARGET_XLSX))
    fill_config(wb)
    seed = seed_records(rows, n=8)
    fill_records(wb, seed)
    fill_status(wb)
    wb.save(str(TARGET_XLSX))
    wb.close()

    print(f"\n✅ Demo 项目生成: {TARGET_XLSX}")
    print(f"   初始数据 {len(seed)} 条,seed 固定 42")
    print(f"   因子: base / ligand / solvent / temperature / concentration")
    print(f"   目标: yield (maximize)")
    print()
    print("下一步:")
    print(f"  python portal/run.py --project {TARGET_DIR} --phase ask")
    print(f"  python portal/run.py --project {TARGET_DIR} --phase refresh")


if __name__ == "__main__":
    main()