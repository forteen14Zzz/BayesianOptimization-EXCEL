"""scripts/upgrade_project_template.py — 把老项目升级到新模板结构。

老模板 row 6-8 是本轮推荐(3 行),历轮推荐 row 48。
新模板 row 6-15 是本轮推荐(10 行),历轮推荐 row 53。

升级步骤:
  1. 把 ⚙️ 配置 / 🧪 记录 / _status sheet 内容保留
  2. 用新模板的 📊 看板 / 🚀 开始 替换
  3. 把缓存的最新推荐条件合并到新看板
"""
from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent


def upgrade(project_dir: Path) -> None:
    """升级项目的 xlsx 用新模板结构。"""
    xlsx = next(project_dir.glob("*.xlsx"))
    template = ROOT / "templates" / "project_template.xlsx"

    # 备份
    backup = xlsx.with_suffix(".xlsx.bak")
    shutil.copy2(xlsx, backup)
    print(f"备份: {backup}")

    # 读旧 xlsx 数据
    wb_old = load_workbook(str(xlsx))
    old_data = {}
    for sheet_name in ["⚙️ 配置", "🧪 记录", "_status"]:
        if sheet_name in wb_old.sheetnames:
            ws = wb_old[sheet_name]
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append(list(row))
            old_data[sheet_name] = rows
    wb_old.close()

    # 用新模板替换
    shutil.copy2(template, xlsx)

    # 写回旧数据
    wb = load_workbook(str(xlsx))
    for sheet_name, rows in old_data.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for r_idx, row in enumerate(rows, start=1):
            for c_idx, v in enumerate(row, start=1):
                if v is None:
                    continue
                ws.cell(row=r_idx, column=c_idx, value=v)
    wb.save(str(xlsx))
    wb.close()

    print(f"✅ 升级完成: {xlsx}")
    print(f"   保留: ⚙️ 配置 / 🧪 记录 / _status")
    print(f"   替换: 🚀 开始 / 📊 看板(用新模板)")

    # 重新触发一次 refresh 把 cache 推荐条件写到新看板
    from portal.core import pipeline, recommender
    cache = recommender.find_latest_cache(project_dir)
    if cache:
        print(f"   重新合并 cache: {cache.name}")
        result = pipeline.process(project_dir, pipeline.PHASE_REFRESH)
        print(f"   refresh: {result}")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--project", type=Path, required=True)
    args = p.parse_args()
    upgrade(args.project)


if __name__ == "__main__":
    main()