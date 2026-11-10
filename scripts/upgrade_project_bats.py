"""scripts/upgrade_project_bats.py — 给现有项目升级成 3 个独立 bat。

⚠️ 为什么要拆:
表格软件(WPS/Excel)打开 .bat 时**参数会丢**(实测同事 222 项目
点 [📋 生成记录表头] 却跑了 ask)。所以每个按钮一个独立 bat:
  跑_BO_点这里.bat      — ask + refresh
  刷新看板.bat          — 只合并 cache 到看板
  生成记录表头.bat      — 只按配置重建 🧪 记录 表头

同时把 🚀 开始 sheet 的按钮改指向对应的独立 bat。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from bat_templates import write_bats  # noqa: E402


BUTTON_TARGETS = [
    # (row, col, 标签关键词, bat 文件名)
    (17, 1, "启动 BO", "跑_BO_点这里.bat"),
    (17, 4, "刷新", "刷新看板.bat"),
    (20, 1, "生成记录表头", "生成记录表头.bat"),
]


def upgrade(project_dir: Path) -> None:
    from openpyxl import load_workbook
    from openpyxl.worksheet.hyperlink import Hyperlink

    xlsx = next(project_dir.glob("*.xlsx"))
    print(f"项目: {project_dir.name}")

    # 1. 写 3 个 bat
    bats = write_bats(project_dir)
    for name, path in bats.items():
        print(f"  ✅ bat: {name}")

    # 2. 改按钮
    wb = load_workbook(str(xlsx))
    if "🚀 开始" not in wb.sheetnames:
        wb.close()
        print("  ⚠️ 没找到 🚀 开始 sheet")
        return
    ws = wb["🚀 开始"]

    for row, col, keyword, bat_name in BUTTON_TARGETS:
        cell = ws.cell(row=row, column=col)
        v = cell.value
        if not (isinstance(v, str) and "HYPERLINK" in v):
            # 位置不对就全表扫一遍找
            found = None
            for r in range(1, 30):
                for c in range(1, 9):
                    vv = ws.cell(row=r, column=c).value
                    if isinstance(vv, str) and keyword in vv and "HYPERLINK" in vv:
                        found = (r, c, vv)
                        break
                if found:
                    break
            if not found:
                # ⚠️ 老项目缺这个按钮 → 新建一个
                if keyword == "生成记录表头":
                    r = 20
                    while ws.cell(row=r, column=1).value is not None:
                        r += 1
                    found = (r, 1, "")
                    print(f"  ➕ 补建 [生成记录表头] 按钮 (row {r})")
                else:
                    print(f"  ⚠️ 找不到含 {keyword!r} 的按钮")
                    continue
            row, col, v = found
            cell = ws.cell(row=row, column=col)

        label = ""
        if '"' in v:
            parts = v.split('","')
            if len(parts) >= 2:
                label = parts[-1].rstrip('")')

        bat_abs = str((project_dir / bat_name).resolve())
        new_label = {
            "启动 BO": "🚀 启动 BO",
            "刷新": "🔄 单独刷新",
            "生成记录表头": "📋 生成记录表头",
        }.get(keyword, label)

        cell.value = f'=HYPERLINK("{bat_abs}","{new_label}")'
        cell.hyperlink = Hyperlink(
            ref=cell.coordinate,
            target=bat_abs,
            tooltip=new_label,
        )
        print(f"  ✅ ({row},{col}) → {bat_name}  [{new_label}]")

    wb.save(str(xlsx))
    wb.close()


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--project", type=Path, help="项目目录(不给就升级所有)")
    args = p.parse_args()

    if args.project:
        targets = [args.project]
    else:
        targets = sorted(p for p in (ROOT / "projects").iterdir() if p.is_dir())

    for t in targets:
        if not list(t.glob("*.xlsx")):
            continue
        try:
            upgrade(t)
        except Exception as e:
            print(f"  ❌ {t.name}: {e}")


if __name__ == "__main__":
    main()
