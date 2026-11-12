"""scripts/upgrade_project_buttons.py — 给老项目补上缺失的按钮。

老项目(模板更新前创建的)只有:
  - 🚀 启动 BO
  - 🔄 单独刷新

新模板加了:
  - 📋 生成记录表头 (bat sync)

这个脚本把缺的按钮补上(不破坏已有的)。
"""
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.hyperlink import Hyperlink


BUTTON_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
BUTTON_FILL = PatternFill("solid", fgColor="70AD47")
HINT_FONT = Font(name="微软雅黑", size=9, italic=True, color="808080")


def upgrade(xlsx_path: Path) -> None:
    wb = load_workbook(str(xlsx_path))
    if "🚀 开始" not in wb.sheetnames:
        wb.close()
        print("  ⚠️ 没找到 🚀 开始 sheet,跳过")
        return

    ws = wb["🚀 开始"]
    bat_abs = str((xlsx_path.parent / "跑_BO_点这里.bat").resolve())

    # 检查现有按钮
    has_sync = False
    has_bo = False
    for r in range(1, 30):
        for c in range(1, 9):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "HYPERLINK" in v:
                if "sync" in v:
                    has_sync = True
                elif "跑_BO" in v and "refresh" not in v:
                    has_bo = True

    print(f"  现有: 🚀启动BO={has_bo}, 📋生成表头={has_sync}")

    if not has_sync:
        # 找空行放新按钮(row 20 优先)
        row = 20
        while ws.cell(row=row, column=1).value is not None:
            row += 1

        cell = ws.cell(row=row, column=1)
        cell.value = f'=HYPERLINK("{bat_abs} sync","📋 生成记录表头")'
        cell.font = BUTTON_FONT
        cell.fill = BUTTON_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.hyperlink = Hyperlink(
            ref=cell.coordinate,
            target=bat_abs,
            tooltip="按 ⚙️ 配置 重建 🧪 记录 表头",
        )

        hint = ws.cell(row=row + 1, column=1)
        hint.value = "(按 ⚙️ 配置 的因子/目标重建 🧪 记录 表头 — 填完配置先点这个,再去 🧪 记录 填数据)"
        hint.font = HINT_FONT
        ws.merge_cells(start_row=row + 1, end_row=row + 1, start_column=1, end_column=8)
        print(f"  ✅ 已添加 [📋 生成记录表头] 按钮 (row {row})")
    else:
        print("  ℹ️ 按钮已存在,跳过")

    wb.save(str(xlsx_path))
    wb.close()


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--project", type=Path, required=True)
    args = p.parse_args()
    xlsx = next(args.project.glob("*.xlsx"))
    print(f"升级: {xlsx}")
    upgrade(xlsx)


if __name__ == "__main__":
    main()
