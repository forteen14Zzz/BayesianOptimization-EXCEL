"""scripts/fix_button_paths.py — 修复老项目里错误的按钮路径。

老项目里 =HYPERLINK('..\\跑_BO_点这里.bat') 解析到 projects\\跑_BO_点这里.bat
(不存在)。

修法: 改成 =HYPERLINK('跑_BO_点这里.bat') — bat 和 xlsx 同目录。
"""
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook


def fix(xlsx_path: Path) -> None:
    wb = load_workbook(str(xlsx_path))
    if "🚀 开始" not in wb.sheetnames:
        print(f"  ⚠️ 没找到 🚀 开始 sheet,跳过")
        return
    ws = wb["🚀 开始"]
    fixed = 0
    for r in range(1, 30):
        for c in range(1, 9):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            if "..\\跑_BO_点这里.bat" in v:
                new_v = v.replace("..\\跑_BO_点这里.bat", "跑_BO_点这里.bat")
                ws.cell(row=r, column=c).value = new_v
                fixed += 1
                print(f"  ({r},{c}) 修复: {v!r} → {new_v!r}")
    if fixed:
        wb.save(str(xlsx_path))
    wb.close()
    print(f"  共修复 {fixed} 个按钮路径")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--project", type=Path, required=True)
    args = p.parse_args()
    xlsx = next(args.project.glob("*.xlsx"))
    print(f"修复: {xlsx}")
    fix(xlsx)


if __name__ == "__main__":
    main()