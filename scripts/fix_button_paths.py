"""scripts/fix_button_paths.py — 修复老项目里错误的按钮路径 + 补 hyperlink 对象。

两个关键点(2026-10-08 实测):

1. **绝对路径** — Excel 解析相对路径用 Excel 当前工作目录,
   不是 xlsx 所在目录 → 相对路径必失效。必须用绝对路径。

2. **真 hyperlink 对象** — openpyxl 只写 `=HYPERLINK()` 公式时,
   Excel 打开后 `Range.Hyperlinks.Count` = 0,不认这是超链接,点了没反应。
   必须同时设 `cell.hyperlink = Hyperlink(...)`。

用法:
  python scripts/fix_button_paths.py --project projects/z13
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink


def fix(xlsx_path: Path) -> int:
    """修按钮路径 + 补 hyperlink 对象。返回修复数量。"""
    wb = load_workbook(str(xlsx_path))
    if "🚀 开始" not in wb.sheetnames:
        wb.close()
        print("  ⚠️ 没找到 🚀 开始 sheet,跳过")
        return 0

    ws = wb["🚀 开始"]
    bat_abs = str((xlsx_path.parent / "跑_BO_点这里.bat").resolve())
    bat_target = bat_abs  # hyperlink.target 必须是不带参数的纯路径

    fixed = 0
    for r in range(1, 30):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if not isinstance(v, str) or "HYPERLINK" not in v:
                continue

            m = re.search(r'HYPERLINK\("([^"]*)"', v)
            if not m:
                continue
            old_href = m.group(1)

            # ⚠️ 按**标签**识别按钮类型 — 老的项目里
            # "🚀 启动 BO(手动)" 错指到了 '#⚙️ 配置'!A1(内部跳转),
            # 要改成调 bat。
            label = ""
            lm = re.search(r'HYPERLINK\("[^"]*","([^"]*)"\)', v)
            if lm:
                label = lm.group(1)

            is_bo_btn = "启动 BO" in label
            is_refresh_btn = "刷新" in label
            is_sync_btn = "生成记录表头" in label

            if (is_bo_btn or is_refresh_btn or is_sync_btn) and old_href.startswith("#"):
                # 该按钮应指向 bat,不是内部跳转
                param = ""
                if is_refresh_btn:
                    param = " refresh"
                elif is_sync_btn:
                    param = " sync"
                new_href = bat_abs + param
                # 顺便把老标签规范化("🚀 启动 BO(手动)" → "🚀 启动 BO")
                new_label = label
                if is_bo_btn:
                    new_label = "🚀 启动 BO"
                elif is_refresh_btn:
                    new_label = "🔄 单独刷新"
                cell.value = f'=HYPERLINK("{new_href}","{new_label}")'
                cell.hyperlink = Hyperlink(
                    ref=cell.coordinate,
                    target=bat_target,
                    tooltip=new_label,
                )
                fixed += 1
                print(f"  ({r},{c}) 按钮改指向 bat: {label!r} → {new_label!r}")
                continue

            # 内部跳转(#'sheet'!A1)— 只需补 location 型 hyperlink
            if old_href.startswith("#"):
                loc = old_href[1:]  # 去掉 #
                cell.hyperlink = Hyperlink(
                    ref=cell.coordinate,
                    location=loc,
                    tooltip="",
                )
                fixed += 1
                print(f"  ({r},{c}) 补内部链接对象: {loc}")
                continue

            # bat 链接 — 必须绝对路径
            if "跑_BO_点这里.bat" not in old_href:
                continue

            # ⚠️ 参数按**标签**判定(不能只看 href 结尾 — 之前被误剥过)
            param = ""
            if is_refresh_btn:
                param = " refresh"
            elif is_sync_btn:
                param = " sync"
            else:
                # 兜底:href 结尾有参数就保留
                for p in ("refresh", "sync"):
                    if old_href.rstrip().endswith(p):
                        param = f" {p}"
                        break
            new_href = bat_abs + param
            if new_href != old_href:
                cell.value = v.replace(f'HYPERLINK("{old_href}"', f'HYPERLINK("{new_href}"')
                print(f"  ({r},{c}) 修路径: {old_href!r} → {new_href!r}")

            # ⚠️ hyperlink.target 必须是**纯文件路径**(不带参数)
            cell.hyperlink = Hyperlink(
                ref=cell.coordinate,
                target=bat_target,
                tooltip=label or "双击运行 BO",
            )
            fixed += 1

    if fixed:
        wb.save(str(xlsx_path))
    wb.close()
    print(f"  共修复 {fixed} 个按钮")
    return fixed


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--project", type=Path, required=True)
    args = p.parse_args()
    xlsx = next(args.project.glob("*.xlsx"))
    print(f"修复: {xlsx}")
    fix(xlsx)


if __name__ == "__main__":
    main()