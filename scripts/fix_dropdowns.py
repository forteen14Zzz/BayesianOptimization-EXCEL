"""scripts/fix_dropdowns.py — 修复老项目的下拉菜单范围。

问题(2026-10-08 同事 zzz12 反馈):
- ⚙️ 配置 sheet 的 B 列下拉范围是 B11:B105 / B6:B105,
  一路跨到 Outputs(row 14+) / Strategy(row 21+) / Constraints(row 33+)
- 所以同事在"实验目标"、"策略"、"约束"三个表格的 B 列点单元格,
  看到的却是实验因子表格的下拉选项(continuous/discrete/categorical)

修复:
- B6:B11  → Factors.type
- C6:C11  → Factors.trainable
- B14:B16 → Outputs.direction
- B21     → Strategy.策略类型
- B22     → Strategy.采集函数
- B27     → Strategy.存模型
- B33:B35 → Constraints.type

用法:
  python scripts/fix_dropdowns.py --project projects/zzz12
"""
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation


# (范围, 选项列表) — 必须和 build_project_template.py 保持一致
CONFIG_VALIDATIONS = [
    ("B6:B11", '"continuous,discrete,categorical,categorical_descriptor,categorical_molecular"'),
    ("C6:C11", '"TRUE,FALSE"'),
    ("B14:B16", '"maximize,minimize,target"'),
    ("B21", '"random,sobo,mobo,doe"'),
    ("B22", '"qLogEI,qLogNEI,qEI,qUCB"'),
    ("B27", '"TRUE,FALSE"'),
    ("B33:B35", '"linear_inequality,linear_equality"'),
]

RECORD_VALIDATIONS = [
    ("H5:H105", '"0,1"'),
    ("J5:J105", '"measured,pending,failed"'),
]


def fix_sheet(ws, specs: list[tuple[str, str]]) -> int:
    """清掉 ws 上所有 dataValidation,按 specs 重建。"""
    # 记录旧的,清空
    ws.data_validations.dataValidation = []
    count = 0
    for sqref, formula in specs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(sqref)
        count += 1
    return count


def fix(xlsx_path: Path) -> int:
    wb = load_workbook(str(xlsx_path))
    total = 0

    if "⚙️ 配置" in wb.sheetnames:
        n = fix_sheet(wb["⚙️ 配置"], CONFIG_VALIDATIONS)
        print(f"  ⚙️ 配置: {n} 个下拉")
        total += n

    if "🧪 记录" in wb.sheetnames:
        n = fix_sheet(wb["🧪 记录"], RECORD_VALIDATIONS)
        print(f"  🧪 记录: {n} 个下拉")
        total += n

    # ⚠️ atomic write — Excel 开着时普通 save 会 Permission denied
    # 写到 .tmp 再替换;失败重试几次(给 Excel 释放锁的时间)
    import os
    import shutil
    import time

    tmp = xlsx_path.with_suffix(xlsx_path.suffix + ".tmp")
    wb.save(str(tmp))
    wb.close()

    last_err = None
    for _ in range(5):
        try:
            if xlsx_path.exists():
                os.remove(xlsx_path)
            shutil.move(str(tmp), str(xlsx_path))
            break
        except PermissionError as e:
            last_err = e
            time.sleep(1.5)
    else:
        # 锁着 → 提示用户关 Excel
        print("  [WARN] 文件被占用(Excel 可能开着),请关闭 Excel 后重试")
        print(f"         新文件已存到: {tmp}")
        raise last_err  # type: ignore[misc]

    return total


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--project", type=Path, required=True)
    args = p.parse_args()
    xlsx = next(args.project.glob("*.xlsx"))
    print(f"修复: {xlsx}")
    total = fix(xlsx)
    print(f"  共 {total} 个下拉")


if __name__ == "__main__":
    main()
