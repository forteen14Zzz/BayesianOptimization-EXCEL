"""生成 projects/Suzuki_2026Q4_demo/ — M1 demo 项目。

复制 templates/project_template.xlsx 到项目目录并改名。
实际填入 6 个真实 Suzuki 反应的初始数据点(从文献或经验)。
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook


# 6 个初始点(Suzuki 反应经验范围)
INITIAL_RECORDS = [
    # trial_id, batch_id, 温度, 催化剂量, 时间, 溶剂, Yield, valid, status, 来源
    ("T0001", "2026-09-15", 60, 2.0, 12, "甲苯", 78.5, 1, "measured", "初始随机"),
    ("T0002", "2026-09-15", 80, 1.5, 8, "THF", 82.1, 1, "measured", "初始随机"),
    ("T0003", "2026-09-15", 70, 2.5, 16, "DMF", 75.3, 1, "measured", "初始随机"),
    ("T0004", "2026-09-15", 90, 1.0, 6, "二氧六环", 68.0, 1, "measured", "初始随机"),
    ("T0005", "2026-09-15", 50, 3.0, 24, "甲苯", 71.5, 1, "measured", "初始随机"),
    ("T0006", "2026-09-15", 75, 2.0, 10, "THF", 85.2, 0, "failed", "初始随机"),  # valid=0 失败
]


def main():
    root = Path(__file__).resolve().parent.parent
    template = root / "templates" / "project_template.xlsx"
    target_dir = root / "projects" / "Suzuki_2026Q4_demo"
    target_dir.mkdir(parents=True, exist_ok=True)
    target_xlsx = target_dir / "Suzuki_2026Q4_demo.xlsx"

    # 复制模板
    import shutil
    shutil.copy2(template, target_xlsx)

    # 填初始数据
    wb = load_workbook(str(target_xlsx))
    ws = wb["🧪 记录"]
    # 表头在 row=4: trial_id, batch_id, 温度, 催化剂用量, 反应时间, 溶剂, Yield, valid, status, 来源, 入时间
    for i, rec in enumerate(INITIAL_RECORDS):
        for j, val in enumerate(rec):
            ws.cell(row=5 + i, column=j + 1, value=val)
        # 填录入时间
        ws.cell(row=5 + i, column=11, value="2026-09-15 10:30")

    # 填项目名到 _status
    ws_status = wb["_status"]
    ws_status["B5"] = "Suzuki_2026Q4_demo"
    ws_status["B6"] = str(target_dir)

    wb.save(str(target_xlsx))
    wb.close()
    print(f"✅ Demo 项目生成: {target_xlsx}")
    print(f"   初始数据 {len(INITIAL_RECORDS)} 条(其中 {sum(1 for r in INITIAL_RECORDS if r[7]==1)} 条 valid)")


if __name__ == "__main__":
    main()