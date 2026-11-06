"""portal/core/sync_headers.py — 🧪 记录 表头跟随 ⚙️ 配置 自动同步。

问题(2026-10 M4 现场,z13 + zzz12 两次踩到):
同事在 ⚙️ 配置 里定义的因子顺序是 base / ligand / solvent / concentration / temperature,
但 🧪 记录 表头是模板预设的 温度 / 催化剂用量 / 反应时间 / 溶剂。
同事按"自己理解的顺序"填数据 → 列错位 → BO 学到乱七八糟的关系。

解法:
**以 ⚙️ 配置 为唯一真值源** — 每次 BO 跑之前,自动把 🧪 记录 的表头
重建成「trial_id | batch_id | 因子... | 目标... | valid | status | 来源 | 录入时间」。

数据迁移:
- 按**列名**匹配旧表头 → 新表头,把数据搬到正确位置
- 旧表头长度和新表头一致时,按**位置**兜底
- 找不到对应数据的新列留空
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


RECORD_SHEET = "🧪 记录"
HEADER_ROW = 4
DATA_START_ROW = 5

# 表头样式(和 build_project_template.py 保持一致)
TABLE_HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
TABLE_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

# 固定列(在因子/目标之前)
PREFIX_COLS = ["trial_id", "batch_id"]
# 固定列(在因子/目标之后)
SUFFIX_COLS = ["valid", "status", "来源", "录入时间"]


def build_record_headers(factor_keys: list[str], output_keys: list[str]) -> list[str]:
    """按 配置 生成 🧪 记录 的表头列表。"""
    return (
        list(PREFIX_COLS)
        + list(factor_keys)
        + list(output_keys)
        + list(SUFFIX_COLS)
    )


def _norm(s: Any) -> str:
    """归一化表头名(忽略大小写/空格/全角空格)。"""
    return str(s).strip().lower().replace(" ", "").replace("　", "")


def _read_old_headers(ws: Any) -> list[str]:
    """读当前 🧪 记录 的表头(row 4)。"""
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if v is None:
            break
        headers.append(str(v).strip())
    return headers


def _read_data_rows(ws: Any) -> list[dict[str, Any]]:
    """读现有数据行,返回 [{header: value}]。跳过空行和提示行。"""
    headers = _read_old_headers(ws)
    if not headers:
        return []
    rows = []
    for r in range(DATA_START_ROW, ws.max_row + 1):
        row = {}
        has_data = False
        for i, h in enumerate(headers):
            v = ws.cell(row=r, column=i + 1).value
            row[h] = v
            if v is not None and str(v).strip() != "":
                has_data = True
        if not has_data:
            continue
        tid = row.get("trial_id")
        if isinstance(tid, str) and tid.startswith(("💡", "📌", "📎", "提示")):
            continue
        rows.append(row)
    return rows


def _migrate_row(
    row: dict[str, Any],
    old_headers: list[str],
    new_headers: list[str],
) -> dict[str, Any]:
    """把一行数据按列名迁移到新表头。

    优先级:
      1. 同名(精确)
      2. 同名(归一化后)
      3. 位置兜底(新旧表头长度一致时)
    """
    norm_index: dict[str, Any] = {}
    for k, v in row.items():
        norm_index.setdefault(_norm(k), v)

    same_length = len(old_headers) == len(new_headers)
    new_row: dict[str, Any] = {}
    for idx, nh in enumerate(new_headers):
        val = row.get(nh)
        if val is None:
            val = norm_index.get(_norm(nh))
        if val is None and same_length:
            val = row.get(old_headers[idx])
        new_row[nh] = val
    return new_row


def sync_record_headers(
    xlsx_path: Path,
    factor_keys: list[str],
    output_keys: list[str],
) -> dict:
    """同步 🧪 记录 表头到 配置 的因子/目标。

    返回:
      {"changed": bool, "old_headers": [...], "new_headers": [...], "migrated": int}
    """
    new_headers = build_record_headers(factor_keys, output_keys)

    wb = load_workbook(str(xlsx_path))
    try:
        if RECORD_SHEET not in wb.sheetnames:
            return {"changed": False, "old_headers": [], "new_headers": new_headers, "migrated": 0}

        ws = wb[RECORD_SHEET]
        old_headers = _read_old_headers(ws)

        if old_headers == new_headers:
            # 表头一致 → 只刷样式和下拉,不动数据
            _style_header(ws, len(new_headers))
            _style_data(ws, len(new_headers), 0)
            _sync_validations(ws, len(new_headers))
            wb.save(str(xlsx_path))
            return {
                "changed": False,
                "old_headers": old_headers,
                "new_headers": new_headers,
                "migrated": 0,
            }

        # 表头变了 → 迁移数据
        old_rows = _read_data_rows(ws)
        migrated_rows = [
            _migrate_row(row, old_headers, new_headers) for row in old_rows
        ]

        # 清掉旧数据区
        max_col = max(ws.max_column, len(new_headers))
        for r in range(DATA_START_ROW, ws.max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = None

        # 写新表头 + 清多余列
        for i, h in enumerate(new_headers):
            ws.cell(row=HEADER_ROW, column=i + 1).value = h
        for c in range(len(new_headers) + 1, max_col + 1):
            cell = ws.cell(row=HEADER_ROW, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

        # 写数据
        for i, row in enumerate(migrated_rows):
            r = DATA_START_ROW + i
            for j, h in enumerate(new_headers):
                ws.cell(row=r, column=j + 1).value = row.get(h)

        _style_header(ws, len(new_headers))
        _style_data(ws, len(new_headers), len(migrated_rows))
        _sync_validations(ws, len(new_headers))

        wb.save(str(xlsx_path))
        return {
            "changed": True,
            "old_headers": old_headers,
            "new_headers": new_headers,
            "migrated": len(migrated_rows),
        }
    finally:
        wb.close()


def _style_header(ws: Any, n_cols: int) -> None:
    """给表头行上样式。"""
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=HEADER_ROW, column=c)
        cell.font = TABLE_HEADER_FONT
        cell.fill = TABLE_HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[HEADER_ROW].height = 32


def _style_data(ws: Any, n_cols: int, n_rows: int) -> None:
    """给数据区上样式(浅黄底 = 用户填)。"""
    for r in range(DATA_START_ROW, DATA_START_ROW + max(n_rows, 8)):
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = INPUT_FILL
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center")


def _sync_validations(ws: Any, n_cols: int) -> None:
    """同步 🧪 记录 的下拉菜单范围(valid / status 列跟着表头走)。"""
    from openpyxl.worksheet.datavalidation import DataValidation

    valid_col = None
    status_col = None
    for c in range(1, n_cols + 1):
        h = ws.cell(row=HEADER_ROW, column=c).value
        if h == "valid":
            valid_col = c
        elif h == "status":
            status_col = c

    # 清掉本 sheet 所有的旧验证(只影响 🧪 记录)
    ws.data_validations.dataValidation = []

    if valid_col:
        dv_valid = DataValidation(type="list", formula1='"0,1"', allow_blank=True)
        ws.add_data_validation(dv_valid)
        dv_valid.add(f"{get_column_letter(valid_col)}{DATA_START_ROW}:{get_column_letter(valid_col)}105")
    if status_col:
        dv_status = DataValidation(type="list", formula1='"measured,pending,failed"', allow_blank=True)
        ws.add_data_validation(dv_status)
        dv_status.add(f"{get_column_letter(status_col)}{DATA_START_ROW}:{get_column_letter(status_col)}105")
