"""portal/core/recommender.py — 写 recommend / 更新 📊 看板。

设计:
- 不动主 xlsx(用户编辑 sheet)
- 写到 .cache/recommend_<时间戳>.xlsx(单 sheet,简单)
- 用户点 [🔄 刷新最新推荐] 按钮 → pipeline 把 .cache 最新文件内容
  合并到主 xlsx 的 📊 看板

这样避免 Excel 锁冲突。
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook


CACHE_DIR = ".cache"
RECOMMEND_PREFIX = "recommend_"


# ---------- Cache 单文件 ----------
def write_recommend_cache(
    project_dir: Path,
    candidates: list[Any],  # list[Candidate]
    factor_keys: list[str],
    output_keys: list[str],
    round_num: int,
) -> Path:
    """把 candidates 写到 .cache/recommend_<时间戳>.xlsx。返回路径。"""
    cache = project_dir / CACHE_DIR
    cache.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")  # 含毫秒避免同秒重名
    out = cache / f"{RECOMMEND_PREFIX}{ts}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "推荐条件"

    headers = (
        ["trial_id", "round", "asked_at"]
        + list(factor_keys)
        + [f"{o}_mean" for o in output_keys]
        + [f"{o}_sd" for o in output_keys]
        + ["acq_value"]
    )
    ws.append(headers)
    now = datetime.now().isoformat(timespec="seconds")
    for c in candidates:
        row = [c.trial_id, round_num, now]
        for k in factor_keys:
            row.append(c.factors.get(k))
        for o in output_keys:
            row.append(c.pred_mean)  # 单输出时填 mean;多输出后续扩展
        for o in output_keys:
            row.append(c.pred_sd)
        row.append(c.acq_value)
        ws.append(row)
    wb.save(str(out))
    return out


def find_latest_cache(project_dir: Path) -> Path | None:
    """找 .cache/recommend_*.xlsx 里最新的一个。"""
    cache = project_dir / CACHE_DIR
    if not cache.exists():
        return None
    files = sorted(cache.glob(f"{RECOMMEND_PREFIX}*.xlsx"), reverse=True)
    return files[0] if files else None


# ---------- 主 xlsx 合并 ----------
def merge_cache_to_dashboard(
    project_dir: Path,
    xlsx_path: Path,
    cache_path: Path,
    round_num: int,
) -> None:
    """把 cache 推荐条件合并到 📊 看板。

    实现:
    1. 读 cache 全部行
    2. ⚠️ 根据 ⚙️ 配置 动态重写 📊 看板 row 5/48 的 headers
       (cache 列名 = 因子英文 key;同事的因子可能和模板默认不一样)
    3. 打开主 xlsx(⚠️ 可能 Excel 开着 — 失败就抛,提示用户关闭)
    4. 把候选写到 📊 看板 的"本轮推荐"区(row 6-8)
    5. 把候选追加到 "📜 历轮推荐" 区
    6. 把这一轮 BO 的运行摘要写到 "📋 运行日志" 区
    """
    wb = load_workbook(str(xlsx_path))
    try:
        # 读 cache
        cwb = load_workbook(str(cache_path), read_only=True)
        ws_c = cwb.active
        cache_rows = list(ws_c.iter_rows(values_only=True))
        cwb.close()
        headers_c = list(cache_rows[0])
        data_rows = [list(r) for r in cache_rows[1:]]

        # ⚠️ 动态更新 📊 看板 row 5 / row 48 headers,从 ⚙️ 配置读
        # 当前用户的 factor / output keys,避免 cache 英文 key 跟模板中文 header 对不上
        if "⚙️ 配置" in wb.sheetnames:
            cfg_ws = wb["⚙️ 配置"]
            factor_keys = []
            for r in range(6, 12):
                k = cfg_ws.cell(row=r, column=1).value
                if k and str(k).strip():
                    factor_keys.append(str(k).strip())
            output_keys = []
            for r in range(14, 17):
                k = cfg_ws.cell(row=r, column=1).value
                if k and str(k).strip():
                    output_keys.append(str(k).strip())
            # row 5: trial_id + factor + output_mean + output_sd + acq + round + asked_at
            headers_row5 = ["trial_id"] + factor_keys
            for o in output_keys:
                headers_row5 += [f"{o}_mean", f"{o}_sd"]
            headers_row5 += ["acq_value", "round", "asked_at"]
            dash = wb["📊 看板"]
            from openpyxl.cell.cell import MergedCell
            for i, h in enumerate(headers_row5):
                target = dash.cell(row=5, column=i + 1)
                if isinstance(target, MergedCell):
                    continue
                target.value = h
            # row 48 历轮推荐
            headers_row48 = ["round", "asked_at", "trial_id"] + factor_keys
            for o in output_keys:
                headers_row48 += [f"{o}_pred_mean", f"{o}_pred_sd"]
            headers_row48.append(f"{output_keys[0]}_actual" if output_keys else "actual")
            for i, h in enumerate(headers_row48):
                target = dash.cell(row=53, column=i + 1)
                if isinstance(target, MergedCell):
                    continue
                target.value = h

        # 写 📊 看板 (假设 DASHBOARD_SHEET = "📊 看板")
        if "📊 看板" not in wb.sheetnames:
            wb.create_sheet("📊 看板")
        ws = wb["📊 看板"]

        # 行映射:
        # row 5 = headers (我们建模板时写过的)
        # row 6-8 = 本轮推荐区(3 行)
        # row 15-44 = 收敛数据区(30 行)
        # row 49-60 = 历轮推荐(12 行)
        # row 63-90 = 运行日志(27 行)

        # 1. 本轮推荐区(覆盖 row 6-15 — 支持 batch_size=1..10)
        # ⚠️ 不再硬编码 [:3],用 len(data_rows) 让 batch_size 自适应
        n_candidates = len(data_rows)
        for i, row in enumerate(data_rows):
            for j, header in enumerate(headers_c):
                col = _col_letter(ws, header, header_row=5)
                if col is None:
                    continue
                target = ws.cell(row=6 + i, column=col)
                from openpyxl.cell.cell import MergedCell
                if isinstance(target, MergedCell):
                    continue
                target.value = row[j]

        # 2. 历轮推荐区(追加到 row 54 起,找第一个完全空的行)
        next_row = 54
        for r in range(54, ws.max_row + 1):
            # 跳过 section 标题行(_style_section 用了 merge_cells)
            from openpyxl.cell.cell import MergedCell
            v = ws.cell(row=r, column=1).value
            if isinstance(v, MergedCell) or v is None:
                break
            # 跳过 row 53 表头本身
            if r == 53:
                continue
            next_row = r + 1
        for i, row in enumerate(data_rows):
            # 用 headers_c 的列对应写到 📜 历轮推荐 的列
            # 历轮推荐 headers (row=53): round, asked_at, trial_id, 因子..., pred_mean, pred_sd, actual_yield
            history_cols = _map_cache_to_history(headers_c, ws)
            for cache_col, hist_col in history_cols.items():
                src_idx = headers_c.index(cache_col) if cache_col in headers_c else None
                if src_idx is None:
                    continue
                v = row[src_idx]
                # ⚠️ 保护:有些模板 cell 在合并区域(MergedCell),
                # 写入会抛 AttributeError。跳过。
                from openpyxl.cell.cell import MergedCell
                target_cell = ws.cell(row=next_row + i, column=hist_col)
                if isinstance(target_cell, MergedCell):
                    continue
                target_cell.value = v

        # 3. 运行日志追加
        log_next = 76
        for r in range(76, ws.max_row + 1):
            from openpyxl.cell.cell import MergedCell
            v = ws.cell(row=r, column=1).value
            if isinstance(v, MergedCell) or v is None:
                break
            log_next = r + 1
        now = datetime.now().isoformat(timespec="seconds")
        log_row = [
            now, "ask", round_num, "", "sobo", "", len(data_rows),
            "ok", f"merged {len(data_rows)} candidates from {cache_path.name}", "",
        ]
        from openpyxl.cell.cell import MergedCell
        for j, v in enumerate(log_row[:10]):
            target = ws.cell(row=log_next, column=j + 1)
            if isinstance(target, MergedCell):
                continue
            target.value = v

        wb.save(str(xlsx_path))
    finally:
        wb.close()


def _col_letter(ws: Any, header_value: str, header_row: int) -> int | None:
    """在 ws 的指定 row 找 header,返回列号(1-based)。找不到返回 None。"""
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is not None and str(v).strip() == header_value:
            return c
    return None


def _map_cache_to_history(cache_headers: list[str], ws: Any) -> dict[str, int]:
    """把 cache 列名映射到历轮推荐 sheet 的列号。

    历轮推荐 row=53 headers: round, asked_at, trial_id, 因子..., pred_mean, pred_sd, actual_yield
    """
    result = {}
    for j, h in enumerate(cache_headers):
        if h in ("trial_id", "asked_at", "round"):
            col = _col_letter(ws, h, header_row=53)
            if col:
                result[h] = col
        elif h.endswith("_mean"):
            # 简化: 用第一个 output 的 pred mean(单输出场景)
            base = h.replace("_mean", "")
            col = _col_letter(ws, f"{base}_pred_mean", header_row=53)
            if col:
                result[h] = col
        elif h.endswith("_sd"):
            base = h.replace("_sd", "")
            col = _col_letter(ws, f"{base}_pred_sd", header_row=53)
            if col:
                result[h] = col
        else:
            # 因子列
            col = _col_letter(ws, h, header_row=53)
            if col:
                result[h] = col
    return result