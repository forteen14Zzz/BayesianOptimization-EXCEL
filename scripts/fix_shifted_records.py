"""scripts/fix_shifted_records.py — 修复 🧪 记录 数据列错位。

场景(M4 现场 zzz12):
同事填数据时按"自己理解的顺序"填,和 row 4 的表头对不上。
zzz12 的 5 行数据整体右移了 3 列:
    表头: trial_id | batch_id | base | ligand | solvent | concentration | temperature | Yield | valid
    数据: T0001    | B2026... | 90  | 0.1    | 90      | KOPiv           | JackiePhos   | p-Xylene | 8.98
                              ↑温度 ↑浓度 ↑温度

怎么修:
用**值的类型特征**推断每个错位行的正确归属:
- numeric 值 + 该因子是 continuous/discrete → 归到那个因子
- 字符串值 + 该因子是 categorical → 归到那个因子
- 目标列的值必须在 [lower_bound, upper_bound] 范围内

一行数据只要能**唯一匹配**,就自动重排。匹配不唯一或冲突的行跳过(报给用户)。
"""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from portal.core.excel_io import FactorSpec, OutputSpec, read_config


RECORD_SHEET = "🧪 记录"
HEADER_ROW = 4
DATA_START_ROW = 5

FIXED_COLS = {"trial_id", "batch_id", "valid", "status", "来源", "录入时间"}


def _is_number(v: Any) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    try:
        float(v)
        return True
    except (ValueError, TypeError):
        return False


def _matches_factor(v: Any, spec: FactorSpec) -> bool:
    """这个值是否可能是该因子的取值(宽松 — 只用于判断"是否错位")。

    ⚠️ discrete 类型用宽松规则(数值即可):
    GP 推荐的连续值可能不在候选列表里(如 concentration=0.0826,
    候选是 0.057/0.1/0.153),这不算错位。
    """
    if v is None or (isinstance(v, str) and not v.strip()):
        return False  # ⚠️ 空值不参与匹配(否则会匹配所有列)

    if spec.type in ("continuous", "discrete"):
        return _is_number(v)

    # categorical
    cats = [str(c) for c in (spec.categories or [])]
    if not cats:
        return False
    return str(v) in cats


def _matches_factor_strict(v: Any, spec: FactorSpec) -> bool:
    """这个值是否**严格**属于该因子(用于消除歧义)。

    discrete 类型严格匹配候选列表 — 否则温度 90 会同时匹配
    concentration 和 Yield,造成歧义。
    """
    if v is None or (isinstance(v, str) and not v.strip()):
        return False
    if spec.type == "discrete":
        cats = [str(c) for c in (spec.categories or [])]
        if cats:
            for c in cats:
                try:
                    if abs(float(v) - float(c)) < 1e-6:
                        return True
                except (ValueError, TypeError):
                    if str(v) == c:
                        return True
            return False
        return _is_number(v)
    return _matches_factor(v, spec)


def _matches_output(v: Any, spec: OutputSpec) -> bool:
    """这个值是否可能是该目标的值。

    ⚠️ 产率范围通常是 0-100,和温度范围重叠。
    所以目标列的匹配要**更严**:
    - 必须是在 bounds 内
    - 而且**不能**是某个 discrete 因子的候选值(那是因子不是目标)
    """
    if v is None or (isinstance(v, str) and not v.strip()):
        return False
    if not _is_number(v):
        return False
    f = float(v)
    if spec.lower_bound is not None and f < spec.lower_bound:
        return False
    if spec.upper_bound is not None and f > spec.upper_bound:
        return False
    return True


def _is_discrete_candidate(v: Any, factor_specs: list[FactorSpec]) -> bool:
    """这个值是不是某个 discrete 因子的候选值(如 temperature 的 90)。"""
    if not _is_number(v):
        return False
    for spec in factor_specs:
        if spec.type != "discrete":
            continue
        for c in (spec.categories or []):
            try:
                if abs(float(v) - float(c)) < 1e-6:
                    return True
            except (ValueError, TypeError):
                pass
    return False


def fix_row(
    values: list[Any],
    headers: list[str],
    factor_specs: list[FactorSpec],
    output_specs: list[OutputSpec],
) -> tuple[dict[str, Any], bool]:
    """尝试修正一行。返回 (新值 dict, 是否成功)。

    算法:
    1. 对每个值,算出它**可能属于哪些列**(按类型/候选列表/bounds 匹配)
    2. 用贪心 + 唯一性判断做分配:
       - 某个列只有一个候选值 → 直接分配
       - 某个值只能属于一个列 → 直接分配
    3. 剩下的歧义 → 保持原位,标记 ok=False
    """
    factor_keys = [s.key for s in factor_specs]
    output_keys = [o.key for o in output_specs]
    all_keys = factor_keys + output_keys

    # 1. value_idx → possible columns
    #    第一轮:宽松匹配(所有可能)
    #    第二轮:严格匹配(消歧 — 优先用严格结果)
    #
    # ⚠️ 固定列(trial_id/batch_id/valid/status/来源/录入时间)的位置上的值
    # **不参与因子/目标匹配** — 否则 valid 列的 8.98 会被当成 Yield 抢走。
    # 同事填错位时,固定列也错位了(valid 位置存的是 Yield 值),
    # 所以固定列的值要**重新推导**而不是直接抄。
    fixed_idxs = {
        i for i, h in enumerate(headers) if h in FIXED_COLS
    }
    possible: dict[int, list[str]] = {}
    possible_strict: dict[int, list[str]] = {}
    for idx, v in enumerate(values):
        if idx >= len(headers):
            break
        # 固定列位置上的值先记下来,不参与匹配
        if idx in fixed_idxs:
            continue
        cols = []
        cols_strict = []
        for spec in factor_specs:
            if _matches_factor(v, spec):
                cols.append(spec.key)
            if _matches_factor_strict(v, spec):
                cols_strict.append(spec.key)
        # 目标列:在 bounds 内 且 不是某个 discrete 因子的候选值
        # (温度 90 也是合法产率,但显然是温度不是产率)
        for spec in output_specs:
            if _matches_output(v, spec) and not _is_discrete_candidate(v, factor_specs):
                cols.append(spec.key)
                cols_strict.append(spec.key)
        if cols:
            possible[idx] = cols
        if cols_strict:
            possible_strict[idx] = cols_strict

    # 用严格匹配优先(消歧),严格为空时退回宽松
    def _candidates(idx: int) -> list[str]:
        return possible_strict.get(idx) or possible.get(idx, [])

    # 2. 贪心分配
    assigned: dict[str, Any] = {}   # column → value
    used_idx: set[int] = set()

    def try_assign(col: str, idx: int) -> bool:
        """尝试把 idx 分配给 col。成功返回 True。"""
        if col in assigned or idx in used_idx:
            return False
        assigned[col] = values[idx]
        used_idx.add(idx)
        return True

    # 2a. 列只有一个候选 → 先定
    for col in all_keys:
        idxs = [i for i in possible if col in _candidates(i) and i not in used_idx]
        if len(idxs) == 1:
            try_assign(col, idxs[0])

    # 2b. 值只能属于一个列 → 再定
    for idx in possible:
        if idx in used_idx:
            continue
        free = [c for c in _candidates(idx) if c not in assigned]
        if len(free) == 1:
            try_assign(free[0], idx)

    # 2c. 剩下的:如果有多个候选但能唯一匹配(列和值一一对应),也定
    remaining_cols = [c for c in all_keys if c not in assigned]
    remaining_idxs = [i for i in possible if i not in used_idx]
    if len(remaining_cols) == len(remaining_idxs) == 1:
        try_assign(remaining_cols[0], remaining_idxs[0])

    # 3. 组装结果
    result: dict[str, Any] = {}
    ok = True
    for key in all_keys:
        if key in assigned:
            result[key] = assigned[key]
        else:
            # 没匹配上 → 保持原位
            orig_idx = headers.index(key) if key in headers else None
            if orig_idx is not None and orig_idx < len(values):
                result[key] = values[orig_idx]
            else:
                result[key] = None
            ok = False  # 该列没匹配上

    # 检查是否所有列都填上了
    for key in factor_keys + output_keys:
        if result.get(key) is None:
            ok = False

    return result, ok


def fix_fixed_columns(
    values: list[Any],
    headers: list[str],
    assigned: dict[str, Any],
    used_idx: set[int],
    factor_keys: list[str],
    output_keys: list[str],
) -> dict[str, Any]:
    """重新推导固定列的值。

    同事错位时固定列也错位了(比如 valid 位置存的是 Yield 值)。
    推导规则:
    - trial_id / batch_id:原样保留(在最前面,不会错)
    - valid: 值必须是 0/1/空。错位时原位置的值不是 0/1 →
      找"没被因子/目标用掉"的值里有没有 0/1
    - status:measured/pending/failed 之一
    - 来源 / 录入时间:原样保留(在最后,不会错)
    """
    out: dict[str, Any] = {}

    # trial_id / batch_id — 原样
    for key in ("trial_id", "batch_id"):
        idx = headers.index(key) if key in headers else None
        out[key] = values[idx] if (idx is not None and idx < len(values)) else None

    # 找没被用掉的值(可能是被挤走的 valid/status)
    leftover = [
        (i, values[i]) for i in range(len(values))
        if i not in used_idx and i not in fixed_idxs_of(headers)
    ]

    # valid
    valid_idx = headers.index("valid") if "valid" in headers else None
    valid_val = values[valid_idx] if (valid_idx is not None and valid_idx < len(values)) else None
    if valid_val is None or str(valid_val).strip() == "" or str(valid_val) in ("0", "1"):
        out["valid"] = valid_val if valid_val is not None else 1
    else:
        # 找 leftover 里的 0/1
        found = None
        for i, v in leftover:
            if v is not None and str(v).strip() in ("0", "1"):
                found = v
                break
        out["valid"] = found if found is not None else 1

    # status
    status_idx = headers.index("status") if "status" in headers else None
    status_val = values[status_idx] if (status_idx is not None and status_idx < len(values)) else None
    if status_val is not None and str(status_val).strip() in ("measured", "pending", "failed"):
        out["status"] = str(status_val).strip()
    else:
        out["status"] = "measured"

    # 来源 / 录入时间 — 原样
    for key in ("来源", "录入时间"):
        idx = headers.index(key) if key in headers else None
        out[key] = values[idx] if (idx is not None and idx < len(values)) else None

    return out


def fixed_idxs_of(headers: list[str]) -> set[int]:
    return {i for i, h in enumerate(headers) if h in FIXED_COLS}


def fix(xlsx_path: Path, dry_run: bool = False) -> dict:
    cfg = read_config(xlsx_path)
    factor_specs = cfg.factors
    output_specs = cfg.outputs
    factor_keys = [s.key for s in factor_specs]
    output_keys = [o.key for o in output_specs]

    wb = load_workbook(str(xlsx_path))
    try:
        if RECORD_SHEET not in wb.sheetnames:
            return {"fixed": 0, "skipped": 0}

        ws = wb[RECORD_SHEET]
        headers = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=HEADER_ROW, column=c).value
            if v is None:
                break
            headers.append(str(v).strip())

        n_fixed = 0
        n_skipped = 0
        details = []

        for r in range(DATA_START_ROW, ws.max_row + 1):
            values = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
            if not values or not values[0]:
                continue
            if str(values[0]).startswith(("💡", "📌", "提示")):
                continue

            # 当前是否已经对齐?(每个列的值都符合该列类型)
            # ⚠️ 空值不算错位证据(Yield 还没测是正常的)
            aligned = True
            for spec in factor_specs:
                idx = headers.index(spec.key) if spec.key in headers else None
                if idx is None:
                    aligned = False
                    break
                v = values[idx]
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue  # 空值跳过
                if not _matches_factor(v, spec):
                    aligned = False
                    break
            for spec in output_specs:
                idx = headers.index(spec.key) if spec.key in headers else None
                if idx is None:
                    aligned = False
                    break
                v = values[idx]
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue  # 空值跳过
                if not _matches_output(v, spec):
                    aligned = False
                    break

            if aligned:
                continue  # 已经对齐,不动

            new_row, ok = fix_row(values, headers, factor_specs, output_specs)
            if ok:
                n_fixed += 1
                details.append((r, values, new_row))
                if not dry_run:
                    for key, val in new_row.items():
                        if key in headers:
                            ws.cell(row=r, column=headers.index(key) + 1).value = val
                    # 固定列也照抄(它们没参与匹配)
                    for key in FIXED_COLS:
                        if key in headers and key not in new_row:
                            idx = headers.index(key)
                            if idx < len(values):
                                ws.cell(row=r, column=idx + 1).value = values[idx]
            else:
                n_skipped += 1

        if not dry_run and n_fixed:
            wb.save(str(xlsx_path))
        wb.close()

        return {
            "fixed": n_fixed,
            "skipped": n_skipped,
            "details": details,
            "headers": headers,
            "factor_keys": factor_keys,
            "output_keys": output_keys,
        }
    finally:
        wb.close()


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--project", type=Path, required=True)
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args()
    xlsx = next(args.project.glob("*.xlsx"))
    print(f"检查: {xlsx}")
    result = fix(xlsx, dry_run=args.dry_run)

    print(f"\n已修正 {result['fixed']} 行,跳过(有歧义){result['skipped']} 行")
    if result.get("details"):
        print("\n修正详情:")
        for r, old, new in result["details"]:
            print(f"  row {r}:")
            for k in result["factor_keys"] + result["output_keys"]:
                if old[result["headers"].index(k)] != new.get(k):
                    print(f"    {k:<15} {old[result['headers'].index(k)]!r} → {new.get(k)!r}")
    if args.dry_run:
        print("\n(dry-run, 未写入)")


if __name__ == "__main__":
    main()
