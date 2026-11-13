"""portal/core/excel_io.py — Excel 读写 + schema 校验。

设计:
- 读 ⚙️ 配置 / 🧪 记录 / _status 都走 read_only=True(不锁 Excel)
- 写系统 sheet 也走 openpyxl(不动用户编辑的 ⚙️🎯🧪)
- schema 校验用 pydantic(失败给友好错误,不动 Excel)

约定:
- CONFIG_SHEET = "⚙️ 配置"
- RECORD_SHEET = "🧪 记录"
- STATUS_SHEET = "_status"
- DASHBOARD_SHEET = "📊 看板"
"""
from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field, field_validator, model_validator


# ---------- 常量 ----------
CONFIG_SHEET = "⚙️ 配置"
RECORD_SHEET = "🧪 记录"
STATUS_SHEET = "_status"
DASHBOARD_SHEET = "📊 看板"

SUPPORTED_FACTOR_TYPES = {
    "continuous",
    "discrete",
    "categorical",
    "categorical_descriptor",
    "categorical_molecular",
}
SUPPORTED_DIRECTIONS = {"maximize", "minimize", "target"}


# ---------- Pydantic 模型 ----------
class FactorSpec(BaseModel):
    """单因子定义。"""
    key: str = Field(..., min_length=1)
    type: str = Field(..., description="continuous|discrete|categorical|categorical_descriptor|categorical_molecular")
    trainable: bool = True
    low: float | None = None
    high: float | None = None
    categories: list[str] | None = None
    unit: str | None = None
    notes: str | None = None

    @field_validator("type")
    @classmethod
    def _v_type(cls, v: str) -> str:
        if v not in SUPPORTED_FACTOR_TYPES:
            raise ValueError(
                f"type 必须是 {SUPPORTED_FACTOR_TYPES}, 收到: {v!r}"
            )
        return v

    @model_validator(mode="after")
    def _check_ranges(self):
        # continuous 必须 low/high;discrete 可选 low/high(categories 已经隐含范围)
        if self.type == "continuous":
            if self.low is None or self.high is None:
                raise ValueError(f"{self.key}: continuous 必须填 low 和 high")
            if self.low >= self.high:
                raise ValueError(f"{self.key}: low 必须 < high")
        # ⚠️ discrete 也必须填 categories ——
        #    不填的话 BoFire 拿不到候选值,只能当连续变量处理
        #    (实测:同事只填了 low/high,推荐出来的是连续值)
        if self.type == "discrete":
            if not self.categories:
                raise ValueError(
                    f"{self.key}: discrete 必须填 categories(用 ; 分隔),"
                    f"比如 0.5;1.0;2.0;5.0 — "
                    f"只填 low/high 的话会被当成连续变量"
                )
            if len(self.categories) < 2:
                raise ValueError(f"{self.key}: 至少 2 个候选值,当前 {len(self.categories)} 个")
            # 候选值必须能转 float
            for c in self.categories:
                try:
                    float(c)
                except (TypeError, ValueError) as e:
                    raise ValueError(
                        f"{self.key}: discrete 的候选值必须是数字(当前 {c!r}) — "
                        f"非数字的离散选项请用 categorical"
                    ) from e
        if self.type in ("categorical", "categorical_descriptor", "categorical_molecular"):
            if not self.categories:
                raise ValueError(f"{self.key}: {self.type} 必须填 categories(用 ; 分隔)")
            if len(self.categories) < 2:
                raise ValueError(f"{self.key}: 至少 2 个 category,候选仅 {len(self.categories)} 个")
        return self


class OutputSpec(BaseModel):
    """单目标定义。"""
    key: str = Field(..., min_length=1)
    direction: str  # maximize|minimize|target
    target: float | None = None  # 仅 target 有
    lower_bound: float | None = None
    upper_bound: float | None = None
    unit: str | None = None

    @field_validator("direction")
    @classmethod
    def _v_dir(cls, v: str) -> str:
        if v not in SUPPORTED_DIRECTIONS:
            raise ValueError(
                f"direction 必须是 {SUPPORTED_DIRECTIONS}, 收到: {v!r}"
            )
        return v

    @model_validator(mode="after")
    def _check(self):
        if self.direction == "target" and self.target is None:
            raise ValueError(f"{self.key}: direction=target 必须填 target 值")
        return self


class StrategySpec(BaseModel):
    """BO 策略配置。"""
    strategy_type: str = "sobo"  # sobo|random|mobo|doe
    acq_function: str = "qLogEI"  # qLogEI|qLogNEI|qEI|qUCB
    n_initial: int = Field(5, ge=3)
    batch_size: int = Field(3, ge=1, le=20)
    convergence_threshold: float = Field(0.01, gt=0, le=1)
    convergence_window: int = Field(5, ge=3, le=20)
    save_model: bool = False


class ConfigSheet(BaseModel):
    """⚙️ 配置 sheet 的完整解析结果。"""
    factors: list[FactorSpec]
    outputs: list[OutputSpec]
    strategy: StrategySpec
    constraints: list[dict[str, Any]] = Field(default_factory=list)


class RecordRow(BaseModel):
    """🧪 记录 一行(经校验)。"""
    trial_id: str
    batch_id: str | None
    factors: dict[str, Any]  # {factor_key: value}
    output: dict[str, float | None] = Field(default_factory=dict)  # {output_key: value}
    valid: int = 1  # 0=失败
    status: str = "measured"  # measured|pending|failed
    source: str | None = None  # 用户预填 / 第N轮推荐
    created_at: str | None = None


# ---------- 读 Excel ----------
def _col_idx(ws: Worksheet, header_value: str) -> int:
    """在表头行(默认 row=1, 但要兼容表格表头在 row=5 等情况)找 header 的列号。

    简单做法: 扫 row=1..10 找到第一个含 header_value 的 cell。
    """
    for r in range(1, min(15, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip().lower() == header_value.lower():
                return c
    raise KeyError(f"找不到表头 {header_value!r}")


# ⚙️ 配置 sheet 里三个区域的标题标记(第一个词是 emoji,稳定不变)
# ⚠️ 之前用**硬编码行号**(因子 r5/输出 r13/策略 r20)定位,
#    同事一改配置(加因子/插行)后面两个区就错位 ——
#    要么报错("Outputs 表为空"),要么**静默读错**
#    (因子表的 type 列读到别的区的值 → discrete 被当 continuous)。
#    改成按标题文字定位,行数随便变。
_SECTION_MARKERS = {
    "factors": "实验因子",
    "outputs": "实验目标",
    "strategy": "BO 策略",
    "constraints": "约束",
}

# 兜底:模板默认的行号(标题找不到时才用)
_FALLBACK_ROWS = {
    "factors": (5, 6),
    "outputs": (13, 14),
    "strategy": (20, 21),
    "constraints": (32, 33),
}


def _find_section(ws: Worksheet, section: str) -> tuple[int, int]:
    """找到某个区域的 (header_row, start_data_row)。

    策略:扫 A 列找标题(标题里有标记文字),标题的**下一行**是表头,
    表头的**下一行**是数据起点。找不到就用模板默认行号。
    """
    marker = _SECTION_MARKERS[section]
    max_r = min(ws.max_row, 80)

    for r in range(1, max_r + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        text = str(v)
        if marker in text:
            header_row = r + 1
            start = r + 2
            # 表头行确认(有 "key" 之类才认)
            hv = ws.cell(row=header_row, column=1).value
            if hv is None or str(hv).strip().lower() != "key":
                # 标题和数据之间隔了空行 → 往下找 "key" 行
                for rr in range(header_row, min(header_row + 5, max_r + 1)):
                    kv = ws.cell(row=rr, column=1).value
                    if kv is not None and str(kv).strip().lower() == "key":
                        return rr, rr + 1
            return header_row, start

    return _FALLBACK_ROWS[section]


def _read_kv(
    ws: Worksheet,
    header_row: int,
    start_data_row: int,
    stop_at_blank: bool = True,
    stop_markers: tuple[str, ...] = (),
) -> list[dict[str, Any]]:
    """读 key-value 型 sheet:从 header_row 读表头,从 start_data_row 起读行。

    返回 list[dict],每个 dict 是 1 行,key 是表头值。

    参数:
      stop_at_blank: True (默认) — 遇到完全空行就停止(避免读到下一个 region 的标题)
      stop_markers: 遇到 A 列含这些文字的行就停(下一个区域的标题)。
        ⚠️ 比 stop_at_blank 更可靠 —— 同事在因子区中间留空行时,
        因子不该被截断,但下一个区的标题要被挡住。
    """
    headers: list[str] = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            break
        headers.append(str(v).strip())
    rows = []
    for r in range(start_data_row, ws.max_row + 1):
        # ⚠️ 下一个区域的标题 → 停
        if stop_markers:
            a_val = ws.cell(row=r, column=1).value
            if a_val is not None:
                a_text = str(a_val)
                if any(m in a_text for m in stop_markers):
                    break
        row = {}
        all_blank = True
        for i, h in enumerate(headers):
            v = ws.cell(row=r, column=i + 1).value
            if v is not None and str(v).strip() != "":
                all_blank = False
            row[h] = v
        if all_blank:
            if stop_at_blank:
                break  # 空行 = 区域结束
            continue
        rows.append(row)
    return rows


def _parse_categories(raw: Any) -> list[str]:
    """把 '甲苯;THF;DMF' 解析成 ['甲苯','THF','DMF']。"""
    if raw is None:
        return []
    s = str(raw).strip()
    if not s:
        return []
    parts = [p.strip() for p in s.split(";") if p.strip()]
    return parts


def read_config(xlsx_path: Path) -> ConfigSheet:
    """读 ⚙️ 配置 sheet,返回 ConfigSheet。校验失败抛 ValueError(友好)。"""
    wb = load_workbook(str(xlsx_path), read_only=True, data_only=True)
    try:
        if CONFIG_SHEET not in wb.sheetnames:
            raise ValueError(f"xlsx 缺少 sheet: {CONFIG_SHEET!r}")

        ws = wb[CONFIG_SHEET]
        # ⚠️ 各区域按**标题文字**定位,不用硬编码行号 ——
        #    同事加因子/插行都不会错位(见 _find_section 注释)
        f_hdr, f_start = _find_section(ws, "factors")
        o_hdr, o_start = _find_section(ws, "outputs")
        s_hdr, s_start = _find_section(ws, "strategy")

        factor_rows = _read_kv(
            ws,
            header_row=f_hdr,
            start_data_row=f_start,
            stop_at_blank=False,  # ⚠️ 因子区中间有空行也不截断
            stop_markers=(
                _SECTION_MARKERS["outputs"],
                _SECTION_MARKERS["strategy"],
                _SECTION_MARKERS["constraints"],
            ),
        )
        output_rows = _read_kv(
            ws,
            header_row=o_hdr,
            start_data_row=o_start,
            stop_at_blank=False,
            stop_markers=(
                _SECTION_MARKERS["strategy"],
                _SECTION_MARKERS["constraints"],
            ),
        )
        strategy_rows = _read_kv(
            ws,
            header_row=s_hdr,
            start_data_row=s_start,
            stop_at_blank=False,
            stop_markers=(_SECTION_MARKERS["constraints"],),
        )

        factors: list[FactorSpec] = []
        for row in factor_rows:
            # 完全空行(None key)跳过
            if not row.get("key") or not str(row["key"]).strip():
                continue
            try:
                factors.append(
                    FactorSpec(
                        key=row["key"],
                        type=row["type"],
                        trainable=_as_bool(row.get("trainable")),
                        low=_as_float(row.get("low")),
                        high=_as_float(row.get("high")),
                        categories=_parse_categories(row.get("categories")) or None,
                        unit=row.get("unit"),
                        notes=row.get("notes"),
                    )
                )
            except Exception as e:
                raise ValueError(f"⚙️ 配置 - 因子 {row.get('key')!r}: {e}") from e

        outputs: list[OutputSpec] = []
        for row in output_rows:
            if not row.get("key") or not str(row["key"]).strip():
                continue
            try:
                outputs.append(
                    OutputSpec(
                        key=row["key"],
                        direction=row["direction"],
                        target=_as_float(row.get("target")),
                        lower_bound=_as_float(row.get("lower_bound")),
                        upper_bound=_as_float(row.get("upper_bound")),
                        unit=row.get("unit"),
                    )
                )
            except Exception as e:
                raise ValueError(f"⚙️ 配置 - 目标 {row.get('key')!r}: {e}") from e

        # strategy 是 key-value 单列
        strategy_kv = {r["key"]: r.get("value") for r in strategy_rows if r.get("key")}
        try:
            strategy = StrategySpec(
                strategy_type=str(strategy_kv.get("策略类型") or "sobo"),
                acq_function=str(strategy_kv.get("采集函数") or "qLogEI"),
                n_initial=int(strategy_kv.get("最少初始点数") or 5),
                batch_size=int(strategy_kv.get("批大小") or 3),
                convergence_threshold=float(strategy_kv.get("收敛阈值") or 0.01),
                convergence_window=int(strategy_kv.get("收敛窗口") or 5),
                save_model=_as_bool(strategy_kv.get("存模型")),
            )
        except Exception as e:
            raise ValueError(f"⚙️ 配置 - 策略: {e}") from e

        if not factors:
            raise ValueError("⚙️ 配置 - Factors 表为空,请至少填 1 个因子")
        if not outputs:
            raise ValueError("⚙️ 配置 - Outputs 表为空,请至少填 1 个目标")

        return ConfigSheet(
            factors=factors,
            outputs=outputs,
            strategy=strategy,
            constraints=[],  # TODO: 后续版本支持
        )
    finally:
        wb.close()


def read_records(
    xlsx_path: Path,
    factor_keys: list[str],
    output_keys: list[str],
    factor_specs: list[FactorSpec] | None = None,
) -> list[RecordRow]:
    """读 🧪 记录 sheet,只返回 valid=1 的行。

    ⚠️ 如果传了 factor_specs,会校验每个因子列的**值类型**是否匹配
    (continuous 要数值,categorical 要在 categories 列表里)。
    检测到错位会抛 ValueError,让用户知道"列填串了"而不是默默学到错乱数据。
    """
    wb = load_workbook(str(xlsx_path), read_only=True, data_only=True)
    try:
        if RECORD_SHEET not in wb.sheetnames:
            return []
        ws = wb[RECORD_SHEET]
        # 表头在 row=4
        headers: list[str] = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=4, column=c).value
            if v is None:
                break
            headers.append(str(v).strip())
        rows = []
        for r in range(5, ws.max_row + 1):
            row_raw = {}
            for i, h in enumerate(headers):
                row_raw[h] = ws.cell(row=r, column=i + 1).value
            trial_id = row_raw.get("trial_id")
            if not trial_id or not str(trial_id).strip():
                continue
            # 跳过提示行(以"💡"等 emoji 起头)
            if str(trial_id).startswith(("💡", "📌", "📎", "提示")):
                continue
            try:
                valid_raw = row_raw.get("valid")
                valid = int(valid_raw) if valid_raw is not None and str(valid_raw).strip() != "" else 1
            except (ValueError, TypeError):
                valid = 1
            # ⚠️ valid 列值异常(不是 0/1)→ 说明列填串了
            # (M4 现场 zzz12:base 列填了温度 90,导致 valid 列读到 90,
            #  数据被 int(90)!=1 静默过滤掉,BO 只用一半数据训练)
            if valid not in (0, 1):
                raise ValueError(
                    f"🧪 记录 sheet 第 {r} 行(trial_id={trial_id!r})的 `valid` 列值是 {valid_raw!r},"
                    f"只能是 0 或 1(或留空)。\n"
                    f"这通常说明**列填串位了** — 请检查该行的数据是否和 row 4 的表头一一对应。\n"
                    f"当前表头: {headers}\n"
                    f"该行数据: {[row_raw.get(h) for h in headers]}"
                )
            if valid != 1:
                continue
            row = RecordRow(
                trial_id=str(row_raw["trial_id"]),
                batch_id=str(row_raw["batch_id"]) if row_raw.get("batch_id") else None,
                factors={k: row_raw.get(k) for k in factor_keys if k in headers},
                output={k: _as_float(row_raw.get(k)) for k in output_keys if k in headers},
                valid=valid,
                status=str(row_raw.get("status") or "measured"),
                source=str(row_raw["来源"]) if row_raw.get("来源") else None,
                created_at=str(row_raw["录入时间"]) if row_raw.get("录入时间") else None,
            )
            rows.append(row)

        # ⚠️ 列值类型校验 — 检测"列填串了"(M4 现场 z13 / zzz12 都踩过)
        if factor_specs:
            _validate_factor_values(rows, factor_specs)
        return rows
    finally:
        wb.close()


def _validate_factor_values(
    rows: list[RecordRow],
    factor_specs: list[FactorSpec],
) -> None:
    """校验因子列的值类型是否和配置匹配。

    continuous/discrete → 必须是数值(或者空)
    categorical → 必须在 categories 列表里(或者空)

    发现错位 → 抛 ValueError,消息说清楚哪一列不对。
    """
    specs = {s.key: s for s in factor_specs}
    problems: dict[str, list[str]] = {}

    for row in rows:
        for key, val in row.factors.items():
            spec = specs.get(key)
            if spec is None or val is None:
                continue
            if spec.type in ("continuous", "discrete"):
                # 必须是数值
                try:
                    float(val)
                except (ValueError, TypeError):
                    problems.setdefault(key, []).append(
                        f"{row.trial_id}: {val!r}(期望数值)"
                    )
            elif spec.type in ("categorical", "categorical_descriptor", "categorical_molecular"):
                # 必须在 categories 里
                cats = spec.categories or []
                if cats and str(val) not in [str(c) for c in cats]:
                    problems.setdefault(key, []).append(
                        f"{row.trial_id}: {val!r}(不在候选列表里)"
                    )

    if problems:
        lines = []
        for key, items in problems.items():
            sample = items[0]
            more = f" 等 {len(items)} 处" if len(items) > 1 else ""
            lines.append(f"  · {key}: {sample}{more}")
        raise ValueError(
            "🧪 记录 sheet 的数据列和 ⚙️ 配置 对不上!\n"
            + "\n".join(lines)
            + "\n\n"
            "请检查 🧪 记录 sheet:每列的表头(row 4)和该列的数据类型要一致。\n"
            "例如 categorical 列(base/ligand/solvent)应该填字符串选项,\n"
            "continuous 列(temperature/concentration)应该填数字。\n"
            "⚠️ 也可以点 [🚀 启动 BO] 让系统自动重建表头(数据会按列名迁移)。"
        )


def read_status(xlsx_path: Path) -> dict[str, Any]:
    """读 _status sheet,返回 dict。

    _status 的设计: 我们用 'B<row>' 这种引用,所以读时把 key 也按
    '<col><row>' 形式组合:
        row 1: A='key',  B='value'         → result['B1'] = 'value'
        row 2: A='', B='awaiting_config'  → result['B2'] = 'awaiting_config'
        row 3: A='', B='1/4'              → result['B3'] = '1/4'

    所以 key 形式是 '<col_letter><row_num>'(B2、B3、...),方便 read 与
    write_status 对齐。
    """
    wb = load_workbook(str(xlsx_path), read_only=True, data_only=True)
    try:
        if STATUS_SHEET not in wb.sheetnames:
            return {}
        ws = wb[STATUS_SHEET]
        result = {}
        from openpyxl.utils import get_column_letter
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(row=r, column=c).value
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                key = f"{get_column_letter(c)}{r}"
                # 第一列如果是 'key'/'value' 这种表头,跳过
                if r == 1 and str(v).strip().lower() in ("key", "value"):
                    continue
                result[key] = v
        return result
    finally:
        wb.close()


# ---------- 写 ----------
def write_status(xlsx_path: Path, status: dict[str, Any]) -> None:
    """写 _status sheet 的 B 列(单 key 在 A 列,value 在 B 列)。

    设计: _status sheet 用 'B<row>' 直接存值(我们之前在 start sheet 用
    公式引用 _status!B2 等)。这里把所有状态写到 row=2 起 的 B 列。

    写入方式: openpyxl load → 改 _status B 列 → save。⚠️ 这会保存整个 workbook,
    如果 Excel 开着会冲突。所以这条函数只用于后台触发(用户没开文件时)。
    """
    wb = load_workbook(str(xlsx_path))
    try:
        if STATUS_SHEET not in wb.sheetnames:
            ws = wb.create_sheet(STATUS_SHEET)
            ws.sheet_state = "hidden"
            # 初始化
            for k, v in [
                ("key", "value"),
                ("B1", "_status"),
                ("B2", "awaiting_config"),
                ("B3", "1/4"),
                ("B4", ""),
                ("B5", ""),
                ("B6", ""),
                ("B7", 0),
                ("B8", 0),
                ("B9", ""),
                ("B10", ""),
                ("B11", ""),
            ]:
                ws[k] = v
        else:
            ws = wb[STATUS_SHEET]

        # 写 B2..B11 的固定位置
        # B2 = state, B3 = step, B4 = next, B5 = project name, B6 = project dir,
        # B7 = round, B8 = n_data, B9 = last_bo_at, B10 = convergence msg, B11 = error
        if "state" in status:
            ws["B2"] = status["state"]
        if "step" in status:
            ws["B3"] = status["step"]
        if "next" in status:
            ws["B4"] = status["next"]
        if "project_name" in status:
            ws["B5"] = status["project_name"]
        if "project_dir" in status:
            ws["B6"] = status["project_dir"]
        if "round" in status:
            ws["B7"] = status["round"]
        if "n_data" in status:
            ws["B8"] = status["n_data"]
        if "last_bo_at" in status:
            ws["B9"] = status["last_bo_at"]
        if "convergence_msg" in status:
            ws["B10"] = status["convergence_msg"]
        if "error" in status:
            ws["B11"] = status["error"]
        wb.save(str(xlsx_path))
    finally:
        wb.close()


def file_md5(path: Path) -> str:
    """计算文件 md5(用于 watchdog 内容变化检测)。"""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


# ---------- helpers ----------
def _as_float(v: Any) -> float | None:
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _as_bool(v: Any) -> bool:
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    s = str(v).strip().upper()
    return s in ("TRUE", "1", "YES", "是", "T")