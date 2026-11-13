"""tests/test_config_layout.py — ⚙️ 配置 sheet 布局鲁棒性测试。

⚠️ 背景(2026-11 同事现场问题):
同事在配置里设了 discrete 因子,但 BO 推荐出来的还是连续值。

根因分析:
1. read_config 之前用**硬编码行号**定位三个区域
   (因子 r5/输出 r13/策略 r20)—— 同事一改配置(加因子/插行)
   后面两个区就错位:
   - 要么报错("Outputs 表为空")
   - 要么**静默读错**(因子表的 type 列读到别的区的值
     → discrete 被当成 continuous)
2. discrete 因子**没有校验 categories** —— 只填 low/high 时
   BoFire 拿不到候选值,只能当连续变量处理

这些测试锁住修复后的行为。
"""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from portal.core.excel_io import (
    FactorSpec,
    read_config,
)


TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "project_template.xlsx"


# ---------- helpers ----------
def _write_config_sections(
    ws,
    factors: list[dict],
    outputs: list[dict] | None = None,
    strategy: list[dict] | None = None,
    factor_start_row: int = 6,
    factor_gap: int = 2,
    output_gap: int = 3,
) -> None:
    """按"标题 → 空行 → 表头 → 数据"的布局写一个配置 sheet。

    ⚠️ 行号全部由参数决定,用来模拟"同事改了配置行数"。
    """
    outputs = outputs or [{"key": "Yield", "direction": "maximize"}]
    strategy = strategy or [{"key": "策略类型", "value": "sobo"}]

    # --- 因子区 ---
    f_title = 4
    f_hdr = f_title + 1
    ws["A4"] = "📐 实验因子 (Factors)"
    for c, h in enumerate(["key", "type", "trainable", "low", "high", "categories"], 1):
        ws.cell(row=f_hdr, column=c, value=h)
    for i, f in enumerate(factors):
        r = factor_start_row + i
        ws.cell(row=r, column=1, value=f["key"])
        ws.cell(row=r, column=2, value=f["type"])
        ws.cell(row=r, column=3, value=f.get("trainable", "TRUE"))
        if f.get("low") is not None:
            ws.cell(row=r, column=4, value=f["low"])
        if f.get("high") is not None:
            ws.cell(row=r, column=5, value=f["high"])
        if f.get("categories"):
            ws.cell(row=r, column=6, value=f["categories"])

    # --- 输出区(位置随因子数变化) ---
    o_title = factor_start_row + len(factors) + factor_gap
    o_hdr = o_title + 1
    ws.cell(row=o_title, column=1, value="🎯 实验目标 (Outputs)")
    ws.cell(row=o_hdr, column=1, value="key")
    ws.cell(row=o_hdr, column=2, value="direction")
    for i, o in enumerate(outputs):
        r = o_hdr + 1 + i
        ws.cell(row=r, column=1, value=o["key"])
        ws.cell(row=r, column=2, value=o["direction"])

    # --- 策略区 ---
    s_title = o_hdr + len(outputs) + 1 + output_gap
    s_hdr = s_title + 1
    ws.cell(row=s_title, column=1, value="🧠 BO 策略 (Strategy)")
    ws.cell(row=s_hdr, column=1, value="key")
    ws.cell(row=s_hdr, column=2, value="value")
    for i, s in enumerate(strategy):
        r = s_hdr + 1 + i
        ws.cell(row=r, column=1, value=s["key"])
        ws.cell(row=r, column=2, value=s["value"])

    # --- 约束区 ---
    c_title = s_hdr + len(strategy) + 3
    ws.cell(row=c_title, column=1, value="🔒 约束 (Constraints) — 一般不填")


def _make_xlsx(tmp_path: Path, name: str, **kw) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "⚙️ 配置"
    _write_config_sections(ws, **kw)
    p = tmp_path / name
    wb.save(str(p))
    return p


DISCRETE_FACTORS = [
    {"key": "温度", "type": "continuous", "low": 20, "high": 100},
    {"key": "催化剂用量", "type": "discrete", "categories": "0.5;1.0;2.0;5.0"},
    {"key": "溶剂", "type": "categorical", "categories": "甲苯;THF;DMF"},
]


# ---------- 测试 ----------
class TestConfigLayoutRobust:
    """配置行数变化时,三个区域仍要正确读取。"""

    def test_standard_layout(self, tmp_path: Path):
        """模板默认布局。"""
        p = _make_xlsx(tmp_path, "std.xlsx", factors=DISCRETE_FACTORS)
        cfg = read_config(p)
        assert [f.key for f in cfg.factors] == ["温度", "催化剂用量", "溶剂"]
        assert cfg.factors[1].type == "discrete"
        assert cfg.factors[1].categories == ["0.5", "1.0", "2.0", "5.0"]
        assert [o.key for o in cfg.outputs] == ["Yield"]

    def test_many_factors_shift_sections_down(self, tmp_path: Path):
        """⚠️ 因子加到 6 个 → 输出/策略区整体下移,仍要读对。

        这是同事遇到的场景 —— 之前用硬编码行号,这里直接报
        "Outputs 表为空"。
        """
        factors = [
            {"key": "温度", "type": "continuous", "low": 20, "high": 100},
            {"key": "压强", "type": "continuous", "low": 1, "high": 10},
            {"key": "催化剂用量", "type": "discrete", "categories": "0.5;1.0;2.0;5.0"},
            {"key": "时间", "type": "continuous", "low": 1, "high": 24},
            {"key": "浓度", "type": "continuous", "low": 0.1, "high": 1},
            {"key": "溶剂", "type": "categorical", "categories": "甲苯;THF;DMF"},
        ]
        p = _make_xlsx(tmp_path, "many.xlsx", factors=factors)
        cfg = read_config(p)
        assert len(cfg.factors) == 6
        assert cfg.factors[2].type == "discrete"
        assert [o.key for o in cfg.outputs] == ["Yield"]
        assert cfg.strategy.strategy_type == "sobo"

    def test_blank_row_inside_factors(self, tmp_path: Path):
        """⚠️ 因子区中间有空行 → 后面的因子不能被截断。

        之前 _read_kv 用 stop_at_blank=True,空行就停。
        """
        factors = DISCRETE_FACTORS + [
            {"key": "压强", "type": "continuous", "low": 1, "high": 10},
        ]
        # 在第 2 个因子后插空行
        p = _make_xlsx(tmp_path, "blank.xlsx", factors=factors)
        wb = load_workbook(str(p))
        ws = wb["⚙️ 配置"]
        ws.insert_rows(8)  # 在 r8 前插一行 → r8 空
        wb.save(str(p))
        wb.close()

        cfg = read_config(p)
        keys = [f.key for f in cfg.factors]
        assert "催化剂用量" in keys, f"空行不该截断因子,实际读到 {keys}"
        assert "压强" in keys, f"空行后的因子也要读到,实际读到 {keys}"

    def test_factors_start_lower(self, tmp_path: Path):
        """因子从 r8 开始写(前面有空行)→ 仍要读到。"""
        p = _make_xlsx(tmp_path, "lower.xlsx", factors=DISCRETE_FACTORS, factor_start_row=8)
        cfg = read_config(p)
        assert len(cfg.factors) == 3
        assert cfg.factors[1].type == "discrete"

    def test_multiple_outputs(self, tmp_path: Path):
        """多个目标 → 输出区要读完。"""
        p = _make_xlsx(
            tmp_path,
            "multi_out.xlsx",
            factors=DISCRETE_FACTORS,
            outputs=[
                {"key": "Yield", "direction": "maximize"},
                {"key": "cost", "direction": "minimize"},
            ],
        )
        cfg = read_config(p)
        assert [o.key for o in cfg.outputs] == ["Yield", "cost"]


class TestDiscreteValidation:
    """discrete 必须填 categories,否则会被当连续变量。"""

    def test_discrete_without_categories_raises(self):
        """⚠️ 只填 low/high → 明确报错,而不是静默当连续变量。"""
        with pytest.raises(ValueError, match="discrete 必须填 categories"):
            FactorSpec(key="用量", type="discrete", low=0.5, high=5.0)

    def test_discrete_with_categories_ok(self):
        f = FactorSpec(key="用量", type="discrete", categories=["0.5", "1.0", "2.0"])
        assert f.categories == ["0.5", "1.0", "2.0"]

    def test_discrete_non_numeric_raises(self):
        """非数字候选值 → 提示改用 categorical。"""
        with pytest.raises(ValueError, match="必须是数字"):
            FactorSpec(key="溶剂", type="discrete", categories=["甲苯", "THF"])

    def test_discrete_too_few_candidates(self):
        with pytest.raises(ValueError, match="至少 2 个候选值"):
            FactorSpec(key="用量", type="discrete", categories=["1.0"])


class TestRealTemplate:
    """真实模板也要能读(回归保护)。"""

    def test_template_reads(self, tmp_path: Path):
        dst = tmp_path / "tpl.xlsx"
        shutil.copy2(TEMPLATE, dst)
        cfg = read_config(dst)
        assert len(cfg.factors) >= 1
        assert len(cfg.outputs) >= 1
        assert cfg.strategy.strategy_type
