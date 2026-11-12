"""tests/test_excel_io.py — excel_io 单元测试"""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from portal.core.excel_io import (
    FactorSpec,
    OutputSpec,
    StrategySpec,
    ConfigSheet,
    read_config,
    read_records,
    read_status,
    write_status,
    file_md5,
    _parse_categories,
    _as_bool,
    _as_float,
)


@pytest.fixture
def template_xlsx(tmp_path: Path) -> Path:
    """复制 templates/project_template.xlsx 到 tmp,作为测试输入。"""
    src = Path(__file__).resolve().parent.parent / "templates" / "project_template.xlsx"
    dst = tmp_path / "test_project.xlsx"
    shutil.copy2(src, dst)
    return dst


class TestHelpers:
    def test_parse_categories(self):
        assert _parse_categories("甲苯;THF;DMF") == ["甲苯", "THF", "DMF"]
        assert _parse_categories("甲苯") == ["甲苯"]
        assert _parse_categories("") == []
        assert _parse_categories(None) == []
        assert _parse_categories("a;b;;c") == ["a", "b", "c"]  # 中间空忽略,末尾空忽略

    def test_as_bool(self):
        assert _as_bool("TRUE") is True
        assert _as_bool("1") is True
        assert _as_bool("YES") is True
        assert _as_bool("是") is True
        assert _as_bool("FALSE") is False
        assert _as_bool("0") is False
        assert _as_bool(None) is False
        assert _as_bool(True) is True
        assert _as_bool(False) is False

    def test_as_float(self):
        assert _as_float("3.14") == 3.14
        assert _as_float(3) == 3.0
        assert _as_float("") is None
        assert _as_float(None) is None
        assert _as_float("abc") is None


class TestFactorSpec:
    def test_continuous_ok(self):
        f = FactorSpec(key="温度", type="continuous", low=0, high=100)
        assert f.trainable is True

    def test_continuous_missing_range(self):
        with pytest.raises(ValueError, match="必须填 low 和 high"):
            FactorSpec(key="温度", type="continuous")

    def test_continuous_low_ge_high(self):
        with pytest.raises(ValueError, match="low 必须 < high"):
            FactorSpec(key="温度", type="continuous", low=100, high=0)

    def test_categorical_ok(self):
        f = FactorSpec(key="溶剂", type="categorical", categories=["甲苯", "THF"])
        assert f.low is None
        assert f.categories == ["甲苯", "THF"]

    def test_categorical_too_few(self):
        with pytest.raises(ValueError, match="至少 2 个 category"):
            FactorSpec(key="溶剂", type="categorical", categories=["甲苯"])

    def test_invalid_type(self):
        with pytest.raises(ValueError, match="type 必须是"):
            FactorSpec(key="x", type="magic")


class TestOutputSpec:
    def test_maximize_ok(self):
        o = OutputSpec(key="Yield", direction="maximize", lower_bound=0, upper_bound=100)
        assert o.target is None

    def test_target_needs_value(self):
        with pytest.raises(ValueError, match="direction=target 必须填 target"):
            OutputSpec(key="选择性", direction="target")

    def test_invalid_direction(self):
        with pytest.raises(ValueError, match="direction 必须是"):
            OutputSpec(key="x", direction="up")


class TestStrategySpec:
    def test_defaults(self):
        s = StrategySpec()
        assert s.strategy_type == "sobo"
        assert s.batch_size == 3
        assert s.convergence_threshold == 0.01

    def test_batch_too_large(self):
        with pytest.raises(ValueError):
            StrategySpec(batch_size=100)


class TestReadConfig:
    def test_read_template(self, template_xlsx: Path):
        cfg = read_config(template_xlsx)
        assert isinstance(cfg, ConfigSheet)
        assert len(cfg.factors) >= 3  # 模板预填了 4 个
        # 温度
        temp = next(f for f in cfg.factors if f.key == "温度")
        assert temp.type == "continuous"
        assert temp.low == 20
        assert temp.high == 100
        # 溶剂
        sol = next(f for f in cfg.factors if f.key == "溶剂")
        assert sol.type == "categorical"
        assert sol.categories is not None
        assert "甲苯" in sol.categories
        # Yield
        y = cfg.outputs[0]
        assert y.key == "Yield"
        assert y.direction == "maximize"
        # strategy
        assert cfg.strategy.strategy_type == "sobo"
        assert cfg.strategy.batch_size == 3

    def test_missing_sheet(self, tmp_path: Path):
        from openpyxl import Workbook
        wb = Workbook()
        wb.save(str(tmp_path / "empty.xlsx"))
        with pytest.raises(ValueError, match="缺少 sheet"):
            read_config(tmp_path / "empty.xlsx")


class TestReadRecords:
    def test_empty(self, template_xlsx: Path):
        rows = read_records(template_xlsx, factor_keys=["温度"], output_keys=["Yield"])
        assert rows == []

    def test_with_data(self, template_xlsx: Path):
        from openpyxl import load_workbook
        wb = load_workbook(str(template_xlsx))
        ws = wb["🧪 记录"]
        # 写 3 行
        data = [
            ("T0001", "2026-09-29", 65, 2.0, 12, "甲苯", 78.5, 1, "measured", "第1轮推荐", "2026-09-29 14:30"),
            ("T0002", "2026-09-29", 70, 1.5, 8, "THF", 82.1, 1, "measured", "第1轮推荐", "2026-09-29 14:30"),
            ("T0003", "2026-09-29", 75, 1.8, 10, "DMF", 75.0, 0, "failed", "第1轮推荐", "2026-09-29 14:30"),
        ]
        for r_idx, row in enumerate(data, start=5):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=val)
        wb.save(str(template_xlsx))
        wb.close()

        parsed = read_records(
            template_xlsx,
            factor_keys=["温度", "催化剂用量", "反应时间", "溶剂"],
            output_keys=["Yield"],
        )
        # 第 3 行 valid=0 应被过滤
        assert len(parsed) == 2
        assert parsed[0].trial_id == "T0001"
        assert parsed[0].factors["温度"] == 65
        assert parsed[0].output["Yield"] == 78.5


class TestStatus:
    def test_read_status(self, template_xlsx: Path):
        s = read_status(template_xlsx)
        assert s.get("B2") == "awaiting_config"

    def test_write_status(self, template_xlsx: Path):
        write_status(
            template_xlsx,
            {
                "state": "awaiting_bo",
                "step": "1/4",
                "next": "点 [🚀 启动 BO]",
                "project_name": "Suzuki_2026Q4_demo",
                "project_dir": "projects/Suzuki_2026Q4_demo",
                "round": 3,
                "n_data": 9,
                "last_bo_at": "2026-09-29 14:30",
                "convergence_msg": "📈 5 轮无显著提升",
                "error": "",
            },
        )
        s = read_status(template_xlsx)
        assert s["B2"] == "awaiting_bo"
        assert s["B5"] == "Suzuki_2026Q4_demo"
        assert s["B7"] == 3
        assert s["B8"] == 9


class TestFileMd5:
    def test_md5(self, template_xlsx: Path):
        h1 = file_md5(template_xlsx)
        h2 = file_md5(template_xlsx)
        assert h1 == h2
        assert len(h1) == 32