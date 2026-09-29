"""tests/test_domain_builder.py — domain_builder 单元测试。

依赖: bofire 装好才测 Domain 构造;无 bofire 测 fallback 路径 + pure_python_ask。
"""
from __future__ import annotations

from portal.core import domain_builder, strategy_runner, excel_io


def _make_config(factor_type="continuous", categories=None):
    """构造一个简单 ConfigSheet。"""
    factors = [
        excel_io.FactorSpec(
            key="温度",
            type=factor_type,
            low=20.0 if factor_type in ("continuous", "discrete") else None,
            high=100.0 if factor_type in ("continuous", "discrete") else None,
            categories=categories,
        ),
    ]
    outputs = [
        excel_io.OutputSpec(key="Yield", direction="maximize", lower_bound=0, upper_bound=100),
    ]
    strategy = excel_io.StrategySpec()
    return excel_io.ConfigSheet(factors=factors, outputs=outputs, strategy=strategy, constraints=[])


class TestBuildDomain:
    def test_returns_domain_or_none(self):
        """build_domain 在 bofire 装上时返回 Domain,缺时返回 None。"""
        cfg = _make_config()
        result = domain_builder.build_domain(cfg)
        if not domain_builder.HAS_BOFIRE:
            assert result is None
            return
        # bofire 装上了,验证 Domain 构造
        assert result is not None
        # ⚠️ BoFire 0.4.1: Domain.inputs/Outputs 都是 Inputs/Outputs 对象,
        # 不是 list;真正的 features 列表是 .inputs.features / .outputs.features
        assert hasattr(result.inputs, "features")
        assert hasattr(result.outputs, "features")
        assert len(result.inputs.features) == 1
        assert len(result.outputs.features) == 1
        # keys
        assert [f.key for f in result.inputs.features] == ["温度"]
        assert [f.key for f in result.outputs.features] == ["Yield"]


class TestResolveAcf:
    def test_sobo_qLogEI(self):
        assert domain_builder.resolve_acf_function("sobo", "qLogEI", 1) == "qLogEI"

    def test_sobo_unknown_to_qLogEI(self, recwarn):
        # 不在白名单的 acf 降级
        assert domain_builder.resolve_acf_function("sobo", "magic", 1) == "qLogEI"

    def test_mobo_too_few_outputs(self):
        import pytest
        with pytest.raises(ValueError, match="MOBO 需要至少 2 个"):
            domain_builder.resolve_acf_function("mobo", "qLogEHVI", 1)

    def test_mobo_unknown_to_qLogEHVI(self, recwarn):
        assert domain_builder.resolve_acf_function("mobo", "qLogEI", 2) == "qLogEHVI"


class TestOutputWeights:
    def test_maximize(self):
        weights = domain_builder.output_objective_weights(
            [excel_io.OutputSpec(key="Y", direction="maximize")]
        )
        assert weights == [1.0]

    def test_minimize(self):
        weights = domain_builder.output_objective_weights(
            [excel_io.OutputSpec(key="Y", direction="minimize")]
        )
        assert weights == [-1.0]

    def test_target(self):
        weights = domain_builder.output_objective_weights(
            [excel_io.OutputSpec(key="Y", direction="target", target=95.0)]
        )
        assert weights == [1.0]


class TestPrepareData:
    def test_empty(self):
        from dataclasses import dataclass
        cfg = _make_config()
        # 空 records
        @dataclass
        class R:
            trial_id: "T"
            factors: {}
            output: {}
            valid: 1
        prepared = domain_builder.prepare_data([], cfg)
        assert len(prepared.X_cols) == 1
        assert prepared.X_cols == ["温度"]


class TestPurePythonAsk:
    def test_returns_correct_count(self):
        cfg = _make_config(factor_type="continuous")
        cands = strategy_runner._pure_python_ask(cfg, 1)
        assert len(cands) == cfg.strategy.batch_size

    def test_continuous_factors_in_range(self):
        cfg = _make_config(factor_type="continuous")
        cands = strategy_runner._pure_python_ask(cfg, 1)
        for c in cands:
            v = c.factors["温度"]
            assert 20.0 <= v <= 100.0

    def test_categorical_factors_valid(self):
        cfg = _make_config(factor_type="categorical", categories=["甲苯", "THF", "DMF"])
        cands = strategy_runner._pure_python_ask(cfg, 1)
        for c in cands:
            assert c.factors["温度"] in ["甲苯", "THF", "DMF"]

    def test_trial_ids_sequential(self):
        cfg = _make_config(factor_type="continuous")
        cands = strategy_runner._pure_python_ask(cfg, 100)
        assert [c.trial_id for c in cands] == ["T0100", "T0101", "T0102"]


class TestPinNonTrainable:
    def test_pins_to_mode(self):
        cfg = _make_config(factor_type="continuous")
        # 加一个非 trainable 因子
        cfg.factors.append(
            excel_io.FactorSpec(
                key="催化剂",
                trainable=False,
                type="categorical",
                categories=["A", "B"],
            )
        )
        records = []
        # 构造 record
        for tid, cat in [("T1", "A"), ("T2", "A"), ("T3", "B")]:
            r = excel_io.RecordRow(
                trial_id=tid,
                batch_id=None,
                factors={"温度": 60, "催化剂": cat},
                output={"Yield": 70.0},
            )
            records.append(r)

        cands = strategy_runner.pin_non_trainable(
            [strategy_runner.Candidate("T100", {"温度": 65, "催化剂": "B"})],
            records,
            cfg,
        )
        assert cands[0].factors["催化剂"] == "A"  # mode from history