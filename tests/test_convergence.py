"""tests/test_convergence.py — 收敛判定(参考 BoFire 官方公式)。

官方文档:
https://experimental-design.github.io/bofire/docs/userguides/convergence_criteria.html

公式: b_N - b_{N-n_lookback} < min_improvement → 收敛
"""
from __future__ import annotations

import pytest

from portal.core.convergence import evaluate_convergence
from portal.core.excel_io import (
    ConfigSheet,
    FactorSpec,
    OutputSpec,
    RecordRow,
    StrategySpec,
)


def _cfg(threshold=0.01, window=5, batch=1) -> ConfigSheet:
    return ConfigSheet(
        factors=[FactorSpec(key="temp", type="continuous", trainable=True, low=0, high=100)],
        outputs=[OutputSpec(key="Yield", direction="maximize")],
        strategy=StrategySpec(
            strategy_type="sobo",
            acq_function="qLogEI",
            n_initial=3,
            batch_size=batch,
            convergence_threshold=threshold,
            convergence_window=window,
        ),
    )


def _rec(yield_val, valid=1) -> RecordRow:
    return RecordRow(
        trial_id=f"T{yield_val}",
        batch_id=None,
        factors={"temp": 50},
        output={"Yield": yield_val},
        valid=valid,
    )


# ---------- 基本行为 ----------

def test_no_records():
    r = evaluate_convergence([], _cfg())
    assert r.converged is False
    assert r.n_experiments == 0
    assert "暂无" in r.message


def test_not_enough_experiments():
    """官方:至少需要 n_lookback + 1 个实验才能判定。"""
    recs = [_rec(80), _rec(82), _rec(90)]  # 只有 3 个,窗口=5
    r = evaluate_convergence(recs, _cfg(window=5))
    assert r.converged is False
    assert r.n_experiments == 3
    assert r.improvement is None  # 还没算
    assert "需要" in r.message


def test_converged_when_stagnant():
    """窗口内最好成绩提升 < 阈值 → 收敛。"""
    recs = [_rec(80), _rec(81), _rec(82), _rec(83), _rec(84), _rec(84.5)]
    r = evaluate_convergence(recs, _cfg(threshold=1.0, window=3))
    # best_now=84.5, best_before=max(80,81,82)=82 → improvement=2.5 > 1.0 → 未收敛
    assert r.converged is False
    assert r.improvement == pytest.approx(2.5)


def test_converged_true():
    """窗口内提升很小 → 收敛。"""
    recs = [_rec(90), _rec(90.2), _rec(90.1), _rec(90.3)]
    r = evaluate_convergence(recs, _cfg(threshold=1.0, window=3))
    # best_now=90.3, best_before=max(90)=90 → improvement=0.3 < 1.0 → 收敛
    assert r.converged is True
    assert r.improvement == pytest.approx(0.3)
    assert "已收敛" in r.message


def test_invalid_records_excluded():
    """valid=0 的实验不进 reward 序列。"""
    recs = [_rec(80), _rec(999, valid=0), _rec(81), _rec(82), _rec(83)]
    r = evaluate_convergence(recs, _cfg(threshold=1.0, window=3))
    assert r.n_experiments == 4  # 999 那条被排除


# ---------- direction 处理 ----------

def test_minimize_direction():
    """minimize 目标:reward 取负 → 数值越小越好。"""
    cfg = ConfigSheet(
        factors=[FactorSpec(key="t", type="continuous", trainable=True, low=0, high=100)],
        outputs=[OutputSpec(key="cost", direction="minimize")],
        strategy=StrategySpec(convergence_threshold=1.0, convergence_window=3),
    )
    # cost 持续下降 → 已达最优 → 提升(负向)很小 → 收敛
    recs = [
        RecordRow(trial_id="a", batch_id=None, factors={}, output={"cost": 10.0}),
        RecordRow(trial_id="b", batch_id=None, factors={}, output={"cost": 9.9}),
        RecordRow(trial_id="c", batch_id=None, factors={}, output={"cost": 9.85}),
        RecordRow(trial_id="d", batch_id=None, factors={}, output={"cost": 9.84}),
    ]
    r = evaluate_convergence(recs, cfg)
    # reward = -10, -9.9, -9.85, -9.84
    # 官方公式 b_N - b_{N-n_lookback}:n=4, lookback=3 → baseline = max(rewards[:1]) = -10
    # improvement = -9.84 - (-10) = 0.16 < 1.0 → 收敛
    assert r.converged is True
    assert r.improvement == pytest.approx(0.16)


def test_target_direction():
    """target 方向:距离目标越近越好。"""
    cfg = ConfigSheet(
        factors=[FactorSpec(key="t", type="continuous", trainable=True, low=0, high=100)],
        outputs=[OutputSpec(key="pH", direction="target", target=7.0)],
        strategy=StrategySpec(convergence_threshold=1.0, convergence_window=3),
    )
    recs = [
        RecordRow(trial_id="a", batch_id=None, factors={}, output={"pH": 6.0}),
        RecordRow(trial_id="b", batch_id=None, factors={}, output={"pH": 6.9}),
        RecordRow(trial_id="c", batch_id=None, factors={}, output={"pH": 6.95}),
        RecordRow(trial_id="d", batch_id=None, factors={}, output={"pH": 6.96}),
    ]
    r = evaluate_convergence(recs, cfg)
    # reward = -|x-7| = -1.0, -0.1, -0.05 → best_now=-0.05, best_before=-0.1
    # improvement = 0.05 < 1.0 → 收敛
    assert r.converged is True


# ---------- 收敛曲线 ----------

def test_curve_generated():
    """曲线每个有效实验一个点。"""
    recs = [_rec(10), _rec(20), _rec(30)]
    r = evaluate_convergence(recs, _cfg(batch=1))
    assert len(r.curve) == 3
    # best_so_far 单调不减
    bests = [p[1] for p in r.curve]
    assert bests == sorted(bests)
    assert bests[-1] == 30


def test_curve_batch_rounds():
    """batch_size=2 → 每 2 个实验算一轮。"""
    recs = [_rec(10), _rec(20), _rec(30), _rec(40)]
    r = evaluate_convergence(recs, _cfg(batch=2))
    rounds = [p[0] for p in r.curve]
    assert rounds == [1, 1, 2, 2]


def test_mean_of_best3():
    """mean_of_best3 = 最近 top3 的均值。"""
    recs = [_rec(10), _rec(20), _rec(30)]
    r = evaluate_convergence(recs, _cfg())
    last = r.curve[-1]
    assert last[1] == 30  # best_so_far
    assert last[2] == pytest.approx((30 + 20 + 10) / 3)  # mean_of_best3
    assert last[3] == 3  # n_measured
