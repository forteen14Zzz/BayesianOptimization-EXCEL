"""portal/core/convergence.py — 收敛判定 + 收敛曲线数据。

⚠️ 2026-11:参考 BoFire 官方 Convergence Criteria 文档实现。

官方文档:
https://experimental-design.github.io/bofire/docs/userguides/convergence_criteria.html

官方 `ObjectiveImprovementCriterion` 的数学定义:

    每个实验经 domain objectives 映射成标量 reward(越大越好),
    r_1, ..., r_N 按时间顺序;b_k = max_{i<=k} r_i 即前 k 个实验里的最好成绩。
    当最好成绩在最近 n_lookback 个实验内的提升小于 min_improvement 时,判为收敛:

        b_N - b_{N - n_lookback} < min_improvement

    至少需要 n_lookback + 1 个实验,否则不判收敛。

我们的两个配置项与官方参数一一对应:

| ⚙️ 配置    | BoFire 官方            | 含义                        |
|-----------|----------------------|----------------------------|
| 收敛阈值   | `min_improvement`     | 最小有意义的提升量            |
| 收敛窗口   | `n_lookback`         | 看最近几个实验                |

⚠️ **单位是"实验数"不是"轮数"** — 官方注释明确写 "over the last n_lookback
experiments"。一轮 batch_size=3 的推荐 = 3 个实验。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ConvergenceResult:
    """收敛判定结果。"""

    converged: bool = False
    #: 当前最好成绩(已转成"越大越好"的标量 reward)
    best_so_far: float | None = None
    #: 窗口起点之前的最好成绩
    baseline: float | None = None
    #: 窗口内提升量
    improvement: float | None = None
    #: 有效实验数
    n_experiments: int = 0
    #: 人话说明(给 _status / 看板 用)
    message: str = "⏳ 暂无收敛数据"
    #: 收敛曲线数据点 [(round, best_so_far, mean_of_best3, n_measured), ...]
    curve: list[tuple] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "converged": self.converged,
            "best_so_far": self.best_so_far,
            "baseline": self.baseline,
            "improvement": self.improvement,
            "n_experiments": self.n_experiments,
            "message": self.message,
        }


def _records_to_rewards(records: list[Any], config: Any) -> list[float]:
    """把实验记录转成"越大越好"的标量 reward 序列。

    reward = objective(desirability),多目标时取各目标 desirability 之和
    (与官方 `ObjectiveImprovementCriterion` 一致:
     `desirabilities = strategy.domain.outputs(experiments)` → `sum(axis=1)`)。

    ⚠️ 我们的 record 里 invalid 行(valid=0)必须排除 — 它们没有有效目标值。
    """
    rewards: list[float] = []

    outputs = config.outputs
    if not outputs:
        return rewards

    for rec in records:
        # 跳过无效实验(valid=0)
        if not getattr(rec, "valid", True):
            continue

        obj_values = getattr(rec, "output", None) or {}
        vals: list[float] = []
        ok = True
        for o in outputs:
            v = obj_values.get(o.key)
            if v is None:
                ok = False
                break
            try:
                vals.append(float(v))
            except (TypeError, ValueError):
                ok = False
                break
        if not ok or not vals:
            continue

        # 按 objective direction 转成"越大越好"
        total = 0.0
        for o, v in zip(outputs, vals):
            direction = (o.direction or "maximize").lower()
            if direction == "minimize":
                total += -v
            elif direction == "target":
                # close-to-target:距离目标越近越好 → 用 -距离
                target = o.target if o.target is not None else 0.0
                total += -abs(v - float(target))
            else:  # maximize
                total += v
        rewards.append(total)

    return rewards


def evaluate_convergence(
    records: list[Any],
    config: Any,
) -> ConvergenceResult:
    """按官方公式判定是否收敛。

    Args:
        records: 实验记录列表(RecordRow),按时间顺序。
        config:  ConfigSheet(含 strategy.convergence_threshold / .convergence_window)。

    Returns:
        ConvergenceResult
    """
    threshold = float(getattr(config.strategy, "convergence_threshold", 0.01) or 0.01)
    window = int(getattr(config.strategy, "convergence_window", 5) or 5)

    rewards = _records_to_rewards(records, config)
    n = len(rewards)
    result = ConvergenceResult(n_experiments=n)

    if n == 0:
        result.message = "⏳ 暂无收敛数据(还没有有效实验)"
        return result

    # 收敛曲线:按轮次累计(每 batch_size 个实验算一轮)
    batch = max(1, int(getattr(config.strategy, "batch_size", 1) or 1))
    curve: list[tuple] = []
    best = None
    best3: list[float] = []
    for i, r in enumerate(rewards):
        best = r if best is None else max(best, r)
        best3.append(r)
        best3.sort(reverse=True)
        top3 = best3[:3]
        curve.append(
            (
                i // batch + 1,
                round(best, 6),
                round(sum(top3) / len(top3), 6),
                len(best3),
            )
        )
    result.curve = curve

    best_now = max(rewards)

    # ⚠️ 官方:至少需要 n_lookback + 1 个实验才能判定
    if n <= window:
        result.best_so_far = best_now
        result.message = (
            f"⏳ 收敛评估中 — 已积累 {n} 个实验,"
            f"需要 > {window} 个才能判定(收敛窗口={window} 个实验)"
        )
        return result

    best_before = max(rewards[: n - window])
    improvement = best_now - best_before

    result.best_so_far = best_now
    result.baseline = best_before
    result.improvement = improvement
    result.converged = bool(improvement < threshold)

    if result.converged:
        result.message = (
            f"✅ 已收敛 — 最近 {window} 个实验内最好成绩只提升了 "
            f"{improvement:.4f}(阈值 {threshold})"
        )
    else:
        result.message = (
            f"⏳ 未收敛 — 最近 {window} 个实验内最好成绩提升了 "
            f"{improvement:.4f}(阈值 {threshold})"
        )
    return result
