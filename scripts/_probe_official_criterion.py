"""临时:验证 bofire 0.5.0 官方 ObjectiveImprovementCriterion 在真实流程里生效。

场景:构造一个明显停滞的数据集(最后 5 个实验没有提升),
挂上官方 criterion → strategy.has_converged() 应返回 True。
"""
from __future__ import annotations

import pandas as pd

import bofire
from bofire.data_models.domain.api import Domain
from bofire.data_models.features.api import ContinuousInput, ContinuousOutput
from bofire.data_models.objectives.api import MaximizeObjective
from bofire.data_models.strategies.api import SoboStrategy
from bofire.data_models.strategies.convergence_criteria.api import (
    ObjectiveImprovementCriterion,
)
from bofire.strategies.api import SoboStrategy as SoboImpl

print(f"bofire {bofire.__version__}\n")

domain = Domain(
    inputs=[ContinuousInput(key="temp", bounds=(20, 100))],
    outputs=[ContinuousOutput(key="y", objective=MaximizeObjective())],
)

# 20 个实验:前 15 个爬到 90,后 5 个在 89-89.8 之间抖动(**不超过 90** → 明显停滞)
rows = []
for i in range(20):
    if i < 15:
        y = 50 + i * (40 / 14)  # i=14 → 90.0
    else:
        y = 89.0 + (i % 3) * 0.4  # 89.0 / 89.4 / 89.8,都 < 90
    rows.append({"temp": 40 + i * 2, "y": y})

df = pd.DataFrame(rows)
print("数据最后 6 行:")
print(df.tail(6).to_string(index=False))
print()

for label, min_impr, n_lookback in [
    ("宽松(阈值 1.0, 窗口 5)", 1.0, 5),
    ("严格(阈值 0.01, 窗口 5)", 0.01, 5),
]:
    criterion = ObjectiveImprovementCriterion(
        min_improvement=min_impr,
        n_lookback=n_lookback,
    )
    dm = SoboStrategy(
        domain=domain,
        convergence_criterion=criterion,
    )
    strat = SoboImpl(dm)
    strat.tell(df)

    converged = strat.has_converged()
    print(f"{label}: has_converged() = {converged}")

    if converged:
        cands = strat.ask(1)
        print(f"   ✅ 仍能 ask → {cands.to_dict('records')}")
    print()
