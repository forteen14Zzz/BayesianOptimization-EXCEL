"""临时:测 bofire 0.5.0 下我们代码用到的 API 是否还兼容。

用法:
  <venv>/python.exe scripts/_probe_compat.py
"""
from __future__ import annotations

import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

results: list[tuple[str, bool, str]] = []


def check(name, fn):
    try:
        r = fn()
        results.append((name, True, str(r)[:200]))
    except Exception as e:
        results.append((name, False, f"{type(e).__name__}: {e}"))
        traceback.print_exc(limit=1)


def main():
    import bofire

    print(f"bofire {bofire.__version__}\n")

    # --- 1. convergence criteria 模块(0.5.0 新增) ---
    def conv_api():
        from bofire.data_models.strategies.convergence_criteria.api import (
            ObjectiveImprovementCriterion,
        )

        c = ObjectiveImprovementCriterion(min_improvement=0.01, n_lookback=5)
        return f"OK: min_improvement={c.min_improvement}, n_lookback={c.n_lookback}"

    check("convergence_criteria.api.ObjectiveImprovementCriterion", conv_api)

    def conv_has_converged():
        from bofire.strategies.api import SoboStrategy

        assert hasattr(SoboStrategy, "has_converged"), "no has_converged method"
        return "has_converged() 存在"

    check("strategy.has_converged()", conv_has_converged)

    def conv_strategy_field():
        from bofire.data_models.domain.api import Domain
        from bofire.data_models.features.api import ContinuousInput, ContinuousOutput
        from bofire.data_models.objectives.api import MaximizeObjective
        from bofire.data_models.strategies.api import SoboStrategy
        from bofire.data_models.strategies.convergence_criteria.api import (
            ObjectiveImprovementCriterion,
        )

        domain = Domain(
            inputs=[ContinuousInput(key="x", bounds=(0.0, 1.0))],
            outputs=[ContinuousOutput(key="y", objective=MaximizeObjective())],
        )
        dm = SoboStrategy(
            domain=domain,
            convergence_criterion=ObjectiveImprovementCriterion(
                min_improvement=0.01, n_lookback=5
            ),
        )
        return f"convergence_criterion = {dm.convergence_criterion}"

    check("SoboStrategy.convergence_criterion field", conv_strategy_field)

    # --- 2. 我们代码用到的核心 API ---
    def features_api():
        from bofire.data_models.features.api import (
            CategoricalInput,
            ContinuousInput,
            ContinuousOutput,
        )

        return "ContinuousInput/ContinuousOutput/CategoricalInput OK"

    check("features.api", features_api)

    def objectives_api():
        from bofire.data_models.objectives.api import (
            MaximizeObjective,
            MinimizeObjective,
            TargetObjective,
        )

        return "Maximize/Minimize/Target OK"

    check("objectives.api", objectives_api)

    def domain_api():
        from bofire.data_models.domain.api import Domain

        return "Domain OK"

    check("domain.api", domain_api)

    def strategies_api():
        from bofire.strategies.api import RandomStrategy, SoboStrategy

        return "SoboStrategy/RandomStrategy OK"

    check("strategies.api", strategies_api)

    def acq_api():
        from bofire.data_models.acquisition_functions.api import (
            qEI,
            qLogEHVI,
            qLogEI,
            qLogNEI,
            qNEHVI,
            qNEI,
            qUCB,
        )

        return "qLogEI/qLogNEI/qEI/qNEI/qUCB/qLogEHVI/qNEHVI OK"

    check("acquisition_functions.api", acq_api)

    def constraints_api():
        from bofire.data_models.constraints.api import (
            LinearInequalityConstraint,
            NonlinearInequalityConstraint,
        )

        return "Linear/NonlinearInequalityConstraint OK"

    check("constraints.api", constraints_api)

    # --- 3. 我们的 domain_builder / strategy_runner 能否跑 ---
    def our_domain_builder():
        from portal.core.domain_builder import build_domain
        from portal.core.excel_io import ConfigSheet, FactorSpec, OutputSpec, StrategySpec

        cfg = ConfigSheet(
            factors=[
                FactorSpec(key="temp", type="continuous", trainable=True, low=0, high=100),
                FactorSpec(
                    key="solvent",
                    type="categorical",
                    trainable=True,
                    categories=["甲苯", "THF"],
                ),
            ],
            outputs=[OutputSpec(key="Yield", direction="maximize")],
            strategy=StrategySpec(),
        )
        d = build_domain(cfg)
        return f"Domain OK: inputs={len(d.inputs.features)} outputs={len(d.outputs.features)}"

    check("portal.core.domain_builder.build_domain", our_domain_builder)

    def our_strategy_runner():
        from portal.core.domain_builder import build_domain
        from portal.core.excel_io import ConfigSheet, FactorSpec, OutputSpec, RecordRow, StrategySpec
        from portal.core.strategy_runner import run_ask

        cfg = ConfigSheet(
            factors=[
                FactorSpec(key="temp", type="continuous", trainable=True, low=0, high=100),
            ],
            outputs=[OutputSpec(key="Yield", direction="maximize")],
            strategy=StrategySpec(),
        )
        records = [
            RecordRow(
                trial_id=f"T{i}",
                batch_id=None,
                factors={"temp": 50 + i},
                output={"Yield": 70 + i},
            )
            for i in range(6)
        ]
        cands = run_ask(cfg, records, next_trial_id_start=100)
        return f"run_ask → {len(cands)} candidates, first={cands[0].factors if cands else None}"

    check("portal.core.strategy_runner.run_ask", our_strategy_runner)

    def our_convergence():
        from portal.core.convergence import evaluate_convergence
        from portal.core.excel_io import ConfigSheet, FactorSpec, OutputSpec, RecordRow, StrategySpec

        cfg = ConfigSheet(
            factors=[FactorSpec(key="t", type="continuous", trainable=True, low=0, high=100)],
            outputs=[OutputSpec(key="Yield", direction="maximize")],
            strategy=StrategySpec(),
        )
        records = [
            RecordRow(trial_id=f"T{i}", batch_id=None, factors={}, output={"Yield": 80 + i * 0.1})
            for i in range(8)
        ]
        r = evaluate_convergence(records, cfg)
        return f"{r.message} | curve={len(r.curve)} points"

    check("portal.core.convergence.evaluate_convergence", our_convergence)

    # --- 汇总 ---
    print("\n" + "=" * 70)
    ok = 0
    for name, passed, msg in results:
        mark = "✅" if passed else "❌"
        print(f"{mark} {name}\n    {msg}")
        ok += passed
    print("=" * 70)
    print(f"{ok}/{len(results)} 通过")
    return 0 if ok == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
