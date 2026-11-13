"""portal/core/strategy_runner.py — ask / train。

设计:
- run_ask(config, records) -> List[CandidateSpec]
  - 少数据(< n_initial) → random
  - 多数据 → SOBO + qLogEI + SingleTaskGPSurrogate
- train_predict(domain, X, Y) -> GP 预测(mean, sd)供 user 看

⚠️ BoFire 0.4.1 API:
- BoFire 可以直接用 strategies API:
  from bofire.strategies.api import SoboStrategy, RandomStrategy
  strategy = SoboStrategy(domain=domain)
  strategy.tell(failed=failed_df, pending=pending_df, X=X_df)
  candidates = strategy.ask(candidate_count=batch_size)
- 也可以走 Strategy API (strategy_data_model + step):
  from bofire.strategies.strategy import Strategy
  step = Strategy.from_spec(data_model, domain=domain)
  step.tell(...)
  step.ask(...)
  ⚠️ 0.4.1 这一版,后者更稳。

我们走 Strategy API: 直接构造 data_model。
"""
from __future__ import annotations

import warnings
from dataclasses import dataclass
from typing import Any

try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    np = None  # type: ignore[assignment]
    HAS_NUMPY = False

try:
    from bofire.strategies.api import SoboStrategy, RandomStrategy  # noqa: F401
    HAS_BOFIRE = True
except ImportError:
    SoboStrategy = None  # type: ignore[assignment]
    RandomStrategy = None  # type: ignore[assignment]
    HAS_BOFIRE = False

from .domain_builder import (
    build_domain,
    prepare_data,
    resolve_acf_function,
    output_objective_weights,
)


# ---------- 数据结构 ----------
@dataclass
class Candidate:
    """BO 推荐 1 个候选实验。"""
    trial_id: str  # T0001 这种
    factors: dict[str, Any]  # {factor_key: value}
    pred_mean: float | None = None  # GP 预测的 mean
    pred_sd: float | None = None  # GP 预测的 sd
    acq_value: float | None = None  # acquisition function value


# ---------- 主入口 ----------
def run_ask(
    config: Any,  # ConfigSheet
    records: list[Any],  # list[RecordRow]
    next_trial_id_start: int = 1,
) -> list[Candidate]:
    """根据当前 config + 历史数据,推荐 batch_size 个候选。

    主入口;供 pipeline.phase=ask 调用。
    降级链: BoFire SOBO → BoFire Random → local uniform → pure python。
    """
    domain = build_domain(config)
    if domain is None:
        # bofire 全不可用 → 纯 Python 采样
        warnings.warn("bofire 未安装,使用纯 Python 随机采样(degraded)")
        return _pure_python_ask(config, next_trial_id_start)

    n_data = sum(1 for r in records if r.valid == 1)

    if n_data < config.strategy.n_initial:
        return _random_ask(domain, config, next_trial_id_start)

    return _sobo_ask(domain, config, records, next_trial_id_start)


def _pure_python_ask(config: Any, trial_id_start: int) -> list[Candidate]:
    """最朴素 fallback — 不依赖 BoFire,也不依赖 bofire 构造的 Domain。

    仅用 config.factors 列表的 low/high/categories。
    """
    import random
    cands = []
    for i in range(config.strategy.batch_size):
        factors: dict[str, Any] = {}
        for f in config.factors:
            if f.type == "discrete":
                # ⚠️ 离散因子必须从 categories 里挑,不能连续采样
                cats = f.categories or []
                try:
                    vals = [float(v) for v in cats]
                except (TypeError, ValueError):
                    vals = []
                if vals:
                    factors[f.key] = random.choice(vals)
                else:
                    factors[f.key] = random.choice(cats) if cats else None
            elif f.type == "continuous":
                lo = float(f.low or 0.0)
                hi = float(f.high or 1.0)
                factors[f.key] = random.uniform(lo, hi)
            elif f.type in ("categorical", "categorical_descriptor", "categorical_molecular"):
                cats = f.categories or [""]
                factors[f.key] = str(random.choice(cats))
            else:
                factors[f.key] = None
        cands.append(Candidate(trial_id=f"T{trial_id_start + i:04d}", factors=factors))
    return cands


# ---------- Random strategy ----------
def _random_ask(domain: Any, config: Any, trial_id_start: int) -> list[Candidate]:
    """用 BoFire RandomStrategy 采样。"""
    try:
        from bofire.data_models.strategies.api import RandomStrategy as RandomDM
        from bofire.strategies.api import RandomStrategy as RandomImpl

        dm = RandomDM(domain=domain, seed=42)
        strat = RandomImpl(data_model=dm)
        df = strat.ask(candidate_count=config.strategy.batch_size)
        return _df_to_candidates(df, domain, trial_id_start)
    except Exception as e:
        warnings.warn(f"BoFire RandomStrategy 失败: {e}, 降级到本地均匀采样")
        return _local_uniform_ask(domain, config, trial_id_start)


# ---------- SOBO ----------
def _sobo_ask(
    domain: Any,
    config: Any,
    records: list[Any],
    trial_id_start: int,
) -> list[Candidate]:
    """GP-BO 路径: train GP, 用 acf 选 candidates。"""
    prepared = prepare_data(records, config)

    # 构造 SoboStrategy data_model
    acf = resolve_acf_function(
        config.strategy.strategy_type,
        config.strategy.acq_function,
        len(config.outputs),
    )
    try:
        from bofire.data_models.strategies.api import SoboStrategy as SoboDM
        from bofire.strategies.api import SoboStrategy as SoboStrategyImpl
        from bofire.data_models.acquisition_functions.api import qLogEI, qLogEHVI, qNEHVI, qEI, qNEI, qUCB

        # ⚠️ BoFire 0.4.1: acquisition_function 必须传 acf 实例,不是字符串
        acf_obj = {
            "qLogEI": qLogEI(),
            "qLogEHVI": qLogEHVI(),
            "qNEHVI": qNEHVI(),
            "qEI": qEI(),
            "qNEI": qNEI(),
            "qUCB": qUCB(),
        }.get(acf, qLogEI())

        sobo_dm = SoboDM(
            domain=domain,
            acquisition_function=acf_obj,
            seed=42,  # 固定种子保证可复现
            convergence_criterion=_build_convergence_criterion(config),
        )
        # ⚠️ BoFire 0.4.1: from_spec(data_model=...) 或直接 SoboStrategy(data_model=...)
        strat = SoboStrategyImpl(data_model=sobo_dm)

        # 喂数据
        try:
            import pandas as pd
            # ⚠️ BoFire 0.4.1: tell(experiments: DataFrame) 一个 df 包含
            # features + outcomes(不再是 X/Y 分开)
            exp_df = pd.DataFrame(prepared.X, columns=prepared.X_cols)
            for j, y_col in enumerate(prepared.Y_cols):
                exp_df[y_col] = prepared.Y[:, j] if hasattr(prepared.Y, "shape") else [r[j] for r in prepared.Y]
        except ImportError:
            warnings.warn("BoFire SOBO 需要 pandas, 降级到 random")
            return _random_ask(domain, config, trial_id_start)

        try:
            strat.tell(experiments=exp_df)
        except Exception as e:
            warnings.warn(f"strat.tell failed: {e}, 降级到 random")
            return _random_ask(domain, config, trial_id_start)

        df = strat.ask(candidate_count=config.strategy.batch_size)
        candidates = _df_to_candidates(df, domain, trial_id_start)

        # ⚠️ GP 预测 mean/sd: 用 strat.predict(df) 拿
        try:
            preds_df = strat.predict(experiments=df)
            # preds_df 形如: input cols + {output}_pred / {output}_sd
            # 配 candidates[i].pred_mean / pred_sd
            for i, cand in enumerate(candidates):
                if i >= len(preds_df):
                    break
                row = preds_df.iloc[i]
                # 找 *_pred 列(单输出时)
                for y_col in prepared.Y_cols:
                    pred_col = f"{y_col}_pred"
                    sd_col = f"{y_col}_sd"
                    if pred_col in preds_df.columns:
                        cand.pred_mean = float(row[pred_col])
                    if sd_col in preds_df.columns:
                        cand.pred_sd = float(row[sd_col])
        except Exception as e:
            warnings.warn(f"GP predict 失败,pred_mean 留 None: {e}")

        # ⚠️ 采集函数值(acq_value):用 strat.calc_acquisition(df)
        try:
            acq_values = strat.calc_acquisition(candidates=df)
            for i, cand in enumerate(candidates):
                if i < len(acq_values):
                    v = acq_values[i]
                    cand.acq_value = float(v) if hasattr(v, "item") else v
        except Exception as e:
            warnings.warn(f"calc_acquisition 失败,acq_value 留 None: {e}")

        # ⚠️ workaround #7: pin_non_trainable 在 pipeline 层做
        return candidates

    except Exception as e:
        warnings.warn(f"BoFire SOBO 失败: {e}, 降级到 random")
        return _random_ask(domain, config, trial_id_start)


# ---------- 工具 ----------
def _build_convergence_criterion(config: Any) -> Any | None:
    """按配置构造 BoFire 官方 ObjectiveImprovementCriterion。

    ⚠️ 官方模块 `bofire.data_models.strategies.convergence_criteria` 是
    **bofire 0.5.0** 才有的(0.4.1 没有)→ 装 0.4.1 时优雅降级为 None。

    官方公式: b_N - b_{N-n_lookback} < min_improvement → 收敛
    我们的配置:
        收敛阈值 ≡ min_improvement
        收敛窗口 ≡ n_lookback(单位:实验数)
    """
    try:
        from bofire.data_models.strategies.convergence_criteria.api import (
            ObjectiveImprovementCriterion,
        )
    except ImportError:
        # bofire < 0.5.0 — 没有官方模块
        return None

    threshold = float(getattr(config.strategy, "convergence_threshold", 0.01) or 0.01)
    window = int(getattr(config.strategy, "convergence_window", 5) or 5)
    try:
        criterion = ObjectiveImprovementCriterion(
            min_improvement=threshold,
            n_lookback=window,
        )
    except Exception:
        # 参数不合法(比如 threshold=0)→ 不挂 criterion,不影响 BO 本身
        return None

    # ⚠️ 官方判据的适用性由 strategy 在构造时校验,不适配会抛错。
    # 单目标才挂(我们的 converge 判定本身也支持多目标,走本地实现)。
    if len(getattr(config, "outputs", [])) > 1:
        return None

    return criterion


def _df_to_candidates(df: Any, domain: Any, trial_id_start: int) -> list[Candidate]:
    """把 BoFire ask() 返回的 DataFrame 转 Candidate 列表。

    BoFire ask() 返回的 DataFrame 列名就是 factor key。
    categorical 在 DataFrame 里是字符串(分类值)。
    """
    cols = [f.key for f in domain.inputs]
    cands = []
    # 用 row index 生成 trial_id
    for i, (_, row) in enumerate(df.iterrows()):
        factors = {}
        for c in cols:
            if c in df.columns:
                v = row[c]
                # numpy 类型 → Python 原生
                if hasattr(v, "item"):
                    try:
                        v = v.item()
                    except Exception:
                        pass
                factors[c] = v
        cands.append(
            Candidate(
                trial_id=f"T{trial_id_start + i:04d}",
                factors=factors,
            ),
        )
    return cands


def pin_non_trainable(
    candidates: list[Candidate],
    records: list[Any],
    config: Any,
) -> list[Candidate]:
    """workaround #7: BoFire 0.4.1 没有 trainable_inputs 参数。

    我们把 candidates 中"非 trainable 因子"覆盖为历史众数。
    """
    from collections import Counter

    trainable = {f.key for f in config.factors if f.trainable}
    pinned: list[Candidate] = []
    for c in candidates:
        new_factors = dict(c.factors)
        for f in config.factors:
            if f.trainable:
                continue
            # 收集所有 records 中此 factor 的值
            vals = [r.factors.get(f.key) for r in records if f.key in r.factors]
            if not vals:
                continue
            cnt = Counter(str(v) for v in vals)
            mode = cnt.most_common(1)[0][0]
            # 保留原类型
            new_factors[f.key] = type(vals[0])(mode) if vals else mode
        pinned.append(
            Candidate(
                trial_id=c.trial_id,
                factors=new_factors,
                pred_mean=c.pred_mean,
                pred_sd=c.pred_sd,
                acq_value=c.acq_value,
            ),
        )
    return pinned


def _local_uniform_ask(domain: Any, config: Any, trial_id_start: int) -> list[Candidate]:
    """完全本地随机采样 — fallback,不依赖 BoFire 也不依赖 numpy。"""
    import random
    cands = []
    for i in range(config.strategy.batch_size):
        factors: dict[str, Any] = {}
        for f in domain.inputs:
            key = f.key
            if hasattr(f, "bounds") and f.bounds:
                lo, hi = float(f.bounds[0]), float(f.bounds[1])
                factors[key] = random.uniform(lo, hi)
            elif hasattr(f, "categories") and f.categories:
                factors[key] = str(random.choice(list(f.categories)))
            else:
                factors[key] = None
        cands.append(
            Candidate(trial_id=f"T{trial_id_start + i:04d}", factors=factors),
        )
    return cands