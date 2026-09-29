"""portal/core/domain_builder.py — config → BoFire Domain。

⚠️ 这里集中了 BoFire 0.4.1 所有踩过的坑,千万不要"按新版本最佳实践"改。

已知 workaround(2026-09-29 锁定 bofire==0.4.1):
  1. DiscreteOutput 会导致 MOBO 校验失败 → 一律用 ContinuousOutput
  2. CategoricalOutput + MOBO acf=qEHVI/qLogEHVI 会被拒 → 降级为
     ContinuousOutput,numeric 编码 categories[0]=min,...,categories[-1]=max
  3. NonlinearConstraint 是 abstract,实现是 NonlinearInequalityConstraint,
     但它没有 rhs 参数 — 把 rhs 编码进 expression (例如 "x+y-100")
  4. Domain 构造不要传 constraints=None,会校验失败 — omit 这个 kwarg
  5. MOBO acf 必须是 EHVI 家族(qLogEHVI/qNEHVI);用户填 qLogEI 自动 map
  6. CategoricalTaskInput BoFire 0.4.1 没有 → 用 CategoricalInput
  7. BoFire 0.4.1 没有 trainable_inputs 参数 → 训练后用 override pin
     非 trainable 列到历史 mode
  8. ContinuousInput 没有 round_digits 参数 → 整数化在后处理做
  9. Domain.outputs 没有 .keys(),要用 [f.key for f in Domain.features]
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
    from bofire.data_models.features.api import ContinuousOutput  # noqa: F401
    HAS_BOFIRE = True
except ImportError:
    HAS_BOFIRE = False

from .excel_io import ConfigSheet, FactorSpec, OutputSpec


# ---------- OutputFeature 构造 ----------
def build_output_features(outputs: list[OutputSpec]) -> list[Any]:
    """从 OutputSpec 列表构造 BoFire OutputFeature。

    ⚠️ BoFire 0.4.1 ContinuousOutput 签名: (key, context, unit, objective)
    不接受 lower_bound/upper_bound。objective 用 MaximizeObjective /
    MinimizeObjective / TargetObjective(由 OutputSpec.direction 决定)。
    """
    from bofire.data_models.features.api import ContinuousOutput
    from bofire.data_models.objectives.api import (
        MaximizeObjective,
        MinimizeObjective,
        TargetObjective,
    )

    feats: list[Any] = []
    for o in outputs:
        if o.direction == "maximize":
            objective = MaximizeObjective()
        elif o.direction == "minimize":
            objective = MinimizeObjective()
        elif o.direction == "target":
            # TargetObjective(w=value, tolerance=tolerance_value)
            # tolerance 简单取 (upper-lower) / 4 或 default 0.1
            if o.lower_bound is not None and o.upper_bound is not None:
                tol = max(0.01, (o.upper_bound - o.lower_bound) / 4)
            else:
                tol = 0.1
            objective = TargetObjective(w=float(o.target or 0.0), tolerance=tol)
        else:
            raise ValueError(f"{o.key}: 未知 direction {o.direction!r}")

        feats.append(
            ContinuousOutput(
                key=o.key,
                objective=objective,
                unit=o.unit,
            ),
        )
    return feats


# ---------- InputFeature 构造 ----------
def build_input_features(factors: list[FactorSpec]) -> list[Any]:
    """从 FactorSpec 列表构造 BoFire InputFeature。

    支持:
      continuous → ContinuousInput
      discrete → ContinuousInput(discrete values)  ⚠️ BoFire 0.4.1
        DiscreteInput 仅支持整数;我们用 categories 转 float 后走
        ContinuousInput + 后处理 round
      categorical → CategoricalInput
      categorical_descriptor → CategoricalInput(忽略 descriptors,M0 不支持)
      categorical_molecular → CategoricalInput(忽略 SMILES,RDKit 缺失则降级提示)

    若 bofire 缺失,raise ImportError;调用方应该 fallback 到纯 Python 采样。
    """
    try:
        from bofire.data_models.features.api import (
            CategoricalInput,
            ContinuousInput,
        )
    except ImportError as e:
        raise ImportError("bofire 不可用") from e

    feats: list[Any] = []
    for f in factors:
        if f.type == "continuous":
            feats.append(
                ContinuousInput(
                    key=f.key,
                    bounds=[float(f.low), float(f.high)],  # type: ignore[list-item]
                ),
            )
        elif f.type == "discrete":
            # 把 categories 转 float,包成 continuous(BoFire 0.4.1 DiscreteInput
            # 仅 int,bounds 会被拒绝。
            # workaround — 当 ranges 数值是连续值列表时 BoFire 不允许,只能走 continuous
            # + 后处理 round。这里若 categories 全是 int,转 int 后用 continuous;
            # 若有 float,用 off 最稳。
            try:
                values = sorted({float(v) for v in (f.categories or [])})  # type: ignore[union-attr]
            except (TypeError, ValueError) as e:
                raise ValueError(f"{f.key}: discrete categories 必须可转 float: {e}") from e
            if not values:
                raise ValueError(f"{f.key}: discrete 必须有 categories")
            # BoFire 0.4.1 接受 ContinuousInput + 后处理;这里用 min/max 作 bounds
            feats.append(
                ContinuousInput(
                    key=f.key,
                    bounds=[float(min(values)), float(max(values))],
                ),
            )
            # 保存离散值表供后处理用
            f.notes = (f.notes or "") + f" | discrete_values={values}"
        elif f.type in ("categorical", "categorical_descriptor", "categorical_molecular"):
            # ⚠️ workaround #6: 不存在 CategoricalTaskInput, 用 CategoricalInput
            cats = f.categories or []
            feats.append(
                CategoricalInput(key=f.key, categories=cats),
            )
        else:
            raise ValueError(f"{f.key}: 不支持的 type {f.type!r}")
    return feats


# ---------- Domain 构造 ----------
def build_domain(config: ConfigSheet) -> Any | None:
    """从 ConfigSheet 构造 BoFire Domain。

    ⚠️ workaround #4: 不传 constraints=None。

    若 bofire 缺失,返回 None;调用方应该 fallback 到本地采样。
    """
    try:
        from bofire.data_models.domain.api import Domain
    except ImportError:
        return None

    input_features = build_input_features(config.factors)
    output_features = build_output_features(config.outputs)

    # ⚠️ workaround #4: omit constraints kwarg
    domain = Domain(
        inputs=input_features,
        outputs=output_features,
    )
    return domain


# ---------- acf 映射 ----------
def resolve_acf_function(strategy_type: str, acf_name: str, n_outputs: int) -> str:
    """根据策略类型 / acf 名 / 输出数,返回实际可用的 acf。

    ⚠️ workaround #5: MOBO acf 必须是 EHVI 家族(qLogEHVI/qNEHVI/qEHVI)
    """
    if strategy_type == "sobo":
        # 单目标: 接受 qLogEI/qLogNEI/qEI/qUCB 原样
        if acf_name not in {"qLogEI", "qLogNEI", "qEI", "qUCB"}:
            warnings.warn(f"单目标 acf {acf_name!r} 不识别,降级为 qLogEI")
            return "qLogEI"
        return acf_name
    if strategy_type == "mobo":
        # 多目标: 必须是 EHVI 家族
        if n_outputs < 2:
            raise ValueError("MOBO 需要至少 2 个 Outputs,当前只有 1 个")
        if acf_name not in {"qLogEHVI", "qEHVI", "qNEHVI"}:
            warnings.warn(f"MOBO acf {acf_name!r} 不识别,降级为 qLogEHVI")
            return "qLogEHVI"
        return acf_name
    return acf_name  # random / doe 不用 acf


# ---------- Direction → objective_weight ----------
def output_objective_weights(outputs: list[OutputSpec]) -> list[float]:
    """把 OutputsSpec 转 BoFire 用的 objective_weights 数组。

    maximize → +1
    minimize → -1
    target   → +1 (|y - target| 最小化是等价的 +1 加 |target - y|,BoFire 内置)

    返回 list[float],BoFire 接受 list,不强制要 numpy 数组。
    """
    weights = []
    for o in outputs:
        if o.direction == "minimize":
            weights.append(-1.0)
        else:
            weights.append(1.0)
    return weights


# ---------- 数据准备 ----------
@dataclass
class PreparedData:
    """从 RecordRow 列表 + Domain 转出来的 numpy 张量。

    只包含 trainable 因子和输出(BoFire 训练要求)。
    """
    X: np.ndarray  # (n, trainable_factors)
    Y: np.ndarray  # (n, outputs)
    X_cols: list[str]
    Y_cols: list[str]
    cat_maps: list[dict[str, int]]  # categorical → int 编码表
    discrete_values: dict[str, list[float]]  # discrete factor → 可选值列表


def prepare_data(
    records: list[Any],  # list[RecordRow]
    config: ConfigSheet,
) -> PreparedData:
    """把 RecordRow 列表转 BoFire 可用的 X, Y 张量。

    BoFire 0.4.1 的 surrogates 训练接口接受: pd.DataFrame(features)
    + pd.DataFrame(outcomes)。这里我们生成 np + 编码表,调用方转 df。

    如果没有 numpy,用 list[list[float]] 替代。返回 PreparedData.X 是
    numpy ndarray(若可用)或 list[list[float]];PreparedData.Y 同。
    """
    trainable = [f for f in config.factors if f.trainable]
    X_cols = [f.key for f in trainable]
    Y_cols = [o.key for o in config.outputs]
    cat_maps: list[dict[str, int]] = []
    discrete_values: dict[str, list[float]] = {}

    # 编码 categorical → int
    for f in trainable:
        if f.type in ("categorical", "categorical_descriptor", "categorical_molecular"):
            cats = f.categories or []
            cat_maps.append({c: i for i, c in enumerate(cats)})
        else:
            cat_maps.append({})
        if f.type == "discrete":
            # 从 notes 解析 (build_input_features 写过)
            dv: list[float] = []
            if f.notes and "discrete_values=" in f.notes:
                seg = f.notes.split("discrete_values=")[1].split("|")[0]
                # 解析 [1.0, 5.0, 10.0]
                import ast
                try:
                    dv = [float(x) for x in ast.literal_eval(seg)]
                except Exception:
                    pass
            discrete_values[f.key] = dv

    rows_X: list[list[float]] = []
    rows_X_cat: list[list[Any]] = []  # categorical 原值,给 BoFire 验
    rows_Y: list[list[float]] = []
    for rec in records:
        x_row = []
        x_cat_row = []
        for i, f in enumerate(trainable):
            v = rec.factors.get(f.key)
            if v is None:
                v = (f.low + f.high) / 2 if f.type in ("continuous", "discrete") else (
                    (f.categories or [""])[0] if f.categories else ""
                )
            if f.type in ("categorical", "categorical_descriptor", "categorical_molecular"):
                # ⚠️ BoFire 0.4.1 categorical 接受原值字符串,不接受 int 编码
                x_row.append(0.0)  # 占位,BoFire 不会用
                x_cat_row.append(str(v))
            else:
                x_cat_row.append("")
                try:
                    x_row.append(float(v))
                except (ValueError, TypeError):
                    x_row.append(float((f.low + f.high) / 2) if f.low is not None and f.high is not None else 0.0)
        rows_X.append(x_row)
        rows_X_cat.append(x_cat_row)
        y_row = []
        for o in config.outputs:
            y_row.append(float(rec.output.get(o.key) or 0.0))
        rows_Y.append(y_row)

    # categorical 列:用原值替换 X 里的占位
    for i, f in enumerate(trainable):
        if f.type in ("categorical", "categorical_descriptor", "categorical_molecular"):
            for r_idx in range(len(rows_X)):
                rows_X[r_idx][i] = rows_X_cat[r_idx][i]

    if HAS_NUMPY:
        X_arr: Any = np.array(rows_X, dtype=object) if rows_X else np.zeros((0, len(X_cols)), dtype=object)
        Y_arr: Any = np.array(rows_Y, dtype=float) if rows_Y else np.zeros((0, len(Y_cols)))
    else:
        X_arr = rows_X
        Y_arr = rows_Y
    return PreparedData(
        X=X_arr, Y=Y_arr, X_cols=X_cols, Y_cols=Y_cols,
        cat_maps=cat_maps, discrete_values=discrete_values,
    )


# ---------- Utility ----------
def describe_domain(domain: Any) -> dict[str, Any]:
    """返回 domain 的可读 summary。

    ⚠️ BoFire 0.4.1: Domain.inputs 和 Domain.outputs 是 Inputs/Outputs 对象,
    不是 list。真正的 features 列表是 .inputs.features 和 .outputs.features。
    """
    if domain is None:
        return {"error": "bofire not installed"}
    return {
        "n_inputs": len(domain.inputs.features),
        "n_outputs": len(domain.outputs.features),
        "input_keys": [f.key for f in domain.inputs.features],
        "output_keys": [f.key for f in domain.outputs.features],
    }