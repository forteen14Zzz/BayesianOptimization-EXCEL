"""portal/core/convergence_plot.py — 收敛趋势图(PNG)。

画在 📊 看板 的收敛区右侧(G 列起),让同事一眼看出
"BO 还在进步还是已经涨不动了"。

⚠️ 两条曲线:
  - best_so_far    — 历史最好成绩(单调不减,平台期 = 收敛)
  - mean_of_best3 — 最近 3 个实验的均值(抖动大,看趋势)

⚠️ 无 matplotlib 时优雅降级(返回 None,不报错)。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any


# matplotlib 在无显示环境必须用 Agg
# ⚠️ 中文字体 — Windows 上有 Microsoft YaHei / Noto Sans SC,
#    不设的话图上的中文全变方块(Glyph missing from font)
_CJK_FONTS = [
    "Microsoft YaHei",
    "Noto Sans SC",
    "SimHei",
    "PingFang SC",
    "WenQuanYi Zen Hei",
    "Source Han Sans SC",
]


def _setup_fonts():
    """配置 matplotlib 用中文字体。返回是否成功找到。"""
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib import font_manager

    available = {f.name for f in font_manager.fontManager.ttflist}
    picked = [f for f in _CJK_FONTS if f in available]
    if picked:
        matplotlib.rcParams["font.sans-serif"] = picked + list(
            matplotlib.rcParams.get("font.sans-serif", [])
        )
        matplotlib.rcParams["font.family"] = "sans-serif"
    # 负号用 ASCII,否则缺失字形会画成方块
    matplotlib.rcParams["axes.unicode_minus"] = False
    return bool(picked)


def _plt():
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    _setup_fonts()
    return plt


def plot_convergence(
    curve: list[tuple],
    out_path: Path,
    converged: bool = False,
    threshold: float | None = None,
    title: str = "收敛趋势",
) -> Path | None:
    """画收敛趋势图,存成 PNG。

    Args:
        curve:      [(round, best_so_far, mean_of_best3, n_measured), ...]
        out_path:   PNG 输出路径
        converged:  是否已收敛(影响配色和标题)
        threshold:  收敛阈值(画参考线)
        title:      图表标题

    Returns:
        PNG 路径;失败返回 None
    """
    if not curve:
        return None

    try:
        plt = _plt()
    except Exception:
        return None

    try:
        xs = [p[0] for p in curve]
        best = [p[1] for p in curve]
        mean3 = [p[2] for p in curve]
        n_meas = [p[3] for p in curve]

        # 颜色:收敛=橙,未收敛=绿
        color = "#ED7D31" if converged else "#70AD47"

        fig, (ax1, ax2) = plt.subplots(
            2,
            1,
            figsize=(8.6, 4.6),
            gridspec_kw={"height_ratios": [3, 1]},
            layout="constrained",
            dpi=140,
        )

        # --- 上图:成绩曲线 ---
        ax1.plot(xs, best, "-o", color=color, lw=2.2, ms=4.5,
                 label="最好成绩 best_so_far")
        ax1.plot(xs, mean3, "--s", color="#5B9BD5", lw=1.5, ms=3.5,
                 label="最近3个均值 mean_of_best3")
        ax1.set_ylabel("目标值")
        ax1.grid(alpha=0.3, ls="--")
        ax1.legend(loc="lower right", fontsize=9, framealpha=0.9)

        if converged:
            # ⚠️ 不用 ✅ emoji — Microsoft YaHei 没有这个字形(会警告 + 画成方块)
            ax1.set_title(f"{title} — 已收敛", fontsize=12, color=color, weight="bold")
        else:
            ax1.set_title(title, fontsize=12, weight="bold")

        # 平台期高亮:最后几段 best 不变
        if len(best) >= 2:
            plateau = best[-1] - best[-min(len(best), 6)] < (threshold or 0.01)
            if plateau and not converged:
                ax1.axhspan(
                    best[-1] - (threshold or 0.01),
                    best[-1] + (threshold or 0.01),
                    color="#FFC000",
                    alpha=0.18,
                    label=f"±{threshold or 0.01} 阈值带",
                )
                ax1.legend(loc="lower right", fontsize=9, framealpha=0.9)

        # --- 下图:实验数 ---
        ax2.plot(xs, n_meas, "-", color="#A5A5A5", lw=1.5)
        ax2.set_ylabel("累计实验数")
        ax2.set_xlabel("轮次")
        ax2.grid(alpha=0.3, ls="--")

        # ⚠️ 不用 tight_layout — 和 layout="constrained" 冲突(会警告)
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return out_path

    except Exception:
        try:
            plt.close("all")
        except Exception:
            pass
        return None


def plot_feature_importance(
    importances: dict[str, float],
    out_path: Path,
    title: str = "特征重要性 (SHAP)",
) -> Path | None:
    """画特征重要性条形图(可选功能,失败不影响主流程)。"""
    if not importances:
        return None
    try:
        plt = _plt()
    except Exception:
        return None

    try:
        items = sorted(importances.items(), key=lambda kv: kv[1])
        names = [k for k, _ in items]
        vals = [v for _, v in items]

        fig, ax = plt.subplots(figsize=(6.4, 0.42 * len(names) + 1.4),
                               layout="constrained", dpi=140)
        ax.barh(names, vals, color="#5B9BD5")
        ax.set_xlabel("平均 |SHAP|")
        ax.set_title(title, fontsize=12, weight="bold")
        ax.grid(axis="x", alpha=0.3, ls="--")
        for i, v in enumerate(vals):
            ax.text(v, i, f" {v:.3f}", va="center", fontsize=8)

        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return out_path
    except Exception:
        try:
            plt.close("all")
        except Exception:
            pass
        return None
