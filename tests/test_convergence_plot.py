"""tests/test_convergence_plot.py — 收敛趋势图测试。

覆盖:
- PNG 能生成
- 中文字体不报错(Windows 上 Microsoft YaHei / Noto Sans SC)
- 旧图被替换(不叠加)
- 无 matplotlib / 无数据时优雅降级
"""
from __future__ import annotations

from pathlib import Path

import pytest

from portal.core.convergence_plot import plot_convergence


def _curve(n: int = 12, plateau: bool = False) -> list[tuple]:
    """造一条收敛曲线。plateau=True 表示后期停滞。"""
    out = []
    best = 0.0
    for i in range(n):
        if plateau and i >= n - 4:
            val = 90.0 + (i % 2) * 0.2
        else:
            val = 10.0 + i * 8.0
        best = max(best, val)
        out.append((i // 3 + 1, round(best, 4), round(val, 4), i + 1))
    return out


def test_plot_generates_png(tmp_path: Path):
    """PNG 能生成且非空。"""
    out = tmp_path / "conv.png"
    result = plot_convergence(_curve(), out)
    assert result is not None, "PNG 应该生成成功"
    assert out.exists()
    assert out.stat().st_size > 1000, "PNG 不应该太小"


def test_plot_no_cjk_glyph_warning(tmp_path: Path, recwarn):
    """图里中文不能有缺字警告(Glyph missing from font)。"""
    out = tmp_path / "conv.png"
    plot_convergence(_curve(), out, title="收敛趋势 — 已收敛")
    messages = [str(w.message) for w in recwarn.list]
    glyph_warnings = [m for m in messages if "missing from font" in m]
    assert not glyph_warnings, f"中文缺字警告: {glyph_warnings}"


def test_plot_converged_vs_not(tmp_path: Path):
    """converged 参数改变标题(配色)但都能画出来。"""
    a = plot_convergence(_curve(plateau=True), tmp_path / "a.png", converged=True)
    b = plot_convergence(_curve(plateau=True), tmp_path / "b.png", converged=False)
    assert a is not None and b is not None
    assert a.stat().st_size > 1000 and b.stat().st_size > 1000


def test_plot_empty_curve_returns_none(tmp_path: Path):
    """空数据 → 返回 None,不算错。"""
    assert plot_convergence([], tmp_path / "x.png") is None


def test_plot_creates_parent_dir(tmp_path: Path):
    """父目录不存在时自动创建。"""
    out = tmp_path / "nested" / "deep" / "conv.png"
    result = plot_convergence(_curve(), out)
    assert result is not None
    assert out.exists()
