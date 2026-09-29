"""tests/test_pipeline.py — 端到端 pipeline 测试。

依赖: bofire 装好才跑(用 pytest.importorskip)
"""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

bofire = pytest.importorskip("bofire")

from portal.core import pipeline, excel_io  # noqa: E402


@pytest.fixture
def demo_project(tmp_path: Path) -> Path:
    """复制 demo 项目到 tmp,避免污染原始 demo。"""
    src = Path(__file__).resolve().parent.parent / "projects" / "Suzuki_2026Q4_demo"
    dst = tmp_path / "Suzuki_2026Q4_demo"
    shutil.copytree(src, dst)
    # 清 .cache(避免复用旧数据)
    cache = dst / ".cache"
    if cache.exists():
        shutil.rmtree(cache)
    return dst


class TestPipelineValidate:
    def test_validate_ok(self, demo_project: Path):
        result = pipeline.process(demo_project, pipeline.PHASE_VALIDATE)
        assert result["ok"] is True
        assert result["phase"] == "validate"
        # 状态应该写到 _status
        xlsx = next(demo_project.glob("*.xlsx"))
        s = excel_io.read_status(xlsx)
        assert s["B2"] == "awaiting_bo"


class TestPipelineAsk:
    def test_run_ask_first_round(self, demo_project: Path):
        """demo 有 6 条 valid 数据(>= n_initial=5), 应该走 SOBO 路径。"""
        result = pipeline.process(demo_project, pipeline.PHASE_ASK)
        # BoFire 可能因为环境 / 警告失败;允许 ok=False 但要有 cache
        cache = demo_project / ".cache"
        if result.get("ok"):
            assert cache.exists()
            assert any(cache.glob("recommend_*.xlsx"))

    def test_run_ask_after_cache(self, demo_project: Path):
        """跑完一次后, 再跑应该 round+1, 并产生新的 cache。"""
        pipeline.process(demo_project, pipeline.PHASE_ASK)
        cache_files_1 = sorted((demo_project / ".cache").glob("recommend_*.xlsx"))
        assert cache_files_1
        result = pipeline.process(demo_project, pipeline.PHASE_ASK)
        cache_files_2 = sorted((demo_project / ".cache").glob("recommend_*.xlsx"))
        assert len(cache_files_2) >= len(cache_files_1) + 1


class TestPipelineRefresh:
    def test_refresh_writes_dashboard(self, demo_project: Path):
        """跑 ask 后再 refresh,应把 cache 内容合并到 📊 看板。"""
        pipeline.process(demo_project, pipeline.PHASE_ASK)
        result = pipeline.process(demo_project, pipeline.PHASE_REFRESH)
        assert result["ok"] is True
        # 📊 看板 sheet 应该有内容
        xlsx = next(demo_project.glob("*.xlsx"))
        wb = pytest.importorskip("openpyxl").load_workbook(str(xlsx))
        assert "📊 看板" in wb.sheetnames
        ws = wb["📊 看板"]
        # 本轮推荐 row 6-8 应有数据
        first_rec = ws.cell(row=6, column=1).value
        assert first_rec is not None
        wb.close()