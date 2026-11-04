"""tests/test_watcher.py — watcher / recommender 单元测试。"""
from __future__ import annotations

import shutil
import time
from pathlib import Path

import pytest

from portal import watcher
from portal.core import recommender
from portal.core.excel_io import read_status, write_status


@pytest.fixture
def demo_project(tmp_path: Path) -> Path:
    src = Path(__file__).resolve().parent.parent / "projects" / "Suzuki_2026Q4_demo"
    dst = tmp_path / "Suzuki_2026Q4_demo"
    shutil.copytree(src, dst)
    cache = dst / ".cache"
    if cache.exists():
        shutil.rmtree(cache)
    return dst


class TestRecommender:
    def test_write_cache(self, demo_project: Path):
        from portal.core.strategy_runner import Candidate
        cands = [
            Candidate("T0001", {"温度": 60, "催化剂用量": 2.0, "反应时间": 12, "溶剂": "甲苯"}, 78.5, 2.0, None),
            Candidate("T0002", {"温度": 80, "催化剂用量": 1.5, "反应时间": 8, "溶剂": "THF"}, 82.1, 2.5, None),
            Candidate("T0003", {"温度": 70, "催化剂用量": 2.5, "反应时间": 16, "溶剂": "DMF"}, 75.3, 2.2, None),
        ]
        path = recommender.write_recommend_cache(
            demo_project,
            cands,
            factor_keys=["温度", "催化剂用量", "反应时间", "溶剂"],
            output_keys=["Yield"],
            round_num=1,
        )
        assert path.exists()
        assert path.suffix == ".xlsx"
        assert path.parent == demo_project / ".cache"

        # 内容
        from openpyxl import load_workbook
        wb = load_workbook(str(path))
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        assert len(rows) == 4  # 1 header + 3 candidates
        assert rows[0][0] == "trial_id"
        assert rows[1][0] == "T0001"
        assert rows[1][1] == 1  # round

    def test_find_latest_cache(self, demo_project: Path):
        # 空
        assert recommender.find_latest_cache(demo_project) is None
        # 写 2 个 cache,找最新的
        from portal.core.strategy_runner import Candidate
        c1 = Candidate("T0001", {"温度": 60}, 78.5, 2.0, None)
        c2 = Candidate("T0002", {"温度": 80}, 82.1, 2.5, None)
        recommender.write_recommend_cache(demo_project, [c1], ["温度"], ["Yield"], 1)
        time.sleep(1.1)  # 时间戳秒级,确保不同
        recommender.write_recommend_cache(demo_project, [c2], ["温度"], ["Yield"], 2)
        latest = recommender.find_latest_cache(demo_project)
        assert latest is not None
        assert "recommend_" in latest.name

    def test_merge_to_dashboard(self, demo_project: Path):
        """写 cache → merge,验证 📊 看板 收到数据。"""
        from portal.core.strategy_runner import Candidate
        cands = [
            Candidate("T0007", {"温度": 65.0, "催化剂用量": 2.0, "反应时间": 12, "溶剂": "甲苯"}, 78.5, 2.0, None),
            Candidate("T0008", {"温度": 75.0, "催化剂用量": 1.5, "反应时间": 8, "溶剂": "THF"}, 82.1, 2.5, None),
        ]
        cache = recommender.write_recommend_cache(
            demo_project, cands,
            factor_keys=["温度", "催化剂用量", "反应时间", "溶剂"],
            output_keys=["Yield"], round_num=2,
        )
        xlsx = next(demo_project.glob("*.xlsx"))
        recommender.merge_cache_to_dashboard(demo_project, xlsx, cache, 2)

        from openpyxl import load_workbook
        wb = load_workbook(str(xlsx))
        ws = wb["📊 看板"]
        # 本轮推荐 row 6 第一格应该是 T0007
        assert ws.cell(row=6, column=1).value == "T0007"
        # 历轮推荐应有追加
        # 从 row 54 起(新模板结构)
        # cache 列顺序: trial_id, round, asked_at, 因子..., Yield_mean, Yield_sd, acq_value
        # trial_id 在 col 1(不是 col 3)
        history_rows = [ws.cell(row=r, column=1).value for r in range(54, 65)]
        assert "T0007" in history_rows
        assert "T0008" in history_rows


class TestWatcher:
    def test_handler_ignores_temp_files(self, tmp_path: Path):
        """创建 `~$lock.xlsx` 临时文件,handler 应忽略。"""
        handler = watcher.Handler(tmp_path)
        called = []
        # 模拟 watchdog 事件
        from watchdog.events import FileModifiedEvent
        event = FileModifiedEvent(str(tmp_path / "~$lock.xlsx"))
        handler.on_modified(event)
        # 状态不变
        assert handler.last_md5 == {}

    def test_handler_md5_filter(self, demo_project: Path):
        """同内容触发应跳过。"""
        xlsx = next(demo_project.glob("*.xlsx"))
        handler = watcher.Handler(demo_project)
        from watchdog.events import FileModifiedEvent
        event1 = FileModifiedEvent(str(xlsx))
        handler.on_modified(event1)
        # 第二次同文件不触发
        event2 = FileModifiedEvent(str(xlsx))
        handler.on_modified(event2)