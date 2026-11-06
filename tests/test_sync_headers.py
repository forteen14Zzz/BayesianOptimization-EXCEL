"""tests/test_sync_headers.py — 🧪 记录 表头同步测试。"""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from portal.core import sync_headers
from portal.core.excel_io import read_records


@pytest.fixture
def template_xlsx(tmp_path: Path) -> Path:
    src = Path(__file__).resolve().parent.parent / "templates" / "project_template.xlsx"
    dst = tmp_path / "test_project.xlsx"
    shutil.copy2(src, dst)
    return dst


def _write_records(xlsx: Path, headers: list[str], rows: list[list]) -> None:
    """写 🧪 记录 sheet 的表头和数据(先清空整行,避免残留旧表头)。"""
    from openpyxl import load_workbook
    wb = load_workbook(str(xlsx))
    ws = wb["🧪 记录"]
    # 先清空 row 4 整行(模板预设了 11 个表头)
    for c in range(1, 20):
        ws.cell(row=4, column=c).value = None
    for i, h in enumerate(headers):
        ws.cell(row=4, column=i + 1, value=h)
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            ws.cell(row=5 + r_idx, column=c_idx + 1, value=val)
    wb.save(str(xlsx))
    wb.close()


class TestBuildRecordHeaders:
    def test_order(self):
        headers = sync_headers.build_record_headers(
            ["温度", "溶剂"], ["Yield"]
        )
        assert headers == [
            "trial_id", "batch_id",
            "温度", "溶剂",
            "Yield",
            "valid", "status", "来源", "录入时间",
        ]

    def test_empty(self):
        headers = sync_headers.build_record_headers([], [])
        assert headers == [
            "trial_id", "batch_id",
            "valid", "status", "来源", "录入时间",
        ]


class TestSyncRecordHeaders:
    def test_no_change_when_same(self, template_xlsx: Path):
        """表头已经一致 → 不改数据。"""
        factors = ["温度", "催化剂用量", "反应时间", "溶剂"]
        outputs = ["Yield"]
        expected = sync_headers.build_record_headers(factors, outputs)
        _write_records(
            template_xlsx,
            expected,
            [["T0001", "B1", 60, 2.0, 12, "甲苯", 78.5, 1, "measured"]],
        )
        result = sync_headers.sync_record_headers(template_xlsx, factors, outputs)
        assert result["changed"] is False
        assert result["migrated"] == 0

    def test_change_headers_and_migrate(self, template_xlsx: Path):
        """表头变了 → 按列名迁移数据。"""
        old_headers = ["trial_id", "batch_id", "温度", "催化剂用量", "溶剂", "Yield", "valid"]
        _write_records(
            template_xlsx,
            old_headers,
            [["T0001", "B1", 60, 2.0, "甲苯", 78.5, 1]],
        )
        # 新因子顺序不同
        new_factors = ["溶剂", "温度", "催化剂用量"]
        result = sync_headers.sync_record_headers(template_xlsx, new_factors, ["Yield"])

        assert result["changed"] is True
        assert result["migrated"] == 1
        assert result["new_headers"][:5] == [
            "trial_id", "batch_id", "溶剂", "温度", "催化剂用量"
        ]

        # 验证数据位置正确
        from openpyxl import load_workbook
        wb = load_workbook(str(template_xlsx))
        ws = wb["🧪 记录"]
        headers = [ws.cell(row=4, column=c).value for c in range(1, 11)]
        row5 = [ws.cell(row=5, column=c).value for c in range(1, 11)]
        wb.close()

        idx = {h: i for i, h in enumerate(headers)}
        assert row5[idx["溶剂"]] == "甲苯"
        assert row5[idx["温度"]] == 60
        assert row5[idx["催化剂用量"]] == 2.0
        assert row5[idx["Yield"]] == 78.5

    def test_position_fallback(self, template_xlsx: Path):
        """列名对不上但长度一致 → 按位置兜底。"""
        old_headers = ["trial_id", "batch_id", "A", "B", "C", "D", "E"]
        _write_records(
            template_xlsx,
            old_headers,
            [["T0001", "B1", 1, 2, 3, 4, 5]],
        )
        new_factors = ["A", "B", "C", "D", "E"]
        result = sync_headers.sync_record_headers(template_xlsx, new_factors, [])

        assert result["changed"] is True
        assert result["migrated"] == 1

        from openpyxl import load_workbook
        wb = load_workbook(str(template_xlsx))
        ws = wb["🧪 记录"]
        headers = [ws.cell(row=4, column=c).value for c in range(1, 12)]
        row5 = [ws.cell(row=5, column=c).value for c in range(1, 12)]
        wb.close()

        idx = {h: i for i, h in enumerate(headers)}
        assert row5[idx["A"]] == 1
        assert row5[idx["B"]] == 2
        assert row5[idx["C"]] == 3

    def test_validations_follow_headers(self, template_xlsx: Path):
        """valid / status 列位置变了 → 下拉菜单跟着走。"""
        from openpyxl import load_workbook

        # 3 个因子 → valid 在 col 8, status 在 col 9
        sync_headers.sync_record_headers(
            template_xlsx, ["f1", "f2", "f3"], ["Yield"]
        )
        wb = load_workbook(str(template_xlsx))
        ws = wb["🧪 记录"]
        headers = [ws.cell(row=4, column=c).value for c in range(1, 12)]
        dv_info = [
            (str(dv.sqref), dv.formula1)
            for dv in ws.data_validations.dataValidation
        ]
        wb.close()

        valid_idx = headers.index("valid")
        status_idx = headers.index("status")
        assert f"{chr(ord('A') + valid_idx)}5:" in str(dv_info)
        assert any("0,1" in f for _, f in dv_info)
        assert any("measured,pending,failed" in f for _, f in dv_info)

    def test_read_records_after_sync(self, template_xlsx: Path):
        """同步后 read_records 能正确读到因子。"""
        factors = ["base", "ligand"]
        outputs = ["Yield"]
        _write_records(
            template_xlsx,
            ["trial_id", "batch_id", "base", "ligand", "Yield", "valid"],
            [["T0001", "B1", "KOAc", "PPh3", 55.0, 1]],
        )
        sync_headers.sync_record_headers(template_xlsx, factors, outputs)
        records = read_records(template_xlsx, factors, outputs)
        assert len(records) == 1
        assert records[0].factors["base"] == "KOAc"
        assert records[0].factors["ligand"] == "PPh3"
        assert records[0].output["Yield"] == 55.0

    def test_hint_row_not_migrated_as_data(self, template_xlsx: Path):
        """提示行(💡)不该被当成数据。"""
        from openpyxl import load_workbook
        wb = load_workbook(str(template_xlsx))
        ws = wb["🧪 记录"]
        ws.cell(row=4, column=1, value="trial_id")
        ws.cell(row=4, column=2, value="batch_id")
        ws.cell(row=4, column=3, value="温度")
        ws.cell(row=5, column=1, value="💡 提示文字")
        wb.save(str(template_xlsx))
        wb.close()

        result = sync_headers.sync_record_headers(
            template_xlsx, ["温度"], ["Yield"]
        )
        assert result["migrated"] == 0