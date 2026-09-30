"""tests/test_init_new_project.py — init_new_project 单元测试。"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest


SCRIPT = Path(__file__).resolve().parent.parent / "scripts" / "init_new_project.py"


def run(args: list[str], cwd: Path) -> subprocess.CompletedProcess:
    """跑 init_new_project.py 脚本。"""
    return subprocess.run(
        [sys.executable, str(SCRIPT)] + args,
        cwd=str(cwd),
        capture_output=True,
        text=True,
    )


@pytest.fixture
def projects_dir(tmp_path: Path) -> Path:
    """建一个 tmp 项目根目录(含 templates/project_template.xlsx)。"""
    repo = tmp_path / "repo"
    (repo / "scripts").mkdir(parents=True)
    (repo / "templates").mkdir(parents=True)
    # 复制真实模板
    real_template = (
        Path(__file__).resolve().parent.parent / "templates" / "project_template.xlsx"
    )
    if real_template.exists():
        import shutil
        shutil.copy2(real_template, repo / "templates" / "project_template.xlsx")
    return repo


class TestInitNewProject:
    def test_basic_init(self, projects_dir: Path):
        # 我们用真实 repo(因为脚本 hardcode ROOT)
        # 改成直接 import init_new_project 然后跑
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
        import init_new_project as inp

        # 临时改 ROOT
        original_root = inp.ROOT
        inp.ROOT = projects_dir
        try:
            target = inp.init_project("Suzuki_TestUser")
            assert target.exists()
            assert (target / "Suzuki_TestUser.xlsx").exists()
            assert (target / "LOG.txt").exists()
            # bat 复制是 optional — 源 bat 不存在就跳过
            bat = target / "跑_BO_点这里.bat"
            if (inp.ROOT / "projects" / "Suzuki_2026Q4_demo" / "跑_BO_点这里.bat").exists():
                assert bat.exists()
        finally:
            inp.ROOT = original_root
            sys.path.remove(str(Path(__file__).resolve().parent.parent / "scripts"))

    def test_init_status_set_correctly(self, projects_dir: Path):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
        import init_new_project as inp
        from openpyxl import load_workbook

        original_root = inp.ROOT
        inp.ROOT = projects_dir
        try:
            target = inp.init_project("MyProject_2026Q4")
            wb = load_workbook(target / "MyProject_2026Q4.xlsx")
            ws = wb["_status"]
            assert ws["B2"].value == "awaiting_config"
            assert ws["B5"].value == "MyProject_2026Q4"
            wb.close()
        finally:
            inp.ROOT = original_root
            sys.path.remove(str(Path(__file__).resolve().parent.parent / "scripts"))

    def test_empty_name_rejected(self, projects_dir: Path):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
        import init_new_project as inp

        original_root = inp.ROOT
        inp.ROOT = projects_dir
        try:
            with pytest.raises(ValueError, match="不能为空"):
                inp.init_project("")
            with pytest.raises(ValueError, match="不能为空"):
                inp.init_project("   ")
        finally:
            inp.ROOT = original_root
            sys.path.remove(str(Path(__file__).resolve().parent.parent / "scripts"))

    def test_path_traversal_rejected(self, projects_dir: Path):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
        import init_new_project as inp

        original_root = inp.ROOT
        inp.ROOT = projects_dir
        try:
            with pytest.raises(ValueError, match="非法字符"):
                inp.init_project("../escape")
        finally:
            inp.ROOT = original_root
            sys.path.remove(str(Path(__file__).resolve().parent.parent / "scripts"))

    def test_existing_project_rejected(self, projects_dir: Path):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
        import init_new_project as inp

        original_root = inp.ROOT
        inp.ROOT = projects_dir
        try:
            inp.init_project("Dup")
            with pytest.raises(FileExistsError):
                inp.init_project("Dup")
        finally:
            inp.ROOT = original_root
            sys.path.remove(str(Path(__file__).resolve().parent.parent / "scripts"))

    def test_slash_in_name_normalized(self, projects_dir: Path):
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
        import init_new_project as inp

        original_root = inp.ROOT
        inp.ROOT = projects_dir
        try:
            # 带 / 的名字 → 替换成 _
            target = inp.init_project("Suzuki/2026/Q4")
            assert target.exists()
            # 项目目录名应该是 "Suzuki_2026_Q4"
            assert "Suzuki_2026_Q4" in target.name
        finally:
            inp.ROOT = original_root
            sys.path.remove(str(Path(__file__).resolve().parent.parent / "scripts"))