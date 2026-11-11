"""scripts/init_new_project.py — 一键生成新项目骨架。

用法:
  python scripts/init_new_project.py "Suzuki_2026Q4_张三"

会创建:
  projects/Suzuki_2026Q4_张三/
    ├── Suzuki_2026Q4_张三.xlsx       (4 sheet + 公式 + 按钮, ⚙️ 配置 和 🧪 记录 留空)
    ├── 跑_BO_点这里.bat              (应急入口)
    └── LOG.txt                        (空,系统会写)

不依赖任何外部数据,完全用 templates/project_template.xlsx。
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# ⚠️ bat_templates 在 scripts/ 下 — 保证能 import 到(从任意目录运行都行)
_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))


def init_project(project_name: str) -> Path:
    """初始化新项目。返回项目目录。"""
    if not project_name or not project_name.strip():
        raise ValueError("项目名不能为空")

    # 安全检查:避免路径穿越
    safe_name = project_name.strip().replace("/", "_").replace("\\", "_")
    if ".." in safe_name or safe_name.startswith("."):
        raise ValueError(f"项目名包含非法字符: {project_name!r}")

    target_dir = ROOT / "projects" / safe_name
    target_xlsx = target_dir / f"{safe_name}.xlsx"

    if target_dir.exists():
        raise FileExistsError(f"项目已存在: {target_dir}")

    # 复制模板 — 缺失时自动生成(同事拿到包不用先跑别的脚本)
    template = ROOT / "templates" / "project_template.xlsx"
    if not template.exists():
        print("模板不存在,正在自动生成...")
        import subprocess
        subprocess.run(
            [sys.executable, str(Path(__file__).parent / "build_project_template.py")],
            check=True,
        )
    if not template.exists():
        raise FileNotFoundError(f"模板生成失败: {template}")

    target_dir.mkdir(parents=True)
    shutil.copy2(template, target_xlsx)

    # 写 _status 项目名 + 项目目录 + 初始状态
    from openpyxl import load_workbook
    wb = load_workbook(str(target_xlsx))
    ws_status = wb["_status"]
    ws_status["B2"] = "awaiting_config"  # 等用户填 ⚙️ 配置
    ws_status["B3"] = "1/4"
    ws_status["B4"] = "切到 ⚙️ 配置,填写因子和目标"
    ws_status["B5"] = safe_name
    ws_status["B6"] = str(target_dir)
    ws_status["B7"] = 0
    ws_status["B8"] = 0
    ws_status["B9"] = ""
    ws_status["B10"] = "⏳ 暂无收敛数据"
    ws_status["B11"] = ""
    wb.save(str(target_xlsx))
    wb.close()

    # 写 3 个独立 bat(⚠️ 不能用一个 bat + 参数 — 表格软件打开 .bat 时参数会丢)
    from bat_templates import write_bats

    write_bats(target_dir)

    # 创建空 LOG.txt(系统会写)
    (target_dir / "LOG.txt").touch()

    # ⚠️ 替换 🚀 开始 sheet 的按钮占位符 → 绝对路径
    # WPS/Excel 解析相对路径用表格软件当前工作目录,不是 xlsx 目录 → 相对路径必失效
    bat_paths = {
        "{{BO_BAT}}": (target_dir / "跑_BO_点这里.bat").resolve(),
        "{{REFRESH_BAT}}": (target_dir / "刷新看板.bat").resolve(),
        "{{SYNC_BAT}}": (target_dir / "生成记录表头.bat").resolve(),
    }
    _replace_button_placeholder(target_xlsx, bat_paths)

    return target_dir


def _replace_button_placeholder(xlsx_path: Path, bat_paths: dict) -> int:
    """把 sheet 里的 {{BO_BAT}} / {{REFRESH_BAT}} / {{SYNC_BAT}} 占位符
    替换成各自的绝对路径。

    ⚠️ 每个按钮指向**独立的 bat 文件** — 表格软件打开 .bat 时参数会丢,
    所以路径里绝不能带 " refresh" / " sync" 这样的参数。

    ⚠️ 扫**所有 sheet** — 2026-11 起 ⚙️ 配置 sheet 底部也有按钮。
    """
    from openpyxl import load_workbook

    wb = load_workbook(str(xlsx_path))
    replaced = 0
    for sheet_name in wb.sheetnames:
        if sheet_name == "_status":
            continue
        ws = wb[sheet_name]
        for r in range(1, 60):
            for c in range(1, 12):
                cell = ws.cell(row=r, column=c)
                v = cell.value
                if not (isinstance(v, str) and "{{" in v):
                    continue
                hit = False
                for placeholder, path in bat_paths.items():
                    if placeholder in v:
                        bat_str = str(path)
                        cell.value = v.replace(placeholder, bat_str)
                        # ⚠️ hyperlink 对象的 target 必须是**纯文件路径**(不带参数)
                        if cell.hyperlink is not None and cell.hyperlink.target:
                            for ph in bat_paths:
                                if ph in cell.hyperlink.target:
                                    cell.hyperlink.target = bat_str
                        hit = True
                        break
                if hit:
                    replaced += 1
    if replaced:
        wb.save(str(xlsx_path))
    wb.close()
    return replaced


def main():
    if len(sys.argv) < 2:
        print("用法: python scripts/init_new_project.py <项目名>")
        print('  例: python scripts/init_new_project.py "Suzuki_2026Q4_张三"')
        print()
        print("如果报 'No module named openpyxl',你的 python 路径不对。")
        print("试这些 python(任选):")
        print("  venv\\Scripts\\python.exe      (装了 install\\install.bat 就有)")
        print("  conda run -n base python      (Miniforge/Anaconda)")
        print("  python                        (PATH 上的)")
        print()
        print("或者直接双击 init_new_project.bat(自动检测 python)")
        sys.exit(1)
    name = sys.argv[1]

    # 友好错误:依赖缺失
    try:
        from openpyxl import load_workbook  # noqa: F401
    except ImportError:
        print("❌ openpyxl 没装。")
        print("   跑: python -m pip install openpyxl pydantic")
        print("   或者: 双击 install\\install.bat 重装环境")
        sys.exit(1)

    try:
        target = init_project(name)
    except (FileExistsError, ValueError, FileNotFoundError) as e:
        print(f"❌ {e}")
        sys.exit(1)

    print(f"✅ 项目已创建: {target}")
    print(f"   xlsx: {target / (name + '.xlsx')}")
    print(f"   bat:  {target / '跑_BO_点这里.bat'}")
    print()
    print("下一步:")
    print(f"  1. 打开 {target / (name + '.xlsx')}")
    print(f"  2. 填 ⚙️ 配置(因子 / 目标)")
    print(f"  3. 填 🧪 记录(已有的实验数据)")
    print(f"  4. 点 🚀 开始 sheet 的 [🚀 启动 BO]")


if __name__ == "__main__":
    main()