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

    # 复制模板
    template = ROOT / "templates" / "project_template.xlsx"
    if not template.exists():
        raise FileNotFoundError(f"模板不存在: {template}, 请先跑 scripts/build_project_template.py")

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

    # 复制应急 bat(从 demo 项目,已经写好"自动从目录名推断项目名")
    bat_template = ROOT / "projects" / "Suzuki_2026Q4_demo" / "跑_BO_点这里.bat"
    if bat_template.exists():
        target_bat = target_dir / f"跑_BO_点这里.bat"
        shutil.copy2(bat_template, target_bat)
        # ⚠️ Windows cmd 必须 CRLF 终止符解析 bat 文件; git-bash / write_file
        # 默认是 LF,需要转 CRLF。用 Python 直接转,避免依赖 unix2dos。
        try:
            content = target_bat.read_bytes()
            # 把 LF-only 转 CRLF(已 CRLF 的跳过)
            if b"\r\n" not in content:
                content = content.replace(b"\n", b"\r\n")
                target_bat.write_bytes(content)
        except Exception:
            pass  # 转换失败不致命,Windows 下手动转也行

    # 创建空 LOG.txt(系统会写)
    (target_dir / "LOG.txt").touch()

    return target_dir


def main():
    if len(sys.argv) < 2:
        print("用法: python scripts/init_new_project.py <项目名>")
        print('  例: python scripts/init_new_project.py "Suzuki_2026Q4_张三"')
        print()
        print("如果报 'No module named openpyxl',你的 python 路径不对。")
        print("试这些 python(任选):")
        print("  /c/Users/user/AppData/Local/hermes/hermes-agent/venv/Scripts/python.exe")
        print("  /c/Users/user/Miniforge3/python.exe")
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