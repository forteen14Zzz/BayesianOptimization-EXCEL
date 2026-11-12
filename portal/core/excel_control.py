"""portal/core/excel_control.py — WPS/Excel 自动保存/关闭/重开(COM)。

同事需求(M4 现场 zzz12 / 222):
点 [🚀 启动 BO] 后,必须手工关掉表格软件才能让 Python 写盘,
跑完还要手工再打开 — 太烦。

解法:
1. 按钮触发 → bat 调 run.py
2. run.py 通过 COM 找到打开着这个 xlsx 的表格软件实例
3. Save(保存) → Close(关闭) → Quit(退出)
4. 跑 BO(此时文件锁已释放)
5. 跑完 → 重新打开 xlsx

⚠️ 支持两种 COM:
  - **WPS 表格**(同事实际用的): ProgID = `Ket.Application`
  - **Microsoft Excel**: ProgID = `Excel.Application`
  先试 WPS,再试 Excel;两个都连不上就降级(同事手动关)。

⚠️ 降级:
- 没有 pywin32 / COM 不可用 → 跳过关闭,照常跑 BO
- 表格软件没开这个文件 → 跳过关闭
"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any


# 表格软件的 COM ProgID — 按优先级
# 1. WPS 表格(同事实际用)  2. Microsoft Excel
PROGIDS = ["Ket.Application", "Excel.Application"]


def _try_dispatch(progid: str) -> Any | None:
    """尝试连到某个 ProgID 的已运行实例,连不上就返回 None。"""
    try:
        import win32com.client as win32
    except ImportError:
        return None

    # 优先连已运行的实例(同事正在用的那个)
    try:
        return win32.GetActiveObject(progid)
    except Exception:
        pass
    # 连不上就试 Dispatch(会启动新实例)
    try:
        return win32.Dispatch(progid)
    except Exception:
        return None


def _get_app() -> Any | None:
    """拿到表格软件的 COM 对象。先试已运行实例,再试启动新实例。"""
    for progid in PROGIDS:
        app = _try_dispatch(progid)
        if app is not None:
            try:
                # 能访问 Workbooks 算可用
                _ = app.Workbooks
                return app
            except Exception:
                continue
    return None


def _app_name(app: Any) -> str:
    """猜这个 COM 对象是 WPS 还是 Excel(用于日志)。"""
    try:
        name = app.Name
    except Exception:
        name = ""
    if "ket" in str(name).lower() or "wps" in str(name).lower():
        return f"WPS({name})"
    return f"Excel({name})"


def save_and_close(xlsx_path: Path, timeout: float = 10.0) -> dict:
    """保存并关闭打开着 xlsx_path 的表格软件窗口。

    返回 {"ok": bool, "reason": str, "closed": bool, "saved": bool, "app": str}
    """
    result = {"ok": False, "reason": "", "closed": False, "saved": False, "app": ""}

    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        result["reason"] = "file not found"
        return result

    app = _get_app()
    if app is None:
        result["reason"] = "COM unavailable (pywin32 missing or no WPS/Excel)"
        return result
    result["app"] = _app_name(app)

    try:
        # 找目标工作簿
        target = None
        for wb in app.Workbooks:
            try:
                if Path(wb.FullName).resolve() == xlsx_path:
                    target = wb
                    break
            except Exception:
                continue

        if target is None:
            result["reason"] = "workbook not open in WPS/Excel"
            # 没打开就不需要关
            result["ok"] = True
            return result

        # 保存
        target.Save()
        result["saved"] = True

        # 关闭
        target.Close(SaveChanges=False)  # 已经 Save 过了,不用再存
        result["closed"] = True

        # 如果没有其他工作簿了,把表格软件也退出(彻底释放锁)
        try:
            if app.Workbooks.Count == 0:
                app.Quit()
        except Exception:
            pass

        result["ok"] = True
        result["reason"] = "saved and closed"
        return result

    except Exception as e:
        result["reason"] = f"error: {e}"
        return result


def reopen(xlsx_path: Path, visible: bool = True) -> dict:
    """重新打开 xlsx(用 WPS/Excel 打开,窗口在前台)。"""
    result = {"ok": False, "reason": "", "app": ""}

    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        result["reason"] = "file not found"
        return result

    app = _get_app()
    if app is None:
        result["reason"] = "COM unavailable"
        return result
    result["app"] = _app_name(app)

    try:
        app.Visible = visible
        app.Workbooks.Open(str(xlsx_path))
        result["ok"] = True
        result["reason"] = "reopened"
        return result
    except Exception as e:
        result["reason"] = f"error: {e}"
        return result


def wait_for_unlock(xlsx_path: Path, timeout: float = 15.0) -> bool:
    """等文件锁释放(可写即可)。"""
    xlsx_path = Path(xlsx_path).resolve()
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with open(xlsx_path, "r+b"):
                return True
        except PermissionError:
            time.sleep(0.5)
        except Exception:
            return True
    return False
