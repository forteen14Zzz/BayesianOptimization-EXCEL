"""portal/core/excel_control.py — Excel 自动保存/关闭/重开(COM)。

同事需求(M4 现场 zzz12):
点 [🚀 启动 BO] 后,必须手工关掉 Excel 才能让 Python 写盘,
跑完还要手工再打开 — 太烦。

解法:
1. 按钮触发 → bat 调 run.py
2. run.py 先通过 COM 找到打开这个 xlsx 的 Excel 实例
3. Save(保存) → Close(关闭) → Quit(退出 Excel)
4. 跑 BO(此时文件锁已释放)
5. 跑完 → 重新打开 xlsx(可选:启动 watcher 继续监听)

⚠️ 降级:
- 没有 pywin32 / COM 不可用 → 跳过关闭,照常跑 BO(同事手动关)
- Excel 没开这个文件 → 跳过关闭
"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any


def _get_excel() -> Any | None:
    """拿到 Excel COM 对象(优先已运行的实例)。失败返回 None。"""
    try:
        import win32com.client as win32
    except ImportError:
        return None
    try:
        try:
            # 已运行的实例
            return win32.GetActiveObject("Excel.Application")
        except Exception:
            # 没有就新建一个
            return win32.Dispatch("Excel.Application")
    except Exception:
        return None


def save_and_close(xlsx_path: Path, timeout: float = 10.0) -> dict:
    """保存并关闭打开着 xlsx_path 的 Excel。

    返回 {"ok": bool, "reason": str}
    """
    result = {"ok": False, "reason": "", "closed": False, "saved": False}

    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        result["reason"] = "file not found"
        return result

    excel = _get_excel()
    if excel is None:
        result["reason"] = "COM unavailable (pywin32 missing or Excel not installed)"
        return result

    try:
        # 找目标工作簿
        target = None
        for wb in excel.Workbooks:
            try:
                if Path(wb.FullName).resolve() == xlsx_path:
                    target = wb
                    break
            except Exception:
                continue

        if target is None:
            result["reason"] = "workbook not open in Excel"
            # 没有打开就不需要关
            result["ok"] = True
            return result

        # 保存
        target.Save()
        result["saved"] = True

        # 关闭
        target.Close(SaveChanges=False)  # 已经 Save 过了,不用再存
        result["closed"] = True

        # 如果没有其他工作簿了,把 Excel 也退出(释放锁)
        try:
            if excel.Workbooks.Count == 0:
                excel.Quit()
        except Exception:
            pass

        result["ok"] = True
        result["reason"] = "saved and closed"
        return result

    except Exception as e:
        result["reason"] = f"error: {e}"
        return result


def reopen(xlsx_path: Path, visible: bool = True) -> dict:
    """重新打开 xlsx(用 Excel 打开,窗口在前台)。"""
    result = {"ok": False, "reason": ""}

    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        result["reason"] = "file not found"
        return result

    excel = _get_excel()
    if excel is None:
        result["reason"] = "COM unavailable"
        return result

    try:
        excel.Visible = visible
        excel.Workbooks.Open(str(xlsx_path))
        result["ok"] = True
        result["reason"] = "reopened"
        return result
    except Exception as e:
        result["reason"] = f"error: {e}"
        return result


def wait_for_unlock(xlsx_path: Path, timeout: float = 15.0) -> bool:
    """等 Excel 释放文件锁(可写即可)。"""
    xlsx_path = Path(xlsx_path).resolve()
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            # 试着以写模式打开
            with open(xlsx_path, "r+b"):
                return True
        except PermissionError:
            time.sleep(0.5)
        except Exception:
            return True
    return False
