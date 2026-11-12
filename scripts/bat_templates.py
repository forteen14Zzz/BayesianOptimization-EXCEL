"""生成项目的 bat 文件(3 个独立 bat,不用参数)。

⚠️ 为什么不用参数:
Excel/WPS 的 HYPERLINK 打开 .bat 时,**参数可能丢失**(实测同事 222 项目
点 [📋 生成记录表头] 却跑了 ask)。所以每个按钮对应一个独立的 bat 文件,
路径里没有参数,100% 可靠。

生成的 bat:
  跑_BO_点这里.bat      — ask + refresh(默认)
  刷新看板.bat          — 只把 cache 合并到看板
  生成记录表头.bat      — 只按配置重建 🧪 记录 表头

⚠️ **不要硬编码任何用户名/绝对路径** — 别人 clone 到别的机器也要能用。
python 靠 :pick_python 标签自动探测(venv → conda → PATH)。
"""
from __future__ import annotations

from pathlib import Path


# ---------- 共用的 python 探测片段 ----------
# ⚠️ 顺序有讲究:
#   1. 项目本地 venv(装机脚本建的,最可靠)
#   2. conda 环境(bofire 常装在 conda base / miniforge)
#   3. PATH 上的 python
# 全部失败 → 明确报错,不用错的 python 静默跑
_PICK_PYTHON = """:pick_python
REM Find a python that has bofire installed.
REM Order: local venv -> conda -> PATH. Never hardcode a user path.
set "PY="

REM 1) project-local venv (created by install\\install.bat)
if exist "venv\\Scripts\\python.exe" (
    set "PY=venv\\Scripts\\python.exe"
    goto :eof
)

REM 2) conda (bofire is often installed in conda base / miniforge)
REM    ⚠️ `where conda` 的 ERRORLEVEL 在找不到时仍可能是 0(实测),
REM    所以必须用 `conda --version` 的输出来判断 conda 是否真的可用。
for /f "tokens=*" %%V in ('conda --version 2^>nul') do (
    if not "%%V"=="" (
        for /f "tokens=*" %%P in ('conda run -n base python -c "import sys;print(sys.executable)" 2^>nul') do (
            if not "%%P"=="" (
                set "PY=%%P"
                goto :eof
            )
        )
    )
)

REM 3) any python on PATH
where python >nul 2>&1
if %ERRORLEVEL%==0 (
    set "PY=python"
    goto :eof
)

echo [ERROR] No python found.
REM ⚠️ 括号必须转义 ^( ^) — cmd 会把裸 ) 当成块结束符
echo   - Option A: run install\\install.bat to create venv\\  ^(recommended^)
echo   - Option B: install Miniforge/Anaconda, then: conda install bofire
pause
exit /b 1
"""


# ---------- bat 模板 ----------
BAT_MAIN = """@echo off
chcp 65001 >nul
setlocal

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
for %%I in ("%PROJECT_DIR%") do set "PROJECT_NAME=%%~nxI"

cd /d "%PROJECT_DIR%\\..\\.."

echo === Starting BO for project %PROJECT_NAME% (ask + refresh) ===
echo.

set "PYTHONPATH=%PROJECT_DIR%\\..\\.."
call :pick_python
if "%PY%"=="" exit /b 1

echo Step 1/2: Running ask...
"%PY%" "portal/run.py" --project "projects/%PROJECT_NAME%" --phase ask
if errorlevel 1 (
    echo [FAIL] Ask failed.
    goto end
)
echo.

echo Step 2/2: Running refresh...
"%PY%" "portal/run.py" --project "projects/%PROJECT_NAME%" --phase refresh
if errorlevel 1 (
    echo [FAIL] Refresh failed.
    goto end
)

echo.
echo [OK] BO done.
goto end

:end
echo.
echo Log: projects/%PROJECT_NAME%/LOG.txt
pause
exit /b 0

""" + _PICK_PYTHON

BAT_REFRESH = """@echo off
chcp 65001 >nul
setlocal

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
for %%I in ("%PROJECT_DIR%") do set "PROJECT_NAME=%%~nxI"

cd /d "%PROJECT_DIR%\\..\\.."

echo === Refreshing dashboard for project %PROJECT_NAME% ===
echo.

set "PYTHONPATH=%PROJECT_DIR%\\..\\.."
call :pick_python
if "%PY%"=="" exit /b 1

"%PY%" "portal/run.py" --project "projects/%PROJECT_NAME%" --phase refresh
if errorlevel 1 (
    echo [FAIL] Refresh failed.
    goto end
)

echo.
echo [OK] Dashboard refreshed.
goto end

:end
echo.
echo Log: projects/%PROJECT_NAME%/LOG.txt
pause
exit /b 0

""" + _PICK_PYTHON

BAT_SYNC = """@echo off
chcp 65001 >nul
setlocal

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
for %%I in ("%PROJECT_DIR%") do set "PROJECT_NAME=%%~nxI"

cd /d "%PROJECT_DIR%\\..\\.."

echo === Syncing record headers for project %PROJECT_NAME% ===
echo.

set "PYTHONPATH=%PROJECT_DIR%\\..\\.."
call :pick_python
if "%PY%"=="" exit /b 1

"%PY%" "portal/run.py" --project "projects/%PROJECT_NAME%" --phase validate
if errorlevel 1 (
    echo [FAIL] Sync failed.
    goto end
)

echo.
echo [OK] Record headers synced. Go fill data in the record sheet.
goto end

:end
pause
exit /b 0

""" + _PICK_PYTHON


def write_bats(project_dir: Path) -> dict[str, Path]:
    """在项目目录写 3 个 bat(CRLF)。返回 {名字: 路径}。"""
    project_dir.mkdir(parents=True, exist_ok=True)

    files = {
        "跑_BO_点这里.bat": BAT_MAIN,
        "刷新看板.bat": BAT_REFRESH,
        "生成记录表头.bat": BAT_SYNC,
    }

    result = {}
    for name, content in files.items():
        path = project_dir / name
        # ⚠️ 必须 CRLF — cmd.exe 不解析 LF-only 的 bat
        data = content.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8")
        path.write_bytes(data)
        result[name] = path
    return result
