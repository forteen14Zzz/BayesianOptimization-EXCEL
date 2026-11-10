"""生成项目的 bat 文件(3 个独立 bat,不用参数)。

⚠️ 为什么不用参数:
Excel 的 HYPERLINK 打开 .bat 时,**参数可能丢失**(实测同事 222 项目
点 [📋 生成记录表头] 却跑了 ask)。所以每个按钮对应一个独立的 bat 文件,
路径里没有参数,100% 可靠。

生成的 bat:
  跑_BO_点这里.bat      — ask + refresh(默认)
  刷新看板.bat          — 只把 cache 合并到看板
  生成记录表头.bat      — 只按配置重建 🧪 记录 表头
"""
from __future__ import annotations

from pathlib import Path


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

:pick_python
REM Choose python: Miniforge3 first (bofire installed), venv, else global
REM Windows where python priority is wrong; explicit Miniforge3 path
set "MINIFORGE=%LOCALAPPDATA%\Miniforge3\python.exe"
if exist "%MINIFORGE%" (
    set "PY=%MINIFORGE%"
    goto :eof
)
if exist "venv\\Scripts\\python.exe" (
    call "venv\\Scripts\\activate.bat"
    set "PY=python"
    goto :eof
)
echo [WARN] No Miniforge3 and no venv, using global python.
set "PY=python"
goto :eof
"""

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

:pick_python
set "MINIFORGE=%LOCALAPPDATA%\Miniforge3\python.exe"
if exist "%MINIFORGE%" (
    set "PY=%MINIFORGE%"
    goto :eof
)
if exist "venv\\Scripts\\python.exe" (
    call "venv\\Scripts\\activate.bat"
    set "PY=python"
    goto :eof
)
echo [WARN] No Miniforge3 and no venv, using global python.
set "PY=python"
goto :eof
"""

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

:pick_python
set "MINIFORGE=%LOCALAPPDATA%\Miniforge3\python.exe"
if exist "%MINIFORGE%" (
    set "PY=%MINIFORGE%"
    goto :eof
)
if exist "venv\\Scripts\\python.exe" (
    call "venv\\Scripts\\activate.bat"
    set "PY=python"
    goto :eof
)
echo [WARN] No Miniforge3 and no venv, using global python.
set "PY=python"
goto :eof
"""


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
