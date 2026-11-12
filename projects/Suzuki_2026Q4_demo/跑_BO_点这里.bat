@echo off
chcp 65001 >nul
setlocal

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
for %%I in ("%PROJECT_DIR%") do set "PROJECT_NAME=%%~nxI"

cd /d "%PROJECT_DIR%\..\.."

echo === Starting BO for project %PROJECT_NAME% (ask + refresh) ===
echo.

set "PYTHONPATH=%PROJECT_DIR%\..\.."
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

:pick_python
REM Find a python that has bofire installed.
REM Order: local venv -> conda -> PATH. Never hardcode a user path.
set "PY="

REM 1) project-local venv (created by install\install.bat)
if exist "venv\Scripts\python.exe" (
    set "PY=venv\Scripts\python.exe"
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
echo   - Option A: run install\install.bat to create venv\  ^(recommended^)
echo   - Option B: install Miniforge/Anaconda, then: conda install bofire
pause
exit /b 1
