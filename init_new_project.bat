@echo off
chcp 65001 >nul
REM ============================================================
REM  init_new_project.bat - create a new BoFire Lab Portal project
REM
REM  HOW TO USE:
REM    1. Double-click this file
REM    2. Type a project name (ASCII recommended, e.g. Suzuki_2026Q4_Lisi)
REM    3. Press Enter
REM    4. Open the generated .xlsx in projects\<name>\
REM
REM  Or from command line:
REM    init_new_project.bat Suzuki_2026Q4_Lisi
REM
REM  NOTE: All messages here are ASCII on purpose.
REM  Chinese characters inside a .bat break cmd.exe parsing
REM  when combined with ( ) parentheses.
REM ============================================================

setlocal

cd /d "%~dp0"

echo ============================================================
echo   BoFire Lab Portal - Create New Project
echo ============================================================
echo.

REM ---- find a usable python ----
REM priority: project venv > conda base > PATH
REM (never hardcode a user-specific path - this file ships to other machines)
set PYTHON_EXE=

if exist "venv\Scripts\python.exe" (
    set "PYTHON_EXE=venv\Scripts\python.exe"
    goto :found
)

REM 2) conda (bofire is often installed in conda base / miniforge)
REM    ⚠️ `where conda` 的 ERRORLEVEL in "not found" case can still be 0,
REM    so check `conda --version` output instead.
for /f "tokens=*" %%V in ('conda --version 2^>nul') do (
    if not "%%V"=="" (
        for /f "tokens=*" %%P in ('conda run -n base python -c "import sys;print(sys.executable)" 2^>nul') do (
            if not "%%P"=="" (
                set "PYTHON_EXE=%%P"
                goto :found
            )
        )
    )
)

where python >nul 2>&1
if %ERRORLEVEL%==0 (
    set "PYTHON_EXE=python"
    goto :found
)

:found
if "%PYTHON_EXE%"=="" (
    echo [FAIL] No python found.
    echo   - Option A: run install\install.bat first  ^(recommended^)
    echo   - Option B: install Miniforge/Anaconda, make sure conda is on PATH.
    pause
    exit /b 1
)

echo Using python: %PYTHON_EXE%
echo.

REM ---- check core deps ----
%PYTHON_EXE% -c "import openpyxl, pandas, watchdog, pydantic" 2>nul
if errorlevel 1 (
    echo [WARN] Missing packages. Installing now, 1-2 minutes...
    %PYTHON_EXE% -m pip install openpyxl pandas watchdog pydantic
    if errorlevel 1 (
        echo [FAIL] Install failed. Please contact IT.
        pause
        exit /b 1
    )
    echo.
)

REM ---- check bofire (optional) ----
%PYTHON_EXE% -c "import bofire" 2>nul
if errorlevel 1 (
    echo [WARN] bofire not installed - BO will use random sampling.
    echo        For smart GP-BO: pip install "bofire[optimization]==0.5.0"
    echo.
)

REM ---- get project name ----
if "%~1"=="" (
    set /p NAME= "Project name (e.g. Suzuki_2026Q4_Lisi): "
) else (
    set NAME=%~1
)

if "%NAME%"=="" (
    echo [FAIL] No project name given.
    pause
    exit /b 1
)

echo.
echo Creating project: %NAME%
echo.

%PYTHON_EXE% scripts/init_new_project.py "%NAME%"
if errorlevel 1 (
    echo.
    echo [FAIL] Create failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   DONE. Next steps:
echo     1. Open the .xlsx file in projects\%NAME%\
echo     2. Fill in the config sheet (factors and target)
echo     3. Fill in the record sheet (data you already have)
echo     4. Click the green BO button on the start sheet
echo ============================================================
echo.
pause
