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
REM priority: Miniforge3 (bofire installed) > hermes venv > PATH
set PYTHON_EXE=
if exist "C:\Users\user\Miniforge3\python.exe" (
    set "PYTHON_EXE=C:\Users\user\Miniforge3\python.exe"
    goto :found
)
if exist "C:\Users\user\AppData\Local\hermes\hermes-agent\venv\Scripts\python.exe" (
    set "PYTHON_EXE=C:\Users\user\AppData\Local\hermes\hermes-agent\venv\Scripts\python.exe"
    goto :found
)
for %%P in ("python.exe") do (
    if exist %%~P (
        set "PYTHON_EXE=%%~P"
        goto :found
    )
)

:found
if "%PYTHON_EXE%"=="" (
    echo [FAIL] No python found. Please contact IT.
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
    echo        For smart GP-BO: %PYTHON_EXE% -m pip install bofire==0.4.1 botorch
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
