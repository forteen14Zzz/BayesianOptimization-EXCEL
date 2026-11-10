@echo off
chcp 65001 >nul
setlocal

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
for %%I in ("%PROJECT_DIR%") do set "PROJECT_NAME=%%~nxI"

cd /d "%PROJECT_DIR%\..\.."

echo === Refreshing dashboard for project %PROJECT_NAME% ===
echo.

set "PYTHONPATH=%PROJECT_DIR%\..\.."
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
set "MINIFORGE=C:\Users\user\Miniforge3\python.exe"
if exist "%MINIFORGE%" (
    set "PY=%MINIFORGE%"
    goto :eof
)
if exist "venv\Scripts\python.exe" (
    call "venv\Scripts\activate.bat"
    set "PY=python"
    goto :eof
)
echo [WARN] No Miniforge3 and no venv, using global python.
set "PY=python"
goto :eof
