@echo off
chcp 65001 >nul
setlocal

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
for %%I in ("%PROJECT_DIR%") do set "PROJECT_NAME=%%~nxI"

cd /d "%PROJECT_DIR%\..\.."

echo === Syncing record headers for project %PROJECT_NAME% ===
echo.

set "PYTHONPATH=%PROJECT_DIR%\..\.."
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
