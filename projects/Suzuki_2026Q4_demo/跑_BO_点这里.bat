@echo off
REM Suzuki_2026Q4_demo\跑_BO_点这里.bat
REM 应急: 同事点这个手动跑 BO

cd /d %~dp0\..\..\
if not exist "venv\Scripts\activate.bat" (
    echo ❌ 没找到 venv, 请先跑 install\install.bat
    pause
    exit /b 1
)
call venv\Scripts\activate.bat

REM 默认 phase=ask; 用户可以传 phase: refresh / validate
if "%~1"=="" (
    set PHASE=ask
) else (
    set PHASE=%~1
)

echo === 启动 BO phase=%PHASE% project=Suzuki_2026Q4_demo ===
python portal\run.py --project projects\Suzuki_2026Q4_demo --phase %PHASE%
echo.
echo === 完成。日志在 projects\Suzuki_2026Q4_demo\LOG.txt ===
pause