@echo off
REM init_new_project.bat — 一键创建新项目(同事双击运行,无需知道 python 在哪)
REM 自动找可用的 python,如果没装包会提示

setlocal

cd /d %~dp0\..\

echo === BoFire Lab Portal - 新项目创建 ===
echo.

REM 找可用的 python(优先 hermes venv)
set PYTHON_EXE=
for %%P in (
    "C:\Users\user\AppData\Local\hermes\installs\b6d8a9a58d988e18\environments\e9f84127e71e450e986dc42b7947380e\venv\Scripts\python.exe"
    "C:\Users\user\AppData\Local\hermes\hermes-agent\venv\Scripts\python.exe"
    "python.exe"
) do (
    if exist %%~P (
        set PYTHON_EXE=%%~P
        goto :found
    )
)

:found
if "%PYTHON_EXE%"=="" (
    echo ❌ 找不到 Python
    pause
    exit /b 1
)

echo 使用 Python: %PYTHON_EXE%
%PYTHON_EXE% -c "import openpyxl, pandas, watchdog, pydantic" 2>nul
if errorlevel 1 (
    echo.
    echo ⚠️ Python 缺包,正在安装(可能需要 1-2 分钟)...
    %PYTHON_EXE% -m pip install openpyxl pandas watchdog pydantic
    if errorlevel 1 (
        echo ❌ 装包失败,联系 IT
        pause
        exit /b 1
    )
)

if "%~1"=="" (
    set /p NAME= "请输入项目名(例: Suzuki_2026Q4_张三): "
) else (
    set NAME=%~1
)

%PYTHON_EXE% scripts\init_new_project.py "%NAME%"
if errorlevel 1 (
    echo ❌ 创建失败
    pause
    exit /b 1
)

echo.
echo === 完成 ===
pause