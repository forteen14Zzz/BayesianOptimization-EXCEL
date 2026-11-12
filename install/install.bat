@echo off
REM install.bat — BoFire Lab Portal 一次性装机脚本
REM 双击运行(无需 .exe)。会自动:
REM   1. 创建 venv
REM   2. 装 requirements.txt
REM   3. 配自启动(可选)

setlocal

cd /d "%~dp0\..\"

echo === BoFire Lab Portal 装机 ===
echo.

REM 1. 检查 Python
where python >nul 2>&1
if errorlevel 1 (
    echo ❌ 找不到 Python。请 IT 先装 Python 3.11 或 3.12 并加入 PATH。
    pause
    exit /b 1
)

REM 2. 创建 venv
if not exist "venv\" (
    echo [1/4] 创建 venv...
    python -m venv venv
    if errorlevel 1 (
        echo ❌ venv 创建失败
        pause
        exit /b 1
    )
) else (
    echo [1/4] venv 已存在,跳过
)

REM 3. 激活并装包
echo [2/4] 装依赖(bofire==0.4.1, botorch, watchdog, ... 大约 3-5 分钟)...
call "venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r "install\requirements.txt"
if errorlevel 1 (
    echo ❌ 装包失败,可能是 torch 装不上,联系 IT
    pause
    exit /b 1
)

REM 4. 验证 — 检查关键模块,给出友好提示
echo [3/4] 验证...
python -c "
import sys
missing = []
for mod in ['bofire', 'botorch', 'torch', 'openpyxl', 'pydantic', 'pandas', 'watchdog', 'cvxpy', 'pymoo', 'shap']:
    try:
        __import__(mod)
    except ImportError:
        missing.append(mod)
if missing:
    print('❌ 缺这些包:', missing)
    sys.exit(1)
else:
    import bofire, torch
    print(f'  bofire {bofire.__version__}, torch {torch.__version__}, OK')
" || (
    echo ❌ 验证失败,联系 IT
    pause
    exit /b 1
)

REM 5. 配自启动(询问)
set /p autostart= "[4/4] 是否配开机自启动 watchdog? (y/n, 默认 n): "
if /I "%autostart%"=="y" (
    call "install\setup_autostart.bat"
) else (
    echo (跳过自启动,用户可手动双击 start_watcher.bat)
)

echo.
echo ✅ 装机完成!
echo 接下来:
echo   1. 复制 templates\project_template.xlsx 到 projects\新建项目\
echo   2. 改名 + 双击打开,跟着 Excel 里 🚀 开始 sheet 引导走
echo.
pause