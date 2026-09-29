@echo off
REM setup_autostart.bat — 把 start_watcher.bat 复制到 Windows 启动项
REM 这样同事开机后 watchdog 自动启动

setlocal

set STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set WATCHER_BAT=%~dp0\..\start_watcher.bat"

REM 创建 start_watcher.bat(如果还没)
if not exist "%WATCHER_BAT%" (
    (
        echo @echo off
        echo cd /d "%~dp0"
        echo call venv\Scripts\activate.bat
        echo python portal\watcher.py --project "%~dp0projects\"
        echo pause
    ) > "%WATCHER_BAT%"
)

REM 创建快捷方式到启动项
set SHORTCUT=%STARTUP_DIR%\BoFireLabPortal_watcher.lnk

powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell; ^
     $sc = $ws.CreateShortcut('%SHORTCUT%'); ^
     $sc.TargetPath = '%WATCHER_BAT%'; ^
     $sc.WorkingDirectory = '%~dp0\..'; ^
     $sc.WindowStyle = 7; ^
     $sc.Save()"

echo ✅ 启动项已配置: %SHORTCUT%
echo (重启后 watchdog 自动启动,无需手动)
endlocal