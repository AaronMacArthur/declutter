@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
    python "%~dp0declutter_app.py"
    goto :done
)

where py >nul 2>nul
if %errorlevel%==0 (
    py "%~dp0declutter_app.py"
    goto :done
)

echo Python was not found. Install Python 3 from https://www.python.org/downloads/windows/
pause

:done
