@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
chcp 65001 >nul
title SOAR Desktop MCP - Installer V1.0

set "HERE=%~dp0"
pushd "%HERE%"

rem --- 0. Path safety check ---
echo "%HERE%" | findstr /R "[()!^&]" >nul 2>&1
if not errorlevel 1 (
    echo.
    echo ============================================================
    echo   [X] ERROR: Install path contains special characters
    echo       Please move to a simple path like D:\soar-desktop-mcp\
    echo ============================================================
    pause
    popd
    exit /b 1
)

echo.
echo ============================================================
echo   SOAR Desktop MCP - V1.0 Installer
echo   Open Source - Control any Windows app via Accessibility Tree
echo ============================================================
echo.

rem --- 1. Python check ---
echo [1/4] Checking Python ...
py -3.10 -c "import sys; assert sys.version_info[:2]==(3,10)" 2>nul
if not errorlevel 1 (
    set "PYCMD=py -3.10"
    goto :python_ok
)
py -3 -c "import sys; assert sys.version_info >= (3,10)" 2>nul
if not errorlevel 1 (
    set "PYCMD=py -3"
    goto :python_ok
)
where python >nul 2>&1
if errorlevel 1 (
    echo   [X] Python not found. Please install Python 3.10+ from
    echo       https://www.python.org/downloads/
    goto :fail
)
for /f "tokens=2" %%V in ('python --version 2^>^&1') do set "PYVER=%%V"
echo   Warning: py launcher not found, falling back to python
set "PYCMD=python"
:python_ok
echo   Using: !PYCMD!

rem --- 2. venv ---
echo [2/4] Creating virtual environment ...
set "VENV_DIR=%HERE%venv"
if exist "%VENV_DIR%" goto :venv_exists
!PYCMD! -m venv "%VENV_DIR%"
if errorlevel 1 goto :fail
goto :venv_done

:venv_exists
echo   venv exists - verifying health ...
call "%VENV_DIR%\Scripts\activate.bat" >nul 2>&1
python -c "import sys" >nul 2>&1
if not errorlevel 1 (
    echo   venv OK.
    goto :venv_done
)
echo   venv corrupted. Rebuilding ...
rmdir /s /q "%VENV_DIR%" >nul 2>&1
!PYCMD! -m venv "%VENV_DIR%"
if errorlevel 1 goto :fail
echo   venv rebuilt successfully.

:venv_done

rem --- 3. pip + dependencies ---
echo [3/4] Installing dependencies ...
call "%HERE%venv\Scripts\activate.bat"
python -m pip install --upgrade pip --timeout 120 >nul
python -m pip install --timeout 120 -r "%HERE%requirements.txt"
if errorlevel 1 goto :fail

rem --- 4. Verify import ---
echo [4/4] Verifying installation ...
python -c "import uiautomation; import mcp; print('  uiautomation OK'); print('  mcp OK')"
if errorlevel 1 (
    echo   [X] Import check failed.
    goto :fail
)

echo.
echo ============================================================
echo   Installation complete!
echo ============================================================
echo.
echo   To launch:   %HERE%run_mcp.bat
echo.
echo   Add to Claude Cowork / Claude Code config:
echo.
echo   "soar-desktop": {
echo     "command": "%HERE%run_mcp.bat"
echo   }
echo.
pause
popd
exit /b 0

:fail
echo.
echo ============================================================
echo   [X] Installation FAILED - see messages above
echo ============================================================
pause
popd
exit /b 1
