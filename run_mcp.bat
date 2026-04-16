@echo off
setlocal ENABLEEXTENSIONS
chcp 65001 >nul
title SOAR Desktop MCP

set "HERE=%~dp0"
pushd "%HERE%"

if not exist "%HERE%venv\Scripts\activate.bat" (
    echo venv not found. Run install.bat first.
    pause
    exit /b 1
)

call "%HERE%venv\Scripts\activate.bat"
python "%HERE%server.py" %*

popd
endlocal
