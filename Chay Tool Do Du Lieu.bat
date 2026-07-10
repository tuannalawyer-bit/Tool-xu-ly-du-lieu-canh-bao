@echo off
chcp 65001 >nul
title Tool Phan Tich Ton Kho

set "SCRIPT_DIR=%~dp0"
set "PYTHON_PATH=%SCRIPT_DIR%Tool_PhanTichTonKho_Portable\python_embed\python.exe"
set "LIBS_DIR=%SCRIPT_DIR%Tool_PhanTichTonKho_Portable\libs"

set "PYTHONPATH=%LIBS_DIR%;%PYTHONPATH%"

echo Starting Tool...
if exist "%PYTHON_PATH%" (
    "%PYTHON_PATH%" "%SCRIPT_DIR%tool_do_du_lieu.py"
) else (
    python "%SCRIPT_DIR%tool_do_du_lieu.py"
)

if errorlevel 1 (
    echo.
    echo [ERROR] Tool exited with error.
    pause
)
