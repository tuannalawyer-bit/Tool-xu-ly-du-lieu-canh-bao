@echo off
chcp 65001 >nul
title Tool Phan Tich Ton Kho & Tong Quan Chuoi

set "SCRIPT_DIR=%~dp0"
set "PYTHON_PATH=%SCRIPT_DIR%python_embed\python.exe"
set "LIBS_DIR=%SCRIPT_DIR%libs"

set "PYTHONPATH=%LIBS_DIR%;%PYTHONPATH%"

echo Khoi dong Tool...
if exist "%PYTHON_PATH%" (
    start "" "%SCRIPT_DIR%python_embed\pythonw.exe" "%SCRIPT_DIR%tool_do_du_lieu.py"
) else (
    start "" python "%SCRIPT_DIR%tool_do_du_lieu.py"
)
