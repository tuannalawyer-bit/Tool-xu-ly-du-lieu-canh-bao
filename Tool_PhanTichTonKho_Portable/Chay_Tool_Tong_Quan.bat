@echo off
chcp 65001 >nul
title Khởi động Tool Tổng Quan Chuỗi (Portable)

echo ==========================================================
echo   TOOL TỔNG QUAN CHUỖI PHÂN TÍCH TỒN KHO (PHIÊN BẢN DI ĐỘNG)
echo ==========================================================
echo.

:: Thiết lập đường dẫn thư viện ngoài
set "PYTHONPATH=%~dp0libs;%PYTHONPATH%"

:: Khởi chạy ẩn màn hình console đen bằng pythonw.exe nhúng
start "" "%~dp0python_embed\pythonw.exe" "%~dp0tool_tong_quan.py"
exit
