@echo off
chcp 65001 >nul
title Khởi động Tool Phân Tích Tồn Kho v1.1.0 (Di động)

echo ==========================================================
echo   TOOL PHÂN TÍCH TỒN KHO v1.1.0 (PHIÊN BẢN DI ĐỘNG)
echo ==========================================================
echo.

:: Kiểm tra cài đặt Python cục bộ
python --version >nul 2>&1
if errorlevel 1 (
    echo [LỖI] Không tìm thấy môi trường Python cài trên máy tính này.
    echo Vui lòng cài đặt Python (3.10 trở lên) từ: https://www.python.org/downloads/
    echo Hoặc sao chép thư mục Python Embeddable đi kèm.
    echo.
    pause
    exit /b 1
)

:: Thiết lập PYTHONPATH để import thư viện đính kèm ngoại tuyến
set PYTHONPATH=%~dp0libs;%PYTHONPATH%

echo [INFO] Đang khởi chạy ứng dụng (Sử dụng thư viện offline trong thư mục libs)...
python "%~dp0tool_do_du_lieu.py"

if errorlevel 1 (
    echo.
    echo [LỖI] Chương trình kết thúc ngoài ý muốn. Vui lòng xem log lỗi ở trên.
    pause
)
