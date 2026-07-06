@echo off
chcp 65001 >nul
title Tool Do Du Lieu - Phan Tich Ton Kho v1.0.2

echo ========================================================
echo  Tool Do Du Lieu Phan Tich Ton Kho v1.0.2 (Chính thức)
echo ========================================================
echo.

:: Kiểm tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [LỖI] Không tìm thấy Python. Vui lòng cài Python 3.10+
    echo Tải Python tại: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Kiểm tra và nạp thư viện offline nếu thiếu để tránh treo mạng do proxy
echo [INFO] Kiểm tra thư viện...
python -c "import openpyxl, pyxlsb, tkinter" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Phát hiện thiếu thư viện hệ thống. Đang tìm thư viện portable offline...
    if exist "%~dp0Tool_PhanTichTonKho_Portable\libs" (
        set PYTHONPATH=%~dp0Tool_PhanTichTonKho_Portable\libs;%PYTHONPATH%
        echo [INFO] Đã nạp thành công thư viện offline từ thư mục portable!
    ) else (
        echo [CẢNH BÁO] Không tìm thấy thư mục libs offline. 
        echo Đang thử cài đặt online qua pip (giới hạn thời gian chờ)...
        python -m pip install openpyxl pyxlsb --quiet --timeout 5
        if errorlevel 1 (
            echo [LỖI] Không thể tự động cài đặt thư viện do lỗi mạng/proxy.
            echo Vui lòng sử dụng file Chay_Tool.bat bên trong thư mục 'Tool_PhanTichTonKho_Portable'.
            pause
            exit /b 1
        )
    )
)

echo [INFO] Đang khởi động...
python "%~dp0tool_do_du_lieu.py"

if errorlevel 1 (
    echo.
    echo [LỖI] Tool kết thúc với lỗi. Xem chi tiết bên trên.
    pause
)
