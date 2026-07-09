@echo off
chcp 65001 >nul
title Tool Tong Quan Chuoi - Phan Tich Ton Kho

echo ========================================================
echo  Tool Tong Quan Chuoi Phan Tich Ton Kho (Chính thức)
echo ========================================================
echo.

:: Kiểm tra Python hệ thống, nếu thiếu thì thử dùng Python nhúng từ bản Portable
set "PYTHON_CMD=python"
python --version >nul 2>&1
if errorlevel 1 (
    if exist "%~dp0Tool_PhanTichTonKho_Portable\python_embed\python.exe" (
        set "PYTHON_CMD=%~dp0Tool_PhanTichTonKho_Portable\python_embed\python.exe"
        echo [INFO] Không tìm thấy Python hệ thống. Sử dụng Python tích hợp từ thư mục portable...
    ) else (
        echo [LỖI] Không tìm thấy Python. Vui lòng cài Python 3.10+
        echo Tải Python tại: https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

:: Kiểm tra và nạp thư viện offline nếu thiếu để tránh treo mạng do proxy
echo [INFO] Kiểm tra thư viện...
"%PYTHON_CMD%" -c "import openpyxl, tkinter" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Phát hiện thiếu thư viện hệ thống. Đang tìm thư viện portable offline...
    if exist "%~dp0Tool_PhanTichTonKho_Portable\libs" (
        set "PYTHONPATH=%~dp0Tool_PhanTichTonKho_Portable\libs;%PYTHONPATH%"
        echo [INFO] Đã nạp thành công thư viện offline từ thư mục portable!
    ) else (
        echo [CẢNH BÁO] Không tìm thấy thư mục libs offline. 
        echo Đang thử cài đặt online qua pip - giới hạn thời gian chờ...
        "%PYTHON_CMD%" -m pip install openpyxl --quiet --timeout 5
        if errorlevel 1 (
            echo [LỖI] Không thể tự động cài đặt thư viện do lỗi mạng/proxy.
            echo Vui lòng sử dụng file Chay_Tool.bat bên trong thư mục 'Tool_PhanTichTonKho_Portable'.
            pause
            exit /b 1
        )
    )
)

echo [INFO] Đang khởi động...
"%PYTHON_CMD%" "%~dp0tool_tong_quan.py"

if errorlevel 1 (
    echo.
    echo [LỖI] Tool kết thúc với lỗi. Xem chi tiết bên trên.
    pause
)
