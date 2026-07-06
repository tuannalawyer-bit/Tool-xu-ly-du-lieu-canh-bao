import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

src = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\Vương Phi Sơn.xlsx"
dst = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\temp_vuong_phi_son.xlsx"

try:
    import win32file
    import win32con
    
    # Open the file with full share permissions
    handle = win32file.CreateFile(
        src,
        win32con.GENERIC_READ,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None
    )
    
    print("CreateFile success! Reading file...")
    # Read the file contents and write to dst
    with open(dst, 'wb') as f_out:
        chunk_size = 1024 * 1024
        while True:
            err, data = win32file.ReadFile(handle, chunk_size)
            if not data:
                break
            f_out.write(data)
            
    win32file.CloseHandle(handle)
    print("Copy completed successfully via Win32 CreateFile!")
    
    import zipfile
    with zipfile.ZipFile(dst, 'r') as z:
        print("Zip read successful! Files in zip:", len(z.namelist()))
        
except Exception as e:
    print(f"Error: {e}")
