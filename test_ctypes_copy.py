import ctypes
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

src = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\Vương Phi Sơn.xlsx"
dst = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\temp_vuong_phi_son.xlsx"

try:
    # CopyFileW(lpExistingFileName, lpNewFileName, bFailIfExists)
    res = ctypes.windll.kernel32.CopyFileW(src, dst, False)
    if res:
        print("ctypes CopyFileW success!")
        import zipfile
        with zipfile.ZipFile(dst, 'r') as z:
            print("Zip read successful! Files in zip:", len(z.namelist()))
    else:
        err = ctypes.windll.kernel32.GetLastError()
        print(f"ctypes CopyFileW failed. Error code: {err}")
except Exception as e:
    print(f"Error: {e}")
