import openpyxl
import os
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]

f = files[0]
path = os.path.join(folder, f)
print(f"File: {f}")
t0 = time.time()
try:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = wb['Chi tiết data']
    total_rows = 0
    non_empty_af = 0
    for i, row in enumerate(sheet.iter_rows(values_only=True)):
        if i < 2:
            continue
        total_rows += 1
        val = row[31] if len(row) > 31 else None
        if val is not None and str(val).strip() != "":
            non_empty_af += 1
    t1 = time.time()
    print(f"Time taken: {t1-t0:.2f} seconds")
    print(f"Total data rows checked: {total_rows}")
    print(f"Non-empty AF: {non_empty_af}")
except Exception as e:
    print(f"Error: {e}")
