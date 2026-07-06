import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\Bùi Anh Tuấn.xlsx"
try:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = wb['Chi tiết data']
    
    # We want to search for row data to see if any cell contains "Lạp xưởng" or "Giò chả" or "Thịt muối"
    found = False
    for i, row in enumerate(sheet.iter_rows(values_only=True)):
        if i < 2:
            continue
        row_str = str(row)
        if "Lạp xưởng" in row_str or "Giò chả" in row_str or "Thịt muối" in row_str:
            print(f"Row {i} contains matching text!")
            for idx, val in enumerate(row):
                print(f"  Col {idx}: {val}")
            found = True
            break
        if i > 5000:
            break
    if not found:
        print("No matching text found in first 5000 rows.")
except Exception as e:
    print(f"Error: {e}")
