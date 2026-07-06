import os
import sys
from pyxlsb import open_workbook

sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\raw mẫu\updated_data_moi.xlsb"
try:
    with open_workbook(path) as wb:
        print("Sheets in updated_data_moi.xlsb:", wb.sheets)
        # Let's inspect the first sheet
        with wb.get_sheet(1) as sheet:
            print("First sheet rows:")
            for i, row in enumerate(sheet.rows()):
                if i >= 5:
                    break
                row_vals = [cell.v for cell in row]
                print(f"  Row {i}: {row_vals[:15]}")
except Exception as e:
    print(f"Error: {e}")
