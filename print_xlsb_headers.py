import os
import sys
from pyxlsb import open_workbook

sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\raw mẫu\updated_data_moi.xlsb"
try:
    with open_workbook(path) as wb:
        with wb.get_sheet('master article') as sheet:
            for i, row in enumerate(sheet.rows()):
                if i == 0:
                    row_vals = [cell.v for cell in row]
                    for idx, val in enumerate(row_vals):
                        print(f"Col {idx}: {val}")
                    break
except Exception as e:
    print(f"Error: {e}")
