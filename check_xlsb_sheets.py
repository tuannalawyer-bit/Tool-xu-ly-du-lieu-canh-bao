import os
import sys
from pyxlsb import open_workbook

sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\raw mẫu\updated_data_moi.xlsb"
try:
    with open_workbook(path) as wb:
        for sheet_name in ['master article', 'MP', 'bizData']:
            if sheet_name in wb.sheets:
                with wb.get_sheet(sheet_name) as sheet:
                    print(f"Sheet: {sheet_name}")
                    for i, row in enumerate(sheet.rows()):
                        if i >= 5:
                            break
                        row_vals = [cell.v for cell in row]
                        print(f"  Row {i}: {row_vals[:15]}")
except Exception as e:
    print(f"Error: {e}")
