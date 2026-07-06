import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]

for f in files:
    path = os.path.join(folder, f)
    try:
        with zipfile.ZipFile(path, 'r') as z:
            wb_data = z.read('xl/workbook.xml')
            root_wb = ET.fromstring(wb_data)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = [sheet.attrib.get('name') for sheet in root_wb.findall('.//ns:sheet', ns)]
            print(f"{f}: {sheets}")
    except Exception as e:
        print(f"Error reading {f}: {e}")
