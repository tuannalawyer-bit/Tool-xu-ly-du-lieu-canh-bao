import ctypes
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')

src = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\Vương Phi Sơn.xlsx"
dst = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\temp_vuong_phi_son.xlsx"

try:
    res = ctypes.windll.kernel32.CopyFileW(src, dst, False)
    if res:
        with zipfile.ZipFile(dst, 'r') as z:
            wb_data = z.read('xl/workbook.xml')
            root_wb = ET.fromstring(wb_data)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = [sheet.attrib.get('name') for sheet in root_wb.findall('.//ns:sheet', ns)]
            print(f"Vương Phi Sơn.xlsx sheets: {sheets}")
            
            # Find data sheet
            candidates = [s for s in sheets if s not in ['_com.sap.ip.bi.xl.hiddensheet', 'RC']]
            print(f"Candidates: {candidates}")
        os.remove(dst)
except Exception as e:
    print(f"Error: {e}")
