import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\Bùi Anh Tuấn.xlsx"

with zipfile.ZipFile(path, 'r') as z:
    sheet_file = z.open('xl/worksheets/sheet3.xml')
    context = ET.iterparse(sheet_file, events=('start', 'end'))
    
    count = 0
    for event, elem in context:
        if event == 'end' and elem.tag.endswith('}c'):
            r = elem.attrib.get('r')
            if r in ['AF2', 'AF3', 'AF4', 'AF5', 'AF6']:
                # Print all child elements and attributes
                print(f"Cell {r}: attrib={elem.attrib}")
                for child in elem:
                    print(f"  Child tag={child.tag}, attrib={child.attrib}, text={child.text}")
                count += 1
                if count >= 10:
                    break
            elem.clear()
