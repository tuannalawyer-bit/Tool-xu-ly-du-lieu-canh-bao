import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
win_file = os.path.join(folder, "Tong_hop_Kiem_tra_Win.xlsx")

def col_to_idx(col):
    val = 0
    for char in col:
        val = val * 26 + (ord(char.upper()) - 64)
    return val - 1

def main():
    print("Fast header inspection using zipfile (v2)...")
    with zipfile.ZipFile(win_file, 'r') as z:
        sheet_file = z.open('xl/worksheets/sheet1.xml')
        context = ET.iterparse(sheet_file, events=('start', 'end'))
        
        current_row = {}
        current_row_idx = None
        
        def get_val(cell_elem):
            t_type = cell_elem.attrib.get('t')
            if t_type == 'inlineStr':
                t_elem = cell_elem.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if t_elem is not None:
                    return t_elem.text
            v_elem = cell_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if v_elem is not None and v_elem.text is not None:
                return v_elem.text
            return None

        # Read first 3 rows
        row_headers = {}
        for event, elem in context:
            if event == 'start' and elem.tag.endswith('}row'):
                current_row = {}
                r_attr = elem.attrib.get('r')
                current_row_idx = int(r_attr) if r_attr and r_attr.isdigit() else None
                
            elif event == 'end' and elem.tag.endswith('}c'):
                r_ref = elem.attrib.get('r')
                if r_ref:
                    digit_idx = 0
                    for char in r_ref:
                        if char.isdigit():
                            break
                        digit_idx += 1
                    col_letter = r_ref[:digit_idx]
                    current_row[col_letter] = get_val(elem)
                    
            elif event == 'end' and elem.tag.endswith('}row'):
                if current_row_idx:
                    row_headers[current_row_idx] = current_row
                    if current_row_idx >= 3:
                        break
                elem.clear()
                
        # Print headers
        print("\nRow 1 headers:")
        r1 = row_headers.get(1, {})
        for col in sorted(r1.keys(), key=col_to_idx):
            if r1[col] is not None:
                print(f"  {col}: {r1[col]}")
            
        print("\nRow 2 headers:")
        r2 = row_headers.get(2, {})
        for col in sorted(r2.keys(), key=col_to_idx):
            if r2[col] is not None:
                print(f"  {col}: {r2[col]}")
                
        print("\nRow 3 sample:")
        r3 = row_headers.get(3, {})
        for col in sorted(r3.keys(), key=col_to_idx):
            if r3[col] is not None:
                print(f"  {col}: {r3[col]}")

if __name__ == '__main__':
    main()
