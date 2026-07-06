import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
files = [
    "Tong_hop_Kiem_tra_Win.xlsx",
    "Tong_hop_Kiem_tra_Rural.xlsx",
    "Tong_hop_Kiem_tra_Urban.xlsx"
]

def check_file_fast(filename):
    filepath = os.path.join(folder, filename)
    print(f"\nChecking {filename}...")
    if not os.path.exists(filepath):
        print("  File not found.")
        return
        
    with zipfile.ZipFile(filepath, 'r') as z:
        sheet_file = z.open('xl/worksheets/sheet1.xml')
        context = ET.iterparse(sheet_file, events=('start', 'end'))
        
        current_row = {}
        current_row_idx = None
        
        def get_val(cell_elem):
            # Check inlineStr
            t_type = cell_elem.attrib.get('t')
            if t_type == 'inlineStr':
                t_elem = cell_elem.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if t_elem is not None:
                    return t_elem.text
            v_elem = cell_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if v_elem is not None and v_elem.text is not None:
                return v_elem.text
            return None

        count_total = 0
        count_result = 0
        
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
                if current_row_idx and current_row_idx > 2:
                    count_total += 1
                    # Store Code is in column D.
                    # Since Column A (map) is still there, Store Code is D, Store Name is E.
                    store_code = str(current_row.get('D', '')).strip()
                    store_name = str(current_row.get('E', '')).strip()
                    
                    if "result" in store_code.lower() or "result" in store_name.lower():
                        count_result += 1
                        if count_result <= 3:
                            print(f"  Row {current_row_idx}: D={store_code}, E={store_name}, value={current_row.get('R')}")
                elem.clear()
                
        print(f"  Total rows: {count_total:,}")
        print(f"  'Result' rows found: {count_result:,}")

def main():
    for f in files:
        check_file_fast(f)

if __name__ == '__main__':
    main()
