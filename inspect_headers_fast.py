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
    print("Fast header inspection using zipfile...")
    with zipfile.ZipFile(win_file, 'r') as z:
        # Load shared strings
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            print("Loading shared strings...")
            sst_file = z.open('xl/sharedStrings.xml')
            context_sst = ET.iterparse(sst_file, events=('start', 'end'))
            for event, elem in context_sst:
                if event == 'end' and elem.tag.endswith('}si'):
                    text = "".join([t.text for t in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text is not None])
                    shared_strings.append(text)
                    elem.clear()
            print(f"Loaded {len(shared_strings)} shared strings.")
            
        # Read worksheet
        print("Reading sheet...")
        sheet_file = z.open('xl/worksheets/sheet1.xml')
        context = ET.iterparse(sheet_file, events=('start', 'end'))
        
        current_row = {}
        current_row_idx = None
        
        def get_val(cell_elem):
            t_type = cell_elem.attrib.get('t')
            v_elem = cell_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if v_elem is not None and v_elem.text is not None:
                val_text = v_elem.text
                if t_type == 's':
                    idx = int(val_text)
                    if idx < len(shared_strings):
                        return shared_strings[idx]
                return val_text
            return None

        # Let's read first 2 rows (since row 1 and 2 contain headers in these files)
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
                    if current_row_idx >= 2:
                        break
                elem.clear()
                
        # Print headers
        print("\nRow 1 headers:")
        r1 = row_headers.get(1, {})
        for col in sorted(r1.keys(), key=col_to_idx):
            print(f"  {col}: {r1[col]}")
            
        print("\nRow 2 headers:")
        r2 = row_headers.get(2, {})
        for col in sorted(r2.keys(), key=col_to_idx):
            print(f"  {col}: {r2[col]}")

if __name__ == '__main__':
    main()
