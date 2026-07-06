import zipfile
import xml.etree.ElementTree as ET
import os
import sys
import ctypes

# Force unbuffered output
sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
files = [
    "Tong_hop_Kiem_tra_Win.xlsx",
    "Tong_hop_Kiem_tra_Rural.xlsx",
    "Tong_hop_Kiem_tra_Urban.xlsx"
]

def check_file_safe(filename):
    filepath = os.path.join(folder, filename)
    temppath = os.path.join(scratch_dir, f"temp_check_{filename}")
    
    print(f"\nChecking {filename}...", flush=True)
    if not os.path.exists(filepath):
        print("  File not found.", flush=True)
        return
        
    # Copy file to bypass locks
    res = ctypes.windll.kernel32.CopyFileW(filepath, temppath, False)
    if not res:
        print("  Error: Copying file failed.", flush=True)
        return
        
    try:
        with zipfile.ZipFile(temppath, 'r') as z:
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
                        store_code = str(current_row.get('D', '')).strip()
                        store_name = str(current_row.get('E', '')).strip()
                        
                        if "result" in store_code.lower() or "result" in store_name.lower():
                            count_result += 1
                            if count_result <= 3:
                                print(f"  Row {current_row_idx}: D={store_code}, E={store_name}, Q={current_row.get('Q')}", flush=True)
                    elem.clear()
                    
            print(f"  Total rows: {count_total:,}", flush=True)
            print(f"  'Result' rows found: {count_result:,}", flush=True)
    except Exception as e:
        print(f"  Error reading zip: {e}", flush=True)
    finally:
        if os.path.exists(temppath):
            try:
                os.remove(temppath)
            except:
                pass

def main():
    for f in files:
        check_file_safe(f)

if __name__ == '__main__':
    main()
