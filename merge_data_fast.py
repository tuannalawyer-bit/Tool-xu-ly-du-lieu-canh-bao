import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
import ctypes
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
output_filename = "Tong_hop_du_lieu_kiem_tra.xlsx"
output_path = os.path.join(folder, output_filename)

def clean_sheet_title(title):
    for char in [':', '\\', '/', '?', '*', '[', ']']:
        title = title.replace(char, '')
    return title[:31]

def process_file_to_sheet(filename, wb_out):
    src_path = os.path.join(folder, filename)
    temp_path = os.path.join(scratch_dir, f"temp_m_{filename}")
    
    rsm_name = os.path.splitext(filename)[0]
    sheet_title = clean_sheet_title(rsm_name)
    
    t_start = time.time()
    
    # Copy to temp to bypass Excel lock
    res = ctypes.windll.kernel32.CopyFileW(src_path, temp_path, False)
    if not res:
        print(f"  Error: Copy failed for {filename}")
        return False
        
    try:
        with zipfile.ZipFile(temp_path, 'r') as z:
            # 1. Read workbook.xml to get sheet names
            wb_data = z.read('xl/workbook.xml')
            root_wb = ET.fromstring(wb_data)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            sheets = []
            for sheet in root_wb.findall('.//ns:sheet', ns):
                name = sheet.attrib.get('name')
                rid = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                sheets.append((name, rid))
                
            candidates = [s for s in sheets if s[0] not in ['_com.sap.ip.bi.xl.hiddensheet', 'RC']]
            if not candidates:
                print(f"  Error: No candidate sheets in {filename}")
                return False
                
            sheet_name, target_rid = candidates[0]
            
            # 2. Find sheet file path from rels
            rels_data = z.read('xl/_rels/workbook.xml.rels')
            root_rels = ET.fromstring(rels_data)
            r_ns = {'rels': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            sheet_path_in_zip = None
            for rel in root_rels.findall('.//rels:Relationship', r_ns):
                rid = rel.attrib.get('Id')
                if rid == target_rid:
                    target = rel.attrib.get('Target')
                    sheet_path_in_zip = f"xl/{target}"
                    break
                    
            if not sheet_path_in_zip:
                print(f"  Error: Sheet path not found for {sheet_name}")
                return False
                
            # 3. Load shared strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                sst_file = z.open('xl/sharedStrings.xml')
                context_sst = ET.iterparse(sst_file, events=('start', 'end'))
                for event, elem in context_sst:
                    if event == 'end' and elem.tag.endswith('}si'):
                        text = "".join([t.text for t in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text is not None])
                        shared_strings.append(text)
                        elem.clear()
                        
            # 4. Stream parse worksheet XML
            sheet_file = z.open(sheet_path_in_zip)
            context = ET.iterparse(sheet_file, events=('start', 'end'))
            
            ws_out = wb_out.create_sheet(title=sheet_title)
            
            # Helper to convert col letter to index (A->0, B->1, etc.)
            def col_to_idx(col):
                val = 0
                for char in col:
                    val = val * 26 + (ord(char.upper()) - 64)
                return val - 1
                
            # We want to reconstruct rows from cell elements
            current_row = {}
            current_row_idx = None
            total_copied = 0
            
            def get_val(cell_elem):
                t_type = cell_elem.attrib.get('t')
                v_elem = cell_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                if v_elem is not None and v_elem.text is not None:
                    val_text = v_elem.text
                    if t_type == 's':
                        idx = int(val_text)
                        if idx < len(shared_strings):
                            return shared_strings[idx]
                    # Check if it's float or int
                    if val_text.isdigit():
                        return int(val_text)
                    try:
                        return float(val_text)
                    except:
                        return val_text
                return None
                
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
                        if current_row_idx <= 2:
                            # Headers (Row 1 and 2 in Excel)
                            # Reconstruct row as list
                            row_list = [None] * 32
                            for col_letter, val in current_row.items():
                                idx = col_to_idx(col_letter)
                                if idx < 32:
                                    row_list[idx] = val
                            ws_out.append(row_list)
                        else:
                            # Data rows
                            af_val = current_row.get('AF')
                            if af_val is not None and str(af_val).strip() != "":
                                # Reconstruct row list
                                row_list = [None] * 32
                                for col_letter, val in current_row.items():
                                    idx = col_to_idx(col_letter)
                                    if idx < 32:
                                        row_list[idx] = val
                                ws_out.append(row_list)
                                total_copied += 1
                                
                    elem.clear()
                    
            print(f"  Processed {filename} in {time.time()-t_start:.2f}s: sheet='{sheet_title}', rows_copied={total_copied}")
            return True
            
    except Exception as e:
        print(f"  Error processing {filename}: {e}")
        return False
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def main():
    t_start = time.time()
    
    files = sorted([f for f in os.listdir(folder) if f.endswith(".xlsx") and f != output_filename])
    print(f"Starting merging of {len(files)} files into {output_filename}...")
    
    wb_out = openpyxl.Workbook(write_only=True)
    
    for idx, f in enumerate(files):
        print(f"[{idx+1}/{len(files)}] Processing {f}...")
        process_file_to_sheet(f, wb_out)
        
    print("Saving merged workbook...")
    t_save = time.time()
    wb_out.save(output_path)
    print(f"Successfully saved to: {output_path}")
    print(f"Total time taken: {time.time()-t_start:.2f}s (Save time: {time.time()-t_save:.2f}s)")

if __name__ == '__main__':
    main()
