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

# Define chains and their corresponding files
chains = {
    "Rural": [
        "Lê Thị Hồng Thu.xlsx",
        "Hoàng Nguyễn Tú Anh.xlsx",
        "Bùi Anh Tuấn.xlsx",
        "Lạc Nhật Minh.xlsx"
    ],
    "Urban": [
        "Phạm Thị Ngọc Xuyến.xlsx",
        "Lê Duy Đức.xlsx",
        "Vương Phi Sơn.xlsx",
        "Trần Thị Diệp.xlsx",
        "Đỗ Khắc Chức.xlsx",
        "NGuyễn Văn Tuấn.xlsx"
    ],
    "Win": [
        "Nguyễn Đức Thiên Ân.xlsx",
        "Đỗ Thị Thanh Loan.xlsx",
        "Lê Văn Trí.xlsx",
        "Nguyễn minh Trang.xlsx"
    ]
}

def normalize_rsm_name(name):
    if not name:
        return ""
    name_str = str(name).strip()
    name_lower = name_str.lower()
    if name_lower == "nguyễn văn tuấn":
        return "Nguyễn Văn Tuấn"
    if name_lower == "nguyễn minh trang":
        return "Nguyễn Minh Trang"
    return name_str

def col_to_idx(col):
    val = 0
    for char in col:
        val = val * 26 + (ord(char.upper()) - 64)
    return val - 1

def process_chain(chain_name, file_list):
    t_start = time.time()
    output_filename = f"Tong_hop_Kiem_tra_{chain_name}.xlsx"
    output_path = os.path.join(folder, output_filename)
    print(f"\n--- Processing Chain: {chain_name} -> {output_filename} ---")
    
    wb_out = openpyxl.Workbook(write_only=True)
    ws_out = wb_out.create_sheet(title="Dữ liệu kiểm tra")
    
    headers_written = False
    
    for idx, filename in enumerate(file_list):
        src_path = os.path.join(folder, filename)
        temp_path = os.path.join(scratch_dir, f"temp_split_{filename}")
        
        print(f"  [{idx+1}/{len(file_list)}] Reading {filename}...")
        
        # Copy to temp to bypass lock
        res = ctypes.windll.kernel32.CopyFileW(src_path, temp_path, False)
        if not res:
            print(f"    Error: Copy failed for {filename}")
            continue
            
        try:
            with zipfile.ZipFile(temp_path, 'r') as z:
                # Get sheet list
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
                    print(f"    Error: No data sheets in {filename}")
                    continue
                    
                sheet_name, target_rid = candidates[0]
                
                # Get sheet XML path
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
                    print(f"    Error: Sheet path not found for {sheet_name}")
                    continue
                    
                # Load shared strings
                shared_strings = []
                if 'xl/sharedStrings.xml' in z.namelist():
                    sst_file = z.open('xl/sharedStrings.xml')
                    context_sst = ET.iterparse(sst_file, events=('start', 'end'))
                    for event, elem in context_sst:
                        if event == 'end' and elem.tag.endswith('}si'):
                            text = "".join([t.text for t in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text is not None])
                            shared_strings.append(text)
                            elem.clear()
                            
                # Stream parse worksheet XML
                sheet_file = z.open(sheet_path_in_zip)
                context = ET.iterparse(sheet_file, events=('start', 'end'))
                
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
                                # Row 1 and 2 (Excel headers)
                                # Only write headers once from the first file in the chain
                                if not headers_written:
                                    row_list = [None] * 32
                                    for col_letter, val in current_row.items():
                                        idx = col_to_idx(col_letter)
                                        if idx < 32:
                                            row_list[idx] = val
                                    ws_out.append(row_list)
                            else:
                                # Data rows
                                af_val = current_row.get('AF')
                                if af_val is not None:
                                    af_str = str(af_val).strip()
                                    if af_str != "" and af_str.upper() not in ["N/A", "#N/A"]:
                                        # Reconstruct row list
                                        row_list = [None] * 32
                                        for col_letter, val in current_row.items():
                                            idx = col_to_idx(col_letter)
                                            if idx < 32:
                                                row_list[idx] = val
                                                
                                        # Normalize RSM name in Column B (index 1)
                                        row_list[1] = normalize_rsm_name(row_list[1])
                                        ws_out.append(row_list)
                                        total_copied += 1
                                        
                        elem.clear()
                
                if not headers_written:
                    headers_written = True
                print(f"    Copied {total_copied} rows from {filename}")
                
        except Exception as e:
            print(f"    Error processing {filename}: {e}")
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass
                    
    print("  Saving workbook...")
    t_save = time.time()
    wb_out.save(output_path)
    print(f"  Saved to: {output_path} (Save time: {time.time()-t_save:.2f}s)")
    print(f"  Total time for chain {chain_name}: {time.time()-t_start:.2f}s")

def main():
    t_total_start = time.time()
    
    # Process all 3 chains
    for chain_name, file_list in chains.items():
        process_chain(chain_name, file_list)
        
    # Clean up the old single merge file if it exists
    old_file_path = os.path.join(folder, "Tong_hop_du_lieu_kiem_tra.xlsx")
    if os.path.exists(old_file_path):
        try:
            os.remove(old_file_path)
            print(f"\nRemoved old merged file: {old_file_path}")
        except Exception as e:
            print(f"\nCould not remove old merged file: {e}")
            
    print(f"\nAll chains processed! Total execution time: {time.time()-t_total_start:.2f}s")

if __name__ == '__main__':
    main()
