import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
import ctypes
import concurrent.futures

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"

def extract_stores_from_file(filename):
    src_path = os.path.join(folder, filename)
    temp_path = os.path.join(scratch_dir, f"temp_stores_{filename}")
    
    try:
        res = ctypes.windll.kernel32.CopyFileW(src_path, temp_path, False)
        if not res:
            return {"filename": filename, "status": "error", "message": "Copy failed"}
            
        stores = {} # store_code -> store_name
        with zipfile.ZipFile(temp_path, 'r') as z:
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
                return {"filename": filename, "status": "error", "message": "No candidate sheets"}
            
            sheet_name, target_rid = candidates[0]
            
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
                return {"filename": filename, "status": "error", "message": "Sheet path not found"}
                
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                sst_file = z.open('xl/sharedStrings.xml')
                context_sst = ET.iterparse(sst_file, events=('start', 'end'))
                for event, elem in context_sst:
                    if event == 'end' and elem.tag.endswith('}si'):
                        text = "".join([t.text for t in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text is not None])
                        shared_strings.append(text)
                        elem.clear()
            
            sheet_file = z.open(sheet_path_in_zip)
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
                    if current_row_idx and current_row_idx >= 3:
                        # Extract store code (Col D) and store name (Col E)
                        # We don't filter by AF here, because the user asked "how many stores in total are there in these data files?"
                        # which usually means ALL stores listed in the files, or at least the ones present.
                        code = current_row.get('D')
                        name = current_row.get('E')
                        if code:
                            code_str = str(code).strip()
                            name_str = str(name).strip() if name else ""
                            if code_str != "" and code_str not in stores:
                                stores[code_str] = name_str
                    elem.clear()
            
            return {"filename": filename, "status": "success", "stores": stores}
    except Exception as e:
        return {"filename": filename, "status": "error", "message": str(e)}
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def main():
    t0 = time.time()
    files = sorted([f for f in os.listdir(folder) if f.endswith(".xlsx") and f != "Tong_hop_du_lieu_kiem_tra.xlsx"])
    print(f"Counting unique stores in {len(files)} files...")
    
    workers = min(os.cpu_count() or 4, 6)
    all_stores = {}
    
    with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(extract_stores_from_file, f): f for f in files}
        for future in concurrent.futures.as_completed(futures):
            f = futures[future]
            res = future.result()
            if res["status"] == "success":
                file_stores = res["stores"]
                print(f"  {f} contains {len(file_stores)} stores.")
                for code, name in file_stores.items():
                    if code not in all_stores:
                        all_stores[code] = name
            else:
                print(f"  Error processing {f}: {res['message']}")
                
    print(f"\nTotal unique stores across all files: {len(all_stores)}")
    print(f"Time taken: {time.time()-t0:.2f}s")
    
    # Save the store list to a text file for verification if needed
    out_path = os.path.join(scratch_dir, "unique_stores.txt")
    with open(out_path, "w", encoding="utf-8") as f_out:
        for code in sorted(all_stores.keys()):
            f_out.write(f"{code}: {all_stores[code]}\n")
            
if __name__ == '__main__':
    main()
