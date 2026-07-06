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

def process_file(filename):
    src_path = os.path.join(folder, filename)
    temp_path = os.path.join(scratch_dir, f"temp_{filename}")
    
    t_start = time.time()
    try:
        # Copy file using ctypes to bypass Excel lock
        res = ctypes.windll.kernel32.CopyFileW(src_path, temp_path, False)
        if not res:
            err = ctypes.windll.kernel32.GetLastError()
            return {"filename": filename, "status": "error", "message": f"Copy failed (Win32 Error {err})"}
            
        with zipfile.ZipFile(temp_path, 'r') as z:
            # 1. Read workbook.xml to get sheet names and rIds
            wb_data = z.read('xl/workbook.xml')
            root_wb = ET.fromstring(wb_data)
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                  'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            
            sheets = []
            for sheet in root_wb.findall('.//ns:sheet', ns):
                name = sheet.attrib.get('name')
                rid = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                sheets.append((name, rid))
                
            # Find candidate data sheet
            # Exclude hiddensheet and RC
            candidates = [s for s in sheets if s[0] not in ['_com.sap.ip.bi.xl.hiddensheet', 'RC']]
            if not candidates:
                return {"filename": filename, "status": "error", "message": "No candidate sheets found"}
            
            # Select the first candidate (usually 'Chi tiết data' or 'Sheet1')
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
                return {"filename": filename, "status": "error", "message": f"Sheet relation path not found for {sheet_name}"}
                
            # 3. Load shared strings if it exists
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                sst_file = z.open('xl/sharedStrings.xml')
                context_sst = ET.iterparse(sst_file, events=('start', 'end'))
                for event, elem in context_sst:
                    if event == 'end' and elem.tag.endswith('}si'):
                        text = "".join([t.text for t in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text is not None])
                        shared_strings.append(text)
                        elem.clear()
            
            # 4. Stream parse worksheet to count rows
            sheet_file = z.open(sheet_path_in_zip)
            context = ET.iterparse(sheet_file, events=('start', 'end'))
            
            total_data_rows = 0
            non_empty_af = 0
            header_af = None
            
            # We want to keep track of the maximum row number containing any cell in the sheet
            # to know the total rows count.
            max_row = 0
            
            for event, elem in context:
                if event == 'end' and elem.tag.endswith('}c'):
                    r = elem.attrib.get('r')
                    if r:
                        # Extract row number
                        # Find the first digit index
                        digit_idx = 0
                        for char in r:
                            if char.isdigit():
                                break
                            digit_idx += 1
                        
                        col_part = r[:digit_idx]
                        row_part = r[digit_idx:]
                        
                        if row_part.isdigit():
                            row_num = int(row_part)
                            if row_num > max_row:
                                max_row = row_num
                            
                            if col_part == 'AF':
                                if row_num == 2:
                                    # Header row (Excel row 2 is index 2, 1-based)
                                    t_type = elem.attrib.get('t')
                                    v_elem = elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                                    if v_elem is not None and v_elem.text is not None:
                                        val_text = v_elem.text
                                        if t_type == 's':
                                            idx = int(val_text)
                                            if idx < len(shared_strings):
                                                header_af = shared_strings[idx]
                                        else:
                                            header_af = val_text
                                elif row_num >= 3:
                                    # Data rows
                                    total_data_rows += 1
                                    
                                    t_type = elem.attrib.get('t')
                                    v_elem = elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                                    val = None
                                    if v_elem is not None and v_elem.text is not None:
                                        val_text = v_elem.text
                                        if t_type == 's':
                                            idx = int(val_text)
                                            if idx < len(shared_strings):
                                                val = shared_strings[idx]
                                        else:
                                            val = val_text
                                            
                                    if val is not None and str(val).strip() != "":
                                        non_empty_af += 1
                    elem.clear()
            
            t_end = time.time()
            return {
                "filename": filename,
                "status": "success",
                "sheet_name": sheet_name,
                "header_af": header_af,
                "total_rows": max_row, # max row index in sheet
                "total_data_rows": max_row - 2 if max_row >= 2 else 0, # data rows are from row 3
                "rows_to_check": non_empty_af,
                "time_taken": t_end - t_start
            }
            
    except Exception as e:
        return {"filename": filename, "status": "error", "message": str(e)}
    finally:
        # Clean up temp file
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

if __name__ == '__main__':
    files = sorted([f for f in os.listdir(folder) if f.endswith(".xlsx")])
    print(f"Starting processing of {len(files)} files...")
    
    # Process in parallel using ProcessPoolExecutor
    # Usually we can use 4-6 workers safely
    workers = min(os.cpu_count() or 4, 6)
    print(f"Using {workers} workers.")
    
    results = []
    with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
        # Map the files to the process function
        future_to_file = {executor.submit(process_file, f): f for f in files}
        for future in concurrent.futures.as_completed(future_to_file):
            f = future_to_file[future]
            try:
                res = future.result()
                results.append(res)
                if res["status"] == "success":
                    print(f"Processed {f} successfully in {res['time_taken']:.2f}s: sheet='{res['sheet_name']}', rows_to_check={res['rows_to_check']}/{res['total_data_rows']}")
                else:
                    print(f"Error processing {f}: {res['message']}")
            except Exception as exc:
                print(f"{f} generated an exception: {exc}")
                
    # Sort results by filename
    results = sorted(results, key=lambda x: x["filename"])
    
    print("\n--- SUMMARY RESULTS ---")
    import json
    print(json.dumps(results, indent=2, ensure_ascii=False))
