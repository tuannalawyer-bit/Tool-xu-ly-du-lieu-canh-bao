import os
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
import ctypes
import concurrent.futures
import json
from pyxlsb import open_workbook

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
xlsb_path = os.path.join(folder, "raw mẫu", "updated_data_moi.xlsb")

win_files = [
    "Nguyễn Đức Thiên Ân.xlsx",
    "Đỗ Thị Thanh Loan.xlsx",
    "Lê Văn Trí.xlsx",
    "Nguyễn minh Trang.xlsx"
]

def load_master_articles():
    print("Loading master articles from XLSB...")
    t0 = time.time()
    mapping = {}
    try:
        with open_workbook(xlsb_path) as wb:
            with wb.get_sheet('master article') as sheet:
                for i, row in enumerate(sheet.rows()):
                    if i == 0:
                        continue
                    prod_id = row[0].v
                    mch5_desc = row[22].v if len(row) > 22 else None
                    if prod_id:
                        prod_id_str = str(int(prod_id)) if isinstance(prod_id, float) else str(prod_id).strip()
                        mapping[prod_id_str] = str(mch5_desc).strip() if mch5_desc else "Khác"
        print(f"Loaded {len(mapping)} articles in {time.time()-t0:.2f}s")
    except Exception as e:
        print(f"Error loading master articles: {e}")
    return mapping

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

def process_win_file_raw(filename, product_mch5_map):
    src_path = os.path.join(folder, filename)
    temp_path = os.path.join(scratch_dir, f"temp_win_raw_{filename}")
    
    t_start = time.time()
    try:
        res = ctypes.windll.kernel32.CopyFileW(src_path, temp_path, False)
        if not res:
            return {"filename": filename, "status": "error", "message": "Copy failed"}
            
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
                return {"filename": filename, "status": "error", "message": "No candidate sheets"}
            
            sheet_name, target_rid = candidates[0]
            
            # 2. Find sheet XML path
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
            
            rows_data = []
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
                        af_val = current_row.get('AF')
                        if af_val is not None:
                            af_str = str(af_val).strip()
                            if af_str != "" and af_str.upper() not in ["N/A", "#N/A"]:
                                rsm = normalize_rsm_name(current_row.get('B'))
                                asm = str(current_row.get('C', '')).strip()
                                store_code = str(current_row.get('D', '')).strip()
                                store_name = str(current_row.get('E', '')).strip()
                                mch2 = str(current_row.get('G', '')).strip()
                                mch3 = str(current_row.get('I', '')).strip()
                                mch4 = str(current_row.get('J', '')).strip()
                                art_code = str(current_row.get('K', '')).strip()
                                art_name = str(current_row.get('L', '')).strip()
                                
                                if art_code.endswith('.0'):
                                    art_code = art_code[:-2]
                                    
                                mch5 = product_mch5_map.get(art_code, "Khác")
                                
                                qty_str = current_row.get('Q', '0')
                                val_str = current_row.get('R', '0')
                                
                                try:
                                    qty = float(qty_str) if qty_str else 0.0
                                except:
                                    qty = 0.0
                                    
                                try:
                                    val = float(val_str) if val_str else 0.0
                                except:
                                    val = 0.0
                                    
                                rows_data.append({
                                    "rsm": rsm,
                                    "asm": asm,
                                    "store_code": store_code,
                                    "store_name": store_name,
                                    "mch2": mch2,
                                    "mch3": mch3,
                                    "mch4": mch4,
                                    "mch5": mch5,
                                    "article_code": art_code,
                                    "article_name": art_name,
                                    "qty": qty,
                                    "value": val,
                                    "check_class": af_str
                                })
                                
                    elem.clear()
                    
            t_end = time.time()
            return {
                "filename": filename,
                "status": "success",
                "data": rows_data,
                "time_taken": t_end - t_start
            }
    except Exception as e:
        return {"filename": filename, "status": "error", "message": str(e)}
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def main():
    t_total_start = time.time()
    
    product_mch5_map = load_master_articles()
    
    print(f"\nProcessing Win chain files in parallel...")
    workers = min(os.cpu_count() or 4, 4)
    
    all_rows = []
    
    with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process_win_file_raw, f, product_mch5_map): f for f in win_files}
        for future in concurrent.futures.as_completed(futures):
            f = futures[future]
            try:
                res = future.result()
                if res["status"] == "success":
                    all_rows.extend(res["data"])
                    print(f"  Processed {f} in {res['time_taken']:.2f}s: rows={len(res['data'])}")
                else:
                    print(f"  Error processing {f}: {res['message']}")
            except Exception as e:
                print(f"  Exception processing {f}: {e}")
                
    print(f"\nTotal rows extracted for Win: {len(all_rows)}")
    
    # Compress data using unique string indexing
    print("Compressing data...")
    t_comp = time.time()
    
    rsms = set()
    asms = set()
    stores = {} # store_code -> store_name
    mch2s = set()
    mch3s = set()
    mch4s = set()
    mch5s = set()
    classes = set()
    
    for r in all_rows:
        rsms.add(r["rsm"])
        asms.add(r["asm"])
        stores[r["store_code"]] = r["store_name"]
        mch2s.add(r["mch2"])
        mch3s.add(r["mch3"])
        mch4s.add(r["mch4"])
        mch5s.add(r["mch5"])
        classes.add(r["check_class"])
        
    rsm_list = sorted(list(rsms))
    asm_list = sorted(list(asms))
    
    # Store list of lists: [[code, name], ...] sorted by code
    store_list = [[code, name] for code, name in sorted(stores.items())]
    store_map = {item[0]: idx for idx, item in enumerate(store_list)}
    
    mch2_list = sorted(list(mch2s))
    mch3_list = sorted(list(mch3s))
    mch4_list = sorted(list(mch4s))
    mch5_list = sorted(list(mch5s))
    class_list = sorted(list(classes))
    
    rsm_map = {name: idx for idx, name in enumerate(rsm_list)}
    asm_map = {name: idx for idx, name in enumerate(asm_list)}
    mch2_map = {name: idx for idx, name in enumerate(mch2_list)}
    mch3_map = {name: idx for idx, name in enumerate(mch3_list)}
    mch4_map = {name: idx for idx, name in enumerate(mch4_list)}
    mch5_map = {name: idx for idx, name in enumerate(mch5_list)}
    class_map = {name: idx for idx, name in enumerate(class_list)}
    
    compressed_rows = []
    for r in all_rows:
        row = [
            store_map[r["store_code"]],
            rsm_map[r["rsm"]],
            asm_map[r["asm"]],
            mch2_map[r["mch2"]],
            mch3_map[r["mch3"]],
            mch4_map[r["mch4"]],
            mch5_map[r["mch5"]],
            r["article_code"],
            r["article_name"],
            class_map[r["check_class"]],
            r["qty"],
            r["value"]
        ]
        compressed_rows.append(row)
        
    out_dict = {
        "rsm_list": rsm_list,
        "asm_list": asm_list,
        "store_list": store_list,
        "mch2_list": mch2_list,
        "mch3_list": mch3_list,
        "mch4_list": mch4_list,
        "mch5_list": mch5_list,
        "class_list": class_list,
        "rows": compressed_rows
    }
    
    out_path = os.path.join(scratch_dir, "aggregated_win_dashboard_data.json")
    with open(out_path, "w", encoding="utf-8") as f_out:
        json.dump(out_dict, f_out, ensure_ascii=False)
        
    print(f"Compressed in {time.time()-t_comp:.2f}s")
    print(f"Saved combined Win dashboard data to: {out_path}")
    print(f"Total processing time: {time.time()-t_total_start:.2f}s")

if __name__ == '__main__':
    main()
