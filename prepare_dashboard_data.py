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
output_html_path = os.path.join(folder, "dashboard.html")

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

def process_file_aggregate(filename, product_mch5_map):
    src_path = os.path.join(folder, filename)
    temp_path = os.path.join(scratch_dir, f"temp_agg_{filename}")
    
    t_start = time.time()
    try:
        # Copy file using ctypes to bypass Excel lock
        res = ctypes.windll.kernel32.CopyFileW(src_path, temp_path, False)
        if not res:
            return {"filename": filename, "status": "error", "message": "Copy failed"}
            
        with zipfile.ZipFile(temp_path, 'r') as z:
            # 1. Read workbook.xml to get sheet names and rIds
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
                return {"filename": filename, "status": "error", "message": "Sheet relation path not found"}
                
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
            
            # 4. Stream parse worksheet to count and aggregate rows
            sheet_file = z.open(sheet_path_in_zip)
            context = ET.iterparse(sheet_file, events=('start', 'end'))
            
            aggregated_data = {}
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
                        if af_val is not None and str(af_val).strip() != "":
                            def safe_str(val):
                                return str(val).strip() if val is not None else ""
                            rsm = safe_str(current_row.get('B'))
                            asm = safe_str(current_row.get('C'))
                            mch2 = safe_str(current_row.get('G'))
                            mch3 = safe_str(current_row.get('I'))
                            mch4 = safe_str(current_row.get('J'))
                            art_code = safe_str(current_row.get('K'))
                            art_name = safe_str(current_row.get('L'))
                            
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
                                
                            key = (rsm, asm, mch2, mch3, mch4, mch5, art_code, art_name, af_val)
                            if key not in aggregated_data:
                                aggregated_data[key] = [0.0, 0.0]
                            aggregated_data[key][0] += qty
                            aggregated_data[key][1] += val
                            
                    elem.clear()
                    
            t_end = time.time()
            res_list = []
            for k, v in aggregated_data.items():
                res_list.append({
                    "rsm": k[0],
                    "asm": k[1],
                    "mch2": k[2],
                    "mch3": k[3],
                    "mch4": k[4],
                    "mch5": k[5],
                    "article_code": k[6],
                    "article_name": k[7],
                    "check_class": k[8],
                    "qty": v[0],
                    "value": v[1]
                })
                
            return {
                "filename": filename,
                "status": "success",
                "data": res_list,
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
    t_start = time.time()
    
    product_mch5_map = load_master_articles()
    
    files = sorted([f for f in os.listdir(folder) if f.endswith(".xlsx") and f != "Tong_hop_du_lieu_kiem_tra.xlsx"])
    print(f"Found {len(files)} files to process.")
    
    workers = min(os.cpu_count() or 4, 6)
    print(f"Using {workers} parallel workers.")
    
    all_data = []
    with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process_file_aggregate, f, product_mch5_map): f for f in files}
        for future in concurrent.futures.as_completed(futures):
            f = futures[future]
            try:
                res = future.result()
                if res["status"] == "success":
                    all_data.extend(res["data"])
                    print(f"Processed {f} successfully: aggregated to {len(res['data'])} product rows. Time: {res['time_taken']:.2f}s")
                else:
                    print(f"Error processing {f}: {res['message']}")
            except Exception as e:
                print(f"Exception processing {f}: {e}")
                
    # Now combine the aggregated list (in case different files had the same product,
    # though since each file is a unique RSM, products might repeat across RSMs/ASMs/stores).
    # We will do a final group and merge in Python
    print(f"Combining {len(all_data)} records...")
    combined_data = {}
    for row in all_data:
        key = (
            row["rsm"], row["asm"], row["mch2"], row["mch3"], 
            row["mch4"], row["mch5"], row["article_code"], 
            row["article_name"], row["check_class"]
        )
        if key not in combined_data:
            combined_data[key] = [0.0, 0.0]
        combined_data[key][0] += row["qty"]
        combined_data[key][1] += row["value"]
        
    final_rows = []
    for k, v in combined_data.items():
        final_rows.append({
            "rsm": k[0],
            "asm": k[1],
            "mch2": k[2],
            "mch3": k[3],
            "mch4": k[4],
            "mch5": k[5],
            "article_code": k[6],
            "article_name": k[7],
            "check_class": k[8],
            "qty": v[0],
            "value": v[1]
        })
        
    print(f"Total unique product-level records: {len(final_rows)}")
    
    # Save the JSON data to scratch
    json_path = os.path.join(scratch_dir, "aggregated_dashboard_data.json")
    with open(json_path, "w", encoding="utf-8") as f_out:
        json.dump(final_rows, f_out, ensure_ascii=False, indent=2)
    print(f"Saved aggregated data to {json_path}")
    print(f"Total time taken: {time.time()-t_start:.2f}s")

if __name__ == '__main__':
    main()
