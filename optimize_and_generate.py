import os
import sys
import json
import time

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
json_path = os.path.join(scratch_dir, "aggregated_dashboard_data.json")
template_path = os.path.join(scratch_dir, "dashboard_template.html")
output_html_path = os.path.join(folder, "dashboard.html")

def main():
    t0 = time.time()
    print("Optimizing aggregated data for dashboard...")
    
    if not os.path.exists(json_path):
        print(f"Error: {json_path} not found.")
        return
        
    with open(json_path, 'r', encoding='utf-8') as f:
        raw_rows = json.load(f)
        
    print(f"Loaded {len(raw_rows)} records from JSON.")
    
    # Collect unique string values
    rsms = set()
    asms = set()
    mch2s = set()
    mch3s = set()
    mch4s = set()
    mch5s = set()
    classes = set()
    
    for r in raw_rows:
        rsms.add(r["rsm"])
        asms.add(r["asm"])
        mch2s.add(r["mch2"])
        mch3s.add(r["mch3"])
        mch4s.add(r["mch4"])
        mch5s.add(r["mch5"])
        classes.add(r["check_class"])
        
    # Sort lists to keep them stable
    rsm_list = sorted(list(rsms))
    asm_list = sorted(list(asms))
    mch2_list = sorted(list(mch2s))
    mch3_list = sorted(list(mch3s))
    mch4_list = sorted(list(mch4s))
    mch5_list = sorted(list(mch5s))
    class_list = sorted(list(classes))
    
    # Create mapping dictionaries
    rsm_map = {name: idx for idx, name in enumerate(rsm_list)}
    asm_map = {name: idx for idx, name in enumerate(asm_list)}
    mch2_map = {name: idx for idx, name in enumerate(mch2_list)}
    mch3_map = {name: idx for idx, name in enumerate(mch3_list)}
    mch4_map = {name: idx for idx, name in enumerate(mch4_list)}
    mch5_map = {name: idx for idx, name in enumerate(mch5_list)}
    class_map = {name: idx for idx, name in enumerate(class_list)}
    
    # Map rows to compressed lists
    # Format of each row:
    # [ rsm_idx, asm_idx, mch2_idx, mch3_idx, mch4_idx, mch5_idx, article_code, article_name, class_idx, qty, value ]
    compressed_rows = []
    for r in raw_rows:
        row = [
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
        
    dashboard_data = {
        "rsm_list": rsm_list,
        "asm_list": asm_list,
        "mch2_list": mch2_list,
        "mch3_list": mch3_list,
        "mch4_list": mch4_list,
        "mch5_list": mch5_list,
        "class_list": class_list,
        "rows": compressed_rows
    }
    
    print(f"Compressed rows count: {len(compressed_rows)}")
    
    # Load HTML template
    if not os.path.exists(template_path):
        print(f"Error: {template_path} not found.")
        return
        
    with open(template_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
        
    # Generate the JS injection
    js_injection = f"const dbData = {json.dumps(dashboard_data, ensure_ascii=False)};\n"
    js_injection += "const rawData = dbData.rows.map(r => ({\n"
    js_injection += "    rsm: dbData.rsm_list[r[0]],\n"
    js_injection += "    asm: dbData.asm_list[r[1]],\n"
    js_injection += "    mch2: dbData.mch2_list[r[2]],\n"
    js_injection += "    mch3: dbData.mch3_list[r[3]],\n"
    js_injection += "    mch4: dbData.mch4_list[r[4]],\n"
    js_injection += "    mch5: dbData.mch5_list[r[5]],\n"
    js_injection += "    article_code: r[6],\n"
    js_injection += "    article_name: r[7],\n"
    js_injection += "    check_class: dbData.class_list[r[8]],\n"
    js_injection += "    qty: r[9],\n"
    js_injection += "    value: r[10]\n"
    js_injection += "}));"
    
    # Replace placeholder
    if '// DATA_PLACEHOLDER' in html_content:
        html_content = html_content.replace('// DATA_PLACEHOLDER', js_injection)
        print("Injected compressed data into HTML template successfully.")
    else:
        print("Error: // DATA_PLACEHOLDER not found in template.")
        return
        
    # Save the output HTML
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    t_end = time.time()
    print(f"Successfully generated optimized dashboard at: {output_html_path}")
    print(f"Size: {os.path.getsize(output_html_path) / (1024*1024):.2f} MB")
    print(f"Total processing time: {t_end-t0:.2f}s")

if __name__ == '__main__':
    main()
