import os
import sys
import json
import time

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
# v2: dùng article_list lookup để giảm kích thước file
json_path = os.path.join(scratch_dir, "aggregated_win_dashboard_data_v2.json")
template_path = os.path.join(scratch_dir, "dashboard_win_template.html")
output_html_path = os.path.join(folder, "dashboard_win.html")

def main():
    t0 = time.time()
    print("Generating dashboard_win.html (v2 optimized)...")
    
    if not os.path.exists(json_path):
        print(f"Error: {json_path} not found.")
        return
        
    if not os.path.exists(template_path):
        print(f"Error: {template_path} not found.")
        return
        
    with open(json_path, 'r', encoding='utf-8') as f:
        db_data = json.load(f)
        
    # Read HTML template
    with open(template_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
        
    # Generate the JS injection (v2 format)
    # New row layout: [store_idx(0), rsm_idx(1), asm_idx(2), mch2_idx(3), mch3_idx(4),
    #                  mch4_idx(5), mch5_idx(6), article_idx(7), class_idx(8), qty(9), value(10)]
    js_injection = f"const dbData = {json.dumps(db_data, ensure_ascii=False, separators=(',', ':'))};\n"
    js_injection += "const rawData = dbData.rows.map(r => ({\n"
    js_injection += "    store_code: dbData.store_list[r[0]][0],\n"
    js_injection += "    store_name: dbData.store_list[r[0]][1],\n"
    js_injection += "    rsm: dbData.rsm_list[r[1]],\n"
    js_injection += "    rsm_idx: r[1],\n"
    js_injection += "    asm: dbData.asm_list[r[2]],\n"
    js_injection += "    asm_idx: r[2],\n"
    js_injection += "    mch2: dbData.mch2_list[r[3]],\n"
    js_injection += "    mch2_idx: r[3],\n"
    js_injection += "    mch3: dbData.mch3_list[r[4]],\n"
    js_injection += "    mch3_idx: r[4],\n"
    js_injection += "    mch4: dbData.mch4_list[r[5]],\n"
    js_injection += "    mch4_idx: r[5],\n"
    js_injection += "    mch5: dbData.mch5_list[r[6]],\n"
    js_injection += "    mch5_idx: r[6],\n"
    js_injection += "    article_code: dbData.article_list[r[7]][0],\n"
    js_injection += "    article_name: dbData.article_list[r[7]][1],\n"
    js_injection += "    check_class: dbData.class_list[r[8]],\n"
    js_injection += "    class_idx: r[8],\n"
    js_injection += "    qty: r[9],\n"
    js_injection += "    value: r[10]\n"
    js_injection += "}));"
    
    # Replace placeholder
    if '// DATA_PLACEHOLDER' in html_content:
        html_content = html_content.replace('// DATA_PLACEHOLDER', js_injection)
        print("Injected rawData (v2 format) into HTML template successfully.")
    else:
        print("Error: // DATA_PLACEHOLDER not found in template.")
        return
        
    # Save the output HTML
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    t_end = time.time()
    print(f"Successfully generated Win dashboard at: {output_html_path}")
    print(f"Size: {os.path.getsize(output_html_path) / (1024*1024):.2f} MB")
    print(f"Total processing time: {t_end-t0:.2f}s")

if __name__ == '__main__':
    main()
