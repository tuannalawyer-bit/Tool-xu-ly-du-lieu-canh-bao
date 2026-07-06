import os
import sys
import json
import time

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
template_path = os.path.join(scratch_dir, "dashboard_win_template.html")

def generate_dashboard(chain_name_lower, chain_title):
    t0 = time.time()
    json_filename = f"aggregated_{chain_name_lower}_dashboard_data_v2.json"
    html_filename = f"dashboard_{chain_name_lower}.html"
    
    json_path = os.path.join(scratch_dir, json_filename)
    output_html_path = os.path.join(folder, html_filename)
    
    print(f"\n--- Generating dashboard for {chain_title} -> {html_filename} ---")
    
    if not os.path.exists(json_path):
        print(f"Error: {json_path} not found.")
        return
        
    if not os.path.exists(template_path):
        print(f"Error: {template_path} not found.")
        return
        
    # Read JSON data
    with open(json_path, 'r', encoding='utf-8') as f:
        db_data = json.load(f)
        
    # Read HTML template
    with open(template_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
        
    # Replace WIN references in the template
    # 1. Page title
    html_content = html_content.replace(
        "<title>Báo Cáo Đối Soát Dữ Liệu Kiểm Tra - Chuỗi WIN</title>",
        f"<title>Báo Cáo Đối Soát Dữ Liệu Kiểm Tra - Chuỗi {chain_title}</title>"
    )
    # 2. Header badge
    html_content = html_content.replace(
        '<span class="px-2.5 py-0.5 bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 text-xs font-bold rounded-full">Chuỗi WIN</span>',
        f'<span class="px-2.5 py-0.5 bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 text-xs font-bold rounded-full">Chuỗi {chain_title}</span>'
    )
    # 3. CSV download name
    html_content = html_content.replace(
        '"Bao_cao_chi_tiet_kiem_tra_kho_Win.csv"',
        f'"Bao_cao_chi_tiet_kiem_tra_kho_{chain_title}.csv"'
    )
    
    # Generate the JS injection (v2 format)
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
        print("  Injected rawData (v2 format) into HTML template successfully.")
    else:
        print("  Error: // DATA_PLACEHOLDER not found in template.")
        return
        
    # Save the output HTML
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"  Successfully generated {chain_title} dashboard at: {output_html_path}")
    print(f"  Size: {os.path.getsize(output_html_path) / (1024*1024):.2f} MB")
    print(f"  Total processing time: {time.time()-t0:.2f}s")

def main():
    generate_dashboard("rural", "Rural")
    generate_dashboard("urban", "Urban")

if __name__ == '__main__':
    main()
