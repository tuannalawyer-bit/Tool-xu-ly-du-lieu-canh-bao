import os
import sys
import json

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
json_path = os.path.join(scratch_dir, "aggregated_dashboard_data.json")
template_path = os.path.join(scratch_dir, "dashboard_template.html")
output_path = os.path.join(folder, "dashboard.html")

def main():
    print("Generating dashboard.html...")
    if not os.path.exists(json_path):
        print(f"Error: JSON data file {json_path} does not exist yet. Please wait for aggregation to finish.")
        return
        
    if not os.path.exists(template_path):
        print(f"Error: HTML template file {template_path} does not exist.")
        return
        
    # Read JSON data
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    # Convert data to Javascript declaration
    data_js = f"const rawData = {json.dumps(data, ensure_ascii=False)};"
    
    # Read HTML template
    with open(template_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
        
    # Replace placeholder
    # Check if '// DATA_PLACEHOLDER' exists in template
    if '// DATA_PLACEHOLDER' in html_content:
        html_content = html_content.replace('// DATA_PLACEHOLDER', data_js)
        print("Injected rawData into HTML template successfully.")
    else:
        print("Warning: '// DATA_PLACEHOLDER' comment not found in template. Injected at head of script tag.")
        # fallback
        html_content = html_content.replace('<script>', f'<script>\n        {data_js}', 1)
        
    # Write output HTML file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"Dashboard generated successfully at: {output_path}")
    print(f"File size: {os.path.getsize(output_path) / 1024:.2f} KB")

if __name__ == '__main__':
    main()
