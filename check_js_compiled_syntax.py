import re
import sys
import json
import subprocess

html_path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\dashboard_win.html"

with open(html_path, "r", encoding="utf-8") as f:
    content = f.read()

# Extract script blocks
scripts = re.findall(r"<script>(.*?)</script>", content, re.DOTALL)
print(f"Found {len(scripts)} script blocks.")

for i, script in enumerate(scripts):
    # Skip Tailwind script tag
    if "tailwind.config" in script:
        continue
    
    # Write script content to a temporary js file
    temp_js_path = f"C:\\Users\\Tuan\\.gemini\\antigravity\\brain\\b053fc36-8fe6-4316-bba8-ac8fef252caf\\scratch\\temp_check_{i}.js"
    with open(temp_js_path, "w", encoding="utf-8") as f_js:
        f_js.write(script)
        
    print(f"Checking script block {i+1} syntax using Node.js...")
    # Run a node command to evaluate it
    # We wrap the script in a try-catch syntax evaluator
    node_eval_code = f"""
    const fs = require('fs');
    const code = fs.readFileSync({json.dumps(temp_js_path)}, 'utf8');
    try {{
        new Function(code);
        console.log("Syntax is OK!");
    }} catch (e) {{
        console.error("Syntax Error found in block {i+1}:", e.message);
        process.exit(1);
    }}
    """
    
    # Write evaluator
    evaluator_path = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\evaluator.js"
    with open(evaluator_path, "w", encoding="utf-8") as f_eval:
        f_eval.write(f"const code = {json.dumps(script)};\ntry {{ new Function(code); console.log('OK'); }} catch(e) {{ console.error(e.stack); process.exit(1); }}")
        
    res = subprocess.run(["node", evaluator_path], capture_output=True, text=True)
    if res.returncode != 0:
        print(f"FAILED: {res.stderr}")
    else:
        print(f"SUCCESS: {res.stdout.strip()}")
