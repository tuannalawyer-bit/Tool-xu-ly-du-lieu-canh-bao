import re
import sys
import subprocess
import json

html_path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\dashboard_win.html"

with open(html_path, "r", encoding="utf-8") as f:
    content = f.read()

# Extract script blocks
scripts = re.findall(r"<script>(.*?)</script>", content, re.DOTALL)
print(f"Found {len(scripts)} script blocks.")

# We want script block 3, which is the main logic.
# But block 2 has the DATA_PLACEHOLDER.
# We will combine them so that dbData is defined.
db_data_script = scripts[1]
main_logic_script = scripts[2]

combined_script = db_data_script + "\n" + main_logic_script

# Write a wrapper js file that mocks DOM and runs it
mock_js = f"""
// Mock DOM
const elementMock = {{
    value: "",
    textContent: "",
    innerHTML: "",
    setAttribute: () => {{}},
    appendChild: () => {{}},
    addEventListener: () => {{}},
    querySelector: () => ({{ classList: {{ remove: () => {{}} }} }}),
    classList: {{ remove: () => {{}} }},
    style: {{}}
}};

global.window = {{
    onload: null
}};
global.document = {{
    getElementById: (id) => {{
        // console.log("Mock getElementById called for ID:", id);
        return elementMock;
    }},
    createElement: (tag) => {{
        return elementMock;
    }}
}};
global.Intl = {{
    NumberFormat: function() {{
        return {{
            format: (val) => String(val)
        }};
    }}
}};

// Run the combined script
try {{
    {combined_script}
    console.log("Compilation and execution succeeded!");
    if (window.onload) {{
        console.log("Calling window.onload()...");
        window.onload();
        console.log("window.onload() completed successfully!");
    }} else {{
        console.log("Warning: window.onload is not set.");
    }}
}} catch (e) {{
    console.error("RUNTIME ERROR:");
    console.error(e.stack);
    process.exit(1);
}}
"""

evaluator_path = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\runtime_evaluator.js"
with open(evaluator_path, "w", encoding="utf-8") as f_eval:
    f_eval.write(mock_js)

print("Running runtime evaluator with Node.js...")
res = subprocess.run(["node", evaluator_path], capture_output=True, text=True)
if res.returncode != 0:
    print(f"RUNTIME FAILED:\nSTDOUT:\n{res.stdout}\nSTDERR:\n{res.stderr}")
else:
    print(f"RUNTIME SUCCESS:\n{res.stdout}")
