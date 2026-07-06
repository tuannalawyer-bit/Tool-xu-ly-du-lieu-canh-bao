import os
import sys
import re
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"

def check_file(filename):
    print(f"\nChecking runtime for {filename}...")
    filepath = os.path.join(folder, filename)
    if not os.path.exists(filepath):
        print(f"Error: {filepath} does not exist.")
        return False
        
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    # Find script blocks
    script_blocks = re.findall(r"<script>(.*?)</script>", content, re.DOTALL)
    print(f"  Found {len(script_blocks)} script blocks.")

    if len(script_blocks) < 3:
        print("  Error: Missing expected script blocks (need data + main logic).")
        return False

    # Extract block 2 (data) and block 3 (main logic)
    db_data_script = script_blocks[1]
    main_logic_script = script_blocks[2]

    # Clean the main script of window.onload wrapper to test compilation
    # and mock the DOM elements
    js_runtime_test = """
    // Mock browser DOM
    global.window = { onload: null };
    global.document = {
        getElementById: function(id) {
            return {
                value: "",
                textContent: "",
                innerHTML: "",
                setAttribute: function() {},
                appendChild: function() {},
                addEventListener: function() {},
                querySelector: function() {
                    return { classList: { remove: function() {} } };
                },
                classList: { remove: function() {} },
                style: {}
            };
        },
        createElement: function() {
            return {
                value: "",
                textContent: "",
                innerHTML: "",
                style: {},
                classList: { remove: function() {} },
                querySelector: function() {
                    return { classList: { remove: function() {} } };
                },
                appendChild: function() {}
            };
        },
        body: {
            insertBefore: function() {}
        }
    };
    global.Intl = Intl;

    // Inject data
    """ + db_data_script + """

    // Inject main logic
    """ + main_logic_script + """

    // Run tests
    console.log("Compilation and execution succeeded!");
    if (typeof window.onload === 'function') {
        console.log("Calling window.onload()...");
        window.onload();
        console.log("window.onload() completed successfully!");
    } else {
        console.log("Error: window.onload is not a function.");
        process.exit(1);
    }
    """

    test_js_path = os.path.join(scratch_dir, f"temp_check_{filename.replace('.html', '')}.js")
    with open(test_js_path, "w", encoding="utf-8") as f_out:
        f_out.write(js_runtime_test)

    # Run in Node.js
    res = subprocess.run(["node", test_js_path], capture_output=True, text=True, encoding="utf-8")
    
    # Cleanup
    if os.path.exists(test_js_path):
        os.remove(test_js_path)

    if res.returncode == 0:
        print(f"  SUCCESS: {filename} executes correctly in Node/V8 environment.")
        return True
    else:
        print(f"  FAILURE: {filename} runtime failed.")
        print(res.stderr)
        return False

def main():
    success = True
    for f in ["dashboard_rural.html", "dashboard_urban.html"]:
        if not check_file(f):
            success = False
            
    if not success:
        sys.exit(1)

if __name__ == '__main__':
    main()
