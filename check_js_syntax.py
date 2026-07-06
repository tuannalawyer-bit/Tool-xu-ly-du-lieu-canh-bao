import re
import sys

with open(r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\dashboard_win_template.html", "r", encoding="utf-8") as f:
    content = f.read()

# Extract script blocks
scripts = re.findall(r"<script>(.*?)</script>", content, re.DOTALL)
print(f"Found {len(scripts)} script blocks.")

for i, script in enumerate(scripts):
    print(f"Script block {i+1} length: {len(script)}")
    # We can try to check syntax using python or simply print it
    # Especially check if there are any CSS properties or characters like @, etc.
    # or syntax errors
    lines = script.split("\n")
    for j, line in enumerate(lines):
        if "@" in line or "@keyframes" in line:
            print(f"ERROR: Found @ or keyframes in script block {i+1} line {j+1}: {line}")
        # Check for unescaped stuff or placeholder
        if "DATA_PLACEHOLDER" in line:
            print(f"Placeholder found at line {j+1}: {line}")
