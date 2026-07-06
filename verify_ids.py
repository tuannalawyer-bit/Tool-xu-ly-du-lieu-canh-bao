import re

with open(r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\dashboard_win_template.html", "r", encoding="utf-8") as f:
    content = f.read()

# Find all occurrences of document.getElementById('...') or document.getElementById("...")
ids_in_js = re.findall(r"document\.getElementById\(['\"](.*?)['\"]\)", content)
unique_js_ids = set(ids_in_js)
print("IDs found in JS:", unique_js_ids)

# Check if each ID exists in the HTML as id="..."
missing_ids = []
for id_name in unique_js_ids:
    pattern = f'id="{id_name}"'
    pattern_single = f"id='{id_name}'"
    if pattern not in content and pattern_single not in content:
        missing_ids.append(id_name)

if missing_ids:
    print("ERROR: The following IDs are missing in the HTML:", missing_ids)
else:
    print("SUCCESS: All IDs used in JS exist in the HTML!")
