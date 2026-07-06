import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r"C:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\dashboard_win.html", "r", encoding="utf-8") as f:
    content = f.read()

# Let's check if the rawData array is populated. We can print the first 5 entries of rawData by parsing the HTML script tags.
import re
match = re.search(r"const dbData = (\{.*?\});", content)
if match:
    db_data = json.loads(match.group(1))
    print("class_list:", db_data.get("class_list"))
    print("rsm_list:", db_data.get("rsm_list"))
    print("asm_list:", db_data.get("asm_list"))
    print("mch2_list:", db_data.get("mch2_list")[:5])
    print("Number of rows:", len(db_data.get("rows", [])))
    print("Sample rows:", db_data.get("rows", [])[:5])
else:
    print("dbData not found in HTML!")
