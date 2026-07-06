"""
Phân tích kỹ hơn cấu trúc dữ liệu để tìm cơ hội tối ưu hóa
"""
import json, os, sys
sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
json_path = os.path.join(scratch_dir, "aggregated_win_dashboard_data.json")

with open(json_path, 'r', encoding='utf-8') as f:
    raw = f.read()

db = json.loads(raw)
rows = db['rows']

print(f"=== PHÂN TÍCH KÍCH THƯỚC FILE ({len(raw)/1024/1024:.2f} MB) ===\n")

# 1. Size breakdown
rows_json = json.dumps(rows, ensure_ascii=False)
meta_json = json.dumps({k: v for k, v in db.items() if k != 'rows'}, ensure_ascii=False)
print(f"Phần metadata (lists): {len(meta_json.encode('utf-8'))/1024:.1f} KB")
print(f"Phần rows (dữ liệu):   {len(rows_json.encode('utf-8'))/1024/1024:.2f} MB  ← ĐÂY LÀ VẤN ĐỀ")
print()

# 2. Per-row breakdown
import random
sample = random.sample(rows, 1000)
r0 = rows[0]
print(f"Sample row: {r0}")
print()

# Field sizes (UTF-8 encoded)
code_bytes = sum(len(str(r[7]).encode('utf-8')) for r in sample) / len(sample)
name_bytes = sum(len(str(r[8]).encode('utf-8')) for r in sample) / len(sample)
nums_bytes = sum(len(f"{r[0]},{r[1]},{r[2]},{r[3]},{r[4]},{r[5]},{r[6]},{r[9]},{r[10]},{r[11]}") for r in sample) / len(sample)
json_overhead = 15  # brackets, commas, quotes

print(f"Phân tích mỗi row (trung bình):")
print(f"  - article_code (r[7]):  {code_bytes:.1f} bytes")
print(f"  - article_name (r[8]):  {name_bytes:.1f} bytes")
print(f"  - Các indices+số (r[0-6,9-11]):  {nums_bytes:.1f} bytes")
print(f"  - JSON overhead:        {json_overhead} bytes")
print(f"  - TỔNG ước tính/row:    {code_bytes+name_bytes+nums_bytes+json_overhead:.1f} bytes")
print()

# 3. Check if qty and value are integers
has_float_qty = sum(1 for r in sample if r[10] != int(r[10]))
has_float_val = sum(1 for r in sample if r[11] != int(r[11]))
print(f"Qty có số thập phân: {has_float_qty}/{len(sample)}")
print(f"Value có số thập phân: {has_float_val}/{len(sample)}")
print()

# 4. Optimization potential
n = len(rows)
print(f"=== ƯỚC TÍNH TIẾT KIỆM KHI TỐI ƯU ===\n")

# Build article lookup
articles = {}
for r in rows:
    key = (str(r[7]), str(r[8]))
    if key not in articles:
        articles[key] = len(articles)

article_list = [None] * len(articles)
for (code, name), idx in articles.items():
    article_list[idx] = [code, name]

article_list_size = len(json.dumps(article_list, ensure_ascii=False).encode('utf-8'))
print(f"Article lookup list ({len(articles):,} entries): {article_list_size/1024:.1f} KB")

# Savings per row: replace code+name (avg ~{code_bytes+name_bytes:.0f} bytes) with index (avg 2-4 bytes)
avg_article_idx_len = len(str(len(articles)))  # max digits
saving_per_row = code_bytes + name_bytes + 4  # +4 for quotes around strings
index_cost = avg_article_idx_len
net_saving = saving_per_row - index_cost
print(f"Tiết kiệm per row: {code_bytes:.0f}+{name_bytes:.0f} bytes → {avg_article_idx_len} bytes = -{net_saving:.0f} bytes/row")
print(f"Tổng tiết kiệm article lookup: {n * net_saving / 1024/1024:.2f} MB")
print()

# Savings from int conversion (remove .0)
qty_saving = sum(len(f"{r[10]}") - len(str(int(r[10]))) for r in sample) / len(sample)
val_saving = sum(len(f"{r[11]}") - len(str(int(r[11]))) for r in sample) / len(sample)
print(f"Tiết kiệm int conversion qty: {qty_saving:.2f} bytes/row → {n*qty_saving/1024/1024:.2f} MB")
print(f"Tiết kiệm int conversion val: {val_saving:.2f} bytes/row → {n*val_saving/1024/1024:.2f} MB")
print()

# Total
total_saving = (n * net_saving + n * (qty_saving + val_saving) - article_list_size) / 1024 / 1024
original_mb = len(raw.encode('utf-8')) / 1024 / 1024
final_mb = original_mb - total_saving
print(f"=== KẾT QUẢ DỰ ÁN ===")
print(f"File hiện tại:  {original_mb:.2f} MB")
print(f"Tổng tiết kiệm: -{total_saving:.2f} MB")
print(f"File sau tối ưu: ~{final_mb:.2f} MB  ({(1-final_mb/original_mb)*100:.0f}% nhỏ hơn)")
