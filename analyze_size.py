import json, os, sys
sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
json_path = os.path.join(scratch_dir, "aggregated_win_dashboard_data.json")

with open(json_path, 'r', encoding='utf-8') as f:
    db = json.load(f)

rows = db.get('rows', [])
print(f"Total rows: {len(rows):,}")
print(f"Sample row: {rows[0] if rows else 'N/A'}")
print(f"Row fields: store_idx(0), rsm_idx(1), asm_idx(2), mch2_idx(3), mch3_idx(4), mch4_idx(5), mch5_idx(6), article_code(7), article_name(8), class_idx(9), qty(10), value(11)")
print()
print(f"rsm_list count: {len(db.get('rsm_list', []))}")
print(f"asm_list count: {len(db.get('asm_list', []))}")
print(f"store_list count: {len(db.get('store_list', []))}")
print(f"mch2_list count: {len(db.get('mch2_list', []))}")
print(f"mch3_list count: {len(db.get('mch3_list', []))}")
print(f"mch4_list count: {len(db.get('mch4_list', []))}")
print(f"mch5_list count: {len(db.get('mch5_list', []))}")
print(f"class_list count: {len(db.get('class_list', []))}")
print()

# Check unique values in article data
article_codes = set(r[7] for r in rows)
article_names = set(r[8] for r in rows)
print(f"Unique article_codes: {len(article_codes):,}")
print(f"Unique article_names: {len(article_names):,}")

# Estimate size breakdown
import json as jmod
json_str = jmod.dumps(db, ensure_ascii=False)
print(f"\ndbData JSON size: {len(json_str.encode('utf-8')) / (1024*1024):.2f} MB")

# Estimate rawData expansion size
# Each row becomes an object with ~10 string fields decoded
# article_name avg length (sample)
avg_name_len = sum(len(str(r[8])) for r in rows[:1000]) / 1000
avg_code_len = sum(len(str(r[7])) for r in rows[:1000]) / 1000
print(f"Avg article_name length: {avg_name_len:.1f} chars")
print(f"Avg article_code length: {avg_code_len:.1f} chars")
# rawData overhead per row estimate
print(f"\nEstimated rawData expansion size: {len(rows) * (avg_name_len + avg_code_len + 150) / (1024*1024):.2f} MB")
print("(150 bytes per row for field names, indices, numeric values, JSON syntax)")
