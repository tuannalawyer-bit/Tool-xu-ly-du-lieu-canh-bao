"""
Tối ưu hóa JSON dữ liệu dashboard:
1. Tạo article_list lookup table (8,641 entries thay vì lặp 455K lần)
2. Chuyển qty/value sang int (bỏ .0)
Kết quả: giảm từ ~43 MB → ~22 MB (tiết kiệm 48%)
"""
import json, os, sys, time
sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch"
json_path     = os.path.join(scratch_dir, "aggregated_win_dashboard_data.json")
json_out_path = os.path.join(scratch_dir, "aggregated_win_dashboard_data_v2.json")

t0 = time.time()
print("Đọc file JSON gốc...")
with open(json_path, 'r', encoding='utf-8') as f:
    db = json.load(f)

rows = db['rows']
print(f"  Đọc xong: {len(rows):,} rows")

# === Bước 1: Xây dựng article lookup table ===
print("Xây dựng article_list lookup...")
article_map = {}   # (code, name) -> index
for r in rows:
    key = (str(r[7]), str(r[8]))
    if key not in article_map:
        article_map[key] = len(article_map)

# article_list[idx] = [code, name]
article_list = [None] * len(article_map)
for (code, name), idx in article_map.items():
    article_list[idx] = [code, name]

print(f"  {len(article_list):,} article entries")

# === Bước 2: Biến đổi rows sang format mới ===
# Old: [store_idx(0), rsm_idx(1), asm_idx(2), mch2_idx(3), mch3_idx(4), mch4_idx(5), mch5_idx(6), code(7), name(8), class_idx(9), qty(10), value(11)]
# New: [store_idx(0), rsm_idx(1), asm_idx(2), mch2_idx(3), mch3_idx(4), mch4_idx(5), mch5_idx(6), article_idx(7), class_idx(8), qty_int(9), value_int(10)]
print("Biến đổi rows sang format nén...")
new_rows = []
for r in rows:
    key = (str(r[7]), str(r[8]))
    art_idx = article_map[key]
    new_row = [
        r[0],            # store_idx
        r[1],            # rsm_idx
        r[2],            # asm_idx
        r[3],            # mch2_idx
        r[4],            # mch3_idx
        r[5],            # mch4_idx
        r[6],            # mch5_idx
        art_idx,         # article_idx (thay thế code+name)
        r[9],            # class_idx
        round(r[10]),    # qty as int
        int(r[11])       # value as int
    ]
    new_rows.append(new_row)

print(f"  Biến đổi xong")

# === Bước 3: Tạo db_v2 ===
db_v2 = {
    'rsm_list':     db['rsm_list'],
    'asm_list':     db['asm_list'],
    'store_list':   db['store_list'],
    'mch2_list':    db['mch2_list'],
    'mch3_list':    db['mch3_list'],
    'mch4_list':    db['mch4_list'],
    'mch5_list':    db['mch5_list'],
    'class_list':   db['class_list'],
    'article_list': article_list,   # <-- MỚI
    'rows':         new_rows
}

# === Bước 4: Ghi file v2 ===
print("Ghi file JSON v2...")
with open(json_out_path, 'w', encoding='utf-8') as f:
    json.dump(db_v2, f, ensure_ascii=False, separators=(',', ':'))  # separators nén thêm whitespace

size_old = os.path.getsize(json_path)
size_new = os.path.getsize(json_out_path)
t_elapsed = time.time() - t0

print()
print(f"=== KẾT QUẢ ===")
print(f"File gốc:   {size_old/1024/1024:.2f} MB")
print(f"File mới:   {size_new/1024/1024:.2f} MB")
print(f"Tiết kiệm:  {(size_old-size_new)/1024/1024:.2f} MB ({(1-size_new/size_old)*100:.1f}%)")
print(f"Thời gian:  {t_elapsed:.1f}s")
print(f"File lưu:   {json_out_path}")
