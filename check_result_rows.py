import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu"
files = [
    "Tong_hop_Kiem_tra_Win.xlsx",
    "Tong_hop_Kiem_tra_Rural.xlsx",
    "Tong_hop_Kiem_tra_Urban.xlsx"
]

def check_result_rows(filename):
    filepath = os.path.join(folder, filename)
    print(f"\nChecking {filename}...")
    if not os.path.exists(filepath):
        print("  File not found.")
        return
        
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    
    count_result = 0
    count_total = 0
    
    # Read headers
    headers = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if r_idx <= 2:
            headers.append(row)
            continue
        count_total += 1
        # Store code was in Column D in original file (index 3). 
        # But wait! Since we haven't run the formatting script yet (we cancelled task 675!), 
        # is the column at index 3 (Column D) or is it index 2?
        # Let's inspect the row values.
        # Let's check if 'Result' is in any of the elements, or specifically Column D (index 3).
        # Note: row[3] is Store Code if Column A (map) is still there. 
        # If Column A is there: map(0), RSM(1), ASM(2), Store Code(3).
        store_code = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
        store_name = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
        
        if "result" in store_code.lower() or "result" in store_name.lower():
            count_result += 1
            if count_result <= 5:
                print(f"  Sample Result row: {row}")
                
    wb.close()
    print(f"  Total rows: {count_total:,}")
    print(f"  'Result' rows found: {count_result:,}")

def main():
    for f in files:
        check_result_rows(f)

if __name__ == '__main__':
    main()
