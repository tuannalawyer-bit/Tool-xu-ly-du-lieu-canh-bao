import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

filepath = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\Tong_hop_Kiem_tra_Win.xlsx"

print(f"Checking if {filepath} is accessible...")
try:
    with open(filepath, 'rb') as f:
        # read 10 bytes
        data = f.read(10)
        print(f"Success! Read: {data}")
except Exception as e:
    print(f"Error: {e}")
