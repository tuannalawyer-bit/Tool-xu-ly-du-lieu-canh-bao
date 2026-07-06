import os
path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\dashboard.html"
if os.path.exists(path):
    try:
        os.remove(path)
        print("Removed old dashboard.html successfully.")
    except Exception as e:
        print(f"Error: {e}")
else:
    print("old dashboard.html not found.")
