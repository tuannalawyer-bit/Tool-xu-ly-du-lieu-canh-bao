import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    print("Launching Excel...")
    xl = win32com.client.Dispatch("Excel.Application")
    print(f"Excel version: {xl.Version}")
    xl.Quit()
    print("Excel closed successfully!")
except Exception as e:
    print(f"Error launching Excel: {e}")
