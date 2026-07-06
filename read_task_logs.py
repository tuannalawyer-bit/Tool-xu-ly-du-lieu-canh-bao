import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

tasks_dir = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\.system_generated\tasks"
try:
    files = os.listdir(tasks_dir)
    print("Files in tasks directory:")
    for f in sorted(files):
        path = os.path.join(tasks_dir, f)
        size = os.path.getsize(path)
        print(f"  {f} ({size} bytes)")
        if "84" in f:
            print("--- CONTENT OF task-84.log ---")
            with open(path, 'r', encoding='utf-8', errors='ignore') as log_f:
                print(log_f.read())
            print("------------------------------")
except Exception as e:
    print(f"Error: {e}")
