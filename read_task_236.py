import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

log_path = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\.system_generated\tasks\task-236.log"
if os.path.exists(log_path):
    print("Found task-236.log! Content:")
    with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
        print(f.read())
else:
    print("task-236.log not found yet.")
