import psutil
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_info']):
    try:
        if proc.info['name'] and 'python' in proc.info['name'].lower():
            print(f"PID: {proc.info['pid']}, Name: {proc.info['name']}, CPU: {proc.info['cpu_percent']}%, RAM: {proc.info['memory_info'].rss / 1024 / 1024:.1f} MB")
    except:
        pass
