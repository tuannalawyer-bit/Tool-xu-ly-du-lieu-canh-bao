import subprocess
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    res = subprocess.run('tasklist', shell=True, capture_output=True, text=True)
    lines = res.stdout.splitlines()
    print("Python processes in tasklist:")
    found = False
    for line in lines:
        if "python" in line.lower():
            print(line)
            found = True
    if not found:
        print("No python processes found in tasklist.")
except Exception as e:
    print(f"Error: {e}")
