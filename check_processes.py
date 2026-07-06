import subprocess
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    # Use wmic process get commandline to list python processes on Windows
    cmd = 'wmic process where "name=\'python.exe\'" get commandline,processid'
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    print("Running Python processes:")
    print(res.stdout)
except Exception as e:
    print(f"Error: {e}")
