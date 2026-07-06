import psutil
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

my_pid = os.getpid()
print(f"My PID: {my_pid}")

count = 0
for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
    try:
        if proc.info['name'] and 'python' in proc.info['name'].lower():
            pid = proc.info['pid']
            if pid != my_pid:
                cmdline = proc.info['cmdline']
                print(f"Killing python process PID {pid}: {cmdline}")
                proc.kill()
                count += 1
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        pass

print(f"Killed {count} other python processes.")
