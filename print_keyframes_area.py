import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

with open(r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\dashboard_win_template.html", "r", encoding="utf-8") as f:
    content = f.read()

lines = content.split("\n")
# Print lines 380 to 440
for idx in range(380, 440):
    if idx < len(lines):
        print(f"{idx+1}: {lines[idx]}")
