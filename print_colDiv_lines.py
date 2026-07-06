import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

with open(r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\dashboard_win_template.html", "r", encoding="utf-8") as f:
    content = f.read()

# Let's print out lines around line 750 where colDiv is created
lines = content.split("\n")
for idx in range(730, 785):
    if idx < len(lines):
        print(f"{idx+1}: {lines[idx]}")
