import urllib.request
import os
import ssl

url = "https://cdn.jsdelivr.net/npm/chart.js"
dest = r"C:\Users\Tuan\.gemini\antigravity\brain\b053fc36-8fe6-4316-bba8-ac8fef252caf\scratch\chart.min.js"

try:
    print(f"Downloading Chart.js (ignoring SSL verification) from {url}...")
    context = ssl._create_unverified_context()
    req = urllib.request.Request(
        url, 
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    )
    with urllib.request.urlopen(req, context=context) as response:
        with open(dest, 'wb') as f:
            f.write(response.read())
    print(f"Successfully downloaded Chart.js to {dest}! Size: {os.path.getsize(dest)} bytes")
except Exception as e:
    print(f"Error downloading Chart.js: {e}")
