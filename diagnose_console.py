import sys
import asyncio
from playwright.async_api import async_playwright

async def main():
    html_path = r"c:\Users\Tuan\OneDrive - WIN\Desktop\công việc thực hiện\Thang 7\Đổ dữ liệu\dashboard_win.html"
    print(f"Loading {html_path} in headless browser...")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        
        # Capture console messages
        page.on("console", lambda msg: print(f"CONSOLE {msg.type.upper()}: {msg.text}"))
        page.on("pageerror", lambda err: print(f"PAGE ERROR: {err}"))
        
        try:
            await page.goto(f"file:///{html_path}", timeout=10000)
            await page.wait_for_timeout(2000) # wait a bit for scripts to run
            title = await page.title()
            print(f"Loaded successfully. Title: {title}")
        except Exception as e:
            print(f"Failed to load: {e}")
            
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
