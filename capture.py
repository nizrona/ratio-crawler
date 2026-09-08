import os
from datetime import datetime
from zoneinfo import ZoneInfo  # 파이썬 표준 라이브러리 (추가 설치 불필요)
from playwright.sync_api import sync_playwright

URL = "https://ratio.uwayapply.com/Sl5KMCYlODlKXiUmOiZKN2ZUZg=="

def capture_screenshot():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1280, "height": 1024})

        page.goto(URL, wait_until="networkidle")
        page.wait_for_timeout(3000)

        # 서울 타임존(KST, UTC+9) 기준으로 시각 생성
        kst = ZoneInfo("Asia/Seoul")
        today = datetime.now(kst).strftime("%Y%m%d_%H%M")

        os.makedirs("screenshots", exist_ok=True)
        file_path = f"screenshots/ratio_{today}.png"

        page.screenshot(path=file_path, full_page=True)
        print(f"Captured: {file_path}")
        browser.close()

if __name__ == "__main__":
    capture_screenshot()
