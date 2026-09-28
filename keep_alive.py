# -*- coding: utf-8 -*-
"""
스트림릿 앱이 잠들지 않도록 주기적으로 방문해 깨우는 스크립트
"""
from playwright.sync_api import sync_playwright

APP_URL = "https://kepco-news-monitor-gbff2xm5nzatkkmvjsd9tm.streamlit.app"
WAKE_BUTTON_TEXT = "Yes, get this app back up!"

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        print(f"접속 시도: {APP_URL}")
        page.goto(APP_URL, timeout=30000)
        page.wait_for_timeout(3000)  # 화면 렌더링 대기

        wake_button = page.get_by_text(WAKE_BUTTON_TEXT, exact=False)
        if wake_button.count() > 0:
            print("😴 앱이 잠들어 있음 — 깨우기 버튼 클릭")
            wake_button.first.click()
            page.wait_for_timeout(60000)  # 완전히 깨어날 때까지 대기
            print("✅ 깨우기 완료")
        else:
            print("✅ 앱이 이미 깨어 있음")

        browser.close()

if __name__ == "__main__":
    main()
