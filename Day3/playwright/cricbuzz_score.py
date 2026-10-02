import time
from playwright.sync_api import sync_playwright
from datetime import datetime
with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("https://www.cricbuzz.com/live-cricket-scores/163077/roi-vs-jk-irani-cup-irani-cup-2026")
    #time.sleep(5)
    page.wait_for_selector("div.flex.flex-row.font-bold.text-xl") # the selector you found
    score = page.inner_text("div.flex.flex-row.font-bold.text-xl")
    print(score)
    page.screenshot(path="score.png")
    browser.close()