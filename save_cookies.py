"""
运行此脚本，会打开一个真实浏览器窗口。
请在窗口中完成阳光采购平台的登录，登录成功看到招标列表后，
回到终端按 Enter，脚本会自动保存 cookie 到 cookies_yangguang.json。

用法：
    uv run python save_cookies.py
"""
import json
from playwright.sync_api import sync_playwright

TARGET = "https://zc.szaee.com/#/project?tradeType=4&projectClass="
COOKIE_FILE = "cookies_yangguang.json"

with sync_playwright() as play:
    browser = play.chromium.launch(headless=False)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        ignore_https_errors=True,
    )
    page = context.new_page()
    page.goto(TARGET)

    print("请在弹出的浏览器窗口中完成登录...")
    print("登录成功并看到招标列表后，回到这里按 Enter 保存 cookie。")
    input()

    cookies = context.cookies()
    with open(COOKIE_FILE, "w", encoding="utf-8") as f:
        json.dump(cookies, f, ensure_ascii=False, indent=2)
    print(f"✅ 已保存 {len(cookies)} 条 cookie 到 {COOKIE_FILE}")
    browser.close()
