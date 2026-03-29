"""
Telegram 推送模組 — 截圖版
每日 08:00 台北時間截圖儀表板並推送至 Telegram
"""

import requests
import os
import asyncio
import logging

logger = logging.getLogger(__name__)

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8743919766:AAG6z6YPW7Gqt7rF2KY2xC9mvbm2Ge31tjQ")
TELEGRAM_CHAT_ID   = os.environ.get("TELEGRAM_CHAT_ID", "2117347781")
DASHBOARD_URL      = os.environ.get("DASHBOARD_URL", "https://ma-hunter-macro-dashboard.onrender.com")


async def _take_screenshot(url: str) -> bytes | None:
    """使用 Playwright 截圖儀表板"""
    try:
        from playwright.async_api import async_playwright
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
            )
            page = await browser.new_page(viewport={"width": 1440, "height": 900})
            await page.goto(url, wait_until="networkidle", timeout=30000)
            # 等待數據渲染
            await page.wait_for_timeout(2000)
            screenshot = await page.screenshot(full_page=True)
            await browser.close()
            return screenshot
    except Exception as e:
        logger.error(f"Screenshot failed: {e}")
        return None


def take_screenshot(url: str) -> bytes | None:
    """同步包裝"""
    return asyncio.run(_take_screenshot(url))


def send_telegram_photo(image_bytes: bytes, caption: str) -> bool:
    """發送截圖至 Telegram"""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
    resp = requests.post(url,
        data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption, "parse_mode": "HTML"},
        files={"photo": ("dashboard.png", image_bytes, "image/png")},
        timeout=30,
    )
    return resp.json().get("ok", False)


def send_telegram_text(text: str) -> bool:
    """發送純文字（截圖失敗時的 fallback）"""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    resp = requests.post(url, json={
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": False,
    }, timeout=15)
    return resp.json().get("ok", False)


def push_daily() -> bool:
    """每日推送（由排程器呼叫）"""
    # 強制刷新 Render 端 cache，確保截圖是最新數據
    try:
        resp = requests.get(f"{DASHBOARD_URL}/api/refresh", timeout=90)
        logger.info(f"Remote cache refreshed: {resp.status_code}")
    except Exception as e:
        logger.warning(f"Remote refresh failed: {e}")

    import time; time.sleep(3)  # 等數據抓完

    # 嘗試截圖
    screenshot = take_screenshot(DASHBOARD_URL)

    if screenshot:
        caption = (
            f'📊 <b>M&A Hunter 總經儀表板</b>\n'
            f'🔗 <a href="{DASHBOARD_URL}">開啟互動版</a>'
        )
        success = send_telegram_photo(screenshot, caption)
    else:
        # Fallback: 文字版
        logger.warning("Screenshot failed, falling back to text")
        success = send_telegram_text(
            f'📊 <b>M&A Hunter 總經儀表板</b>\n'
            f'🔗 <a href="{DASHBOARD_URL}">開啟儀表板</a>\n'
            f'<i>（截圖服務暫時不可用）</i>'
        )

    if success:
        logger.info("Telegram push succeeded")
    else:
        logger.error("Telegram push failed")
    return success


if __name__ == "__main__":
    push_daily()
