"""
Telegram 推送模組 — 本機渲染截圖版
本機抓最新數據 → 渲染 HTML → Playwright 截圖 → Telegram sendPhoto
完全繞開 Render cache，確保數據即時。
"""

import requests
import os
import asyncio
import logging
import tempfile
from datetime import datetime
import pytz

logger = logging.getLogger(__name__)

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8743919766:AAG6z6YPW7Gqt7rF2KY2xC9mvbm2Ge31tjQ")
TELEGRAM_CHAT_ID   = os.environ.get("TELEGRAM_CHAT_ID", "2117347781")
DASHBOARD_URL      = os.environ.get("DASHBOARD_URL", "https://ma-hunter-macro-dashboard.onrender.com")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", "dashboard.html")


def _render_html(data: dict) -> str:
    """用 Jinja2 本機渲染 dashboard.html"""
    from jinja2 import Environment, FileSystemLoader
    taipei_tz = pytz.timezone("Asia/Taipei")
    now = datetime.now(taipei_tz)
    weekday_map = {0:"一",1:"二",2:"三",3:"四",4:"五",5:"六",6:"日"}

    env = Environment(loader=FileSystemLoader(os.path.join(os.path.dirname(__file__), "templates")))
    tmpl = env.get_template("dashboard.html")
    return tmpl.render(
        data=data,
        today=now.strftime("%Y/%m/%d"),
        weekday=weekday_map[now.weekday()],
        is_weekend=now.weekday() >= 5,
    )


async def _screenshot_html(html: str) -> bytes | None:
    """將 HTML 字串寫入暫存檔，用 Playwright 截圖"""
    try:
        from playwright.async_api import async_playwright
        with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8") as f:
            f.write(html)
            tmp_path = f.name

        async with async_playwright() as p:
            browser = await p.chromium.launch(
                args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
            )
            page = await browser.new_page(viewport={"width": 1440, "height": 900})
            await page.goto(f"file://{tmp_path}", wait_until="networkidle")
            await page.wait_for_timeout(1000)
            screenshot = await page.screenshot(full_page=True)
            await browser.close()

        os.unlink(tmp_path)
        return screenshot
    except Exception as e:
        logger.error(f"Screenshot failed: {e}")
        return None


def take_screenshot(html: str) -> bytes | None:
    return asyncio.run(_screenshot_html(html))


def send_telegram_photo(image_bytes: bytes, caption: str) -> bool:
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
    resp = requests.post(url,
        data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption, "parse_mode": "HTML"},
        files={"photo": ("dashboard.png", image_bytes, "image/png")},
        timeout=30,
    )
    return resp.json().get("ok", False)


def send_telegram_text(text: str) -> bool:
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    resp = requests.post(url, json={
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": False,
    }, timeout=15)
    return resp.json().get("ok", False)


def push_daily() -> bool:
    """每日推送：本機抓數據 → 渲染 → 截圖 → Telegram"""
    from data_fetcher import fetch_all, clear_cache
    clear_cache()
    data = fetch_all()

    html = _render_html(data)
    screenshot = take_screenshot(html)

    if screenshot:
        caption = (
            f'📊 <b>M&A Hunter 總經儀表板</b>\n'
            f'🔗 <a href="{DASHBOARD_URL}">開啟互動版</a>'
        )
        success = send_telegram_photo(screenshot, caption)
    else:
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
