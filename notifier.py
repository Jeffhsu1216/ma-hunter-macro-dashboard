"""
Telegram 推送模組
每日 08:00 台北時間推送總經儀表板摘要
"""

import requests
import os
from data_fetcher import fetch_all, clear_cache

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8743919766:AAG6z6YPW7Gqt7rF2KY2xC9mvbm2Ge31tjQ")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "2117347781")
DASHBOARD_URL = os.environ.get("DASHBOARD_URL", "https://ma-hunter-macro-dashboard.onrender.com")


def _change_icon(val):
    """漲跌符號"""
    if val is None:
        return ""
    if val > 0:
        return f"🟢 +{val}%"
    elif val < 0:
        return f"🔴 {val}%"
    return f"⚪ {val}%"


def _change_icon_bps(val):
    if val is None:
        return ""
    if val > 0:
        return f"🟢 +{val}bp"
    elif val < 0:
        return f"🔴 {val}bp"
    return f"⚪ {val}bp"


def build_message() -> str:
    """建立推送訊息文字"""
    clear_cache()
    data = fetch_all()

    lines = []
    lines.append(f"📊 <b>M&A Hunter 總經儀表板</b>")
    lines.append(f"📅 數據日期：{data['market_date']}")
    lines.append("")

    # 匯率
    lines.append("━━━ 💱 匯率與利率 ━━━")
    lines.append(f"Fed Rate: <b>{data['rates']['fed_rate']}%</b>")
    for fx in data["fx"]:
        lines.append(f"  {fx['name']}  <b>{fx['price_fmt']}</b>  {_change_icon(fx['change_pct'])}")

    r = data["rates"]
    if r["ten_y"]:
        lines.append(f"  US 10Y  <b>{r['ten_y']:.2f}%</b>  {_change_icon_bps(r['ten_y_chg_bps'])}")
    if r["two_y"]:
        lines.append(f"  US 2Y   <b>{r['two_y']:.2f}%</b>  {_change_icon_bps(r['two_y_chg_bps'])}")
    if r["spread_bps"] is not None:
        lines.append(f"  利差     <b>{int(r['spread_bps'])}bps</b>")
    lines.append("")

    # 股市
    lines.append("━━━ 📈 全球股市 ━━━")
    for idx in data["indices"]:
        name = "📉 VIX" if idx.get("is_vix") else f"{idx['flag']} {idx['name']}"
        lines.append(f"  {name}  <b>{idx['price_fmt']}</b>  {_change_icon(idx['change_pct'])}")
    lines.append("")

    # 原物料
    lines.append("━━━ 🛢️ 原物料 ━━━")
    for c in data["commodities"]:
        lines.append(f"  {c['name']}  <b>${c['price_fmt']}</b>  {_change_icon(c['change_pct'])}")
    lines.append("")

    # Footer
    lines.append(f"🔗 <a href=\"{DASHBOARD_URL}\">開啟完整儀表板</a>")

    return "\n".join(lines)


def send_telegram(text: str) -> bool:
    """發送 Telegram 訊息"""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    resp = requests.post(url, json=payload)
    return resp.json().get("ok", False)


def push_daily():
    """每日推送（由排程器呼叫）"""
    msg = build_message()
    success = send_telegram(msg)
    if success:
        print(f"✅ Telegram 推送成功")
    else:
        print(f"❌ Telegram 推送失敗")
    return success


if __name__ == "__main__":
    push_daily()
