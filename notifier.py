"""
Telegram 推送模組 — 互動連結版
每日推送：傳送儀表板連結 + 數據摘要到 Telegram
點擊後在手機/電腦瀏覽器直接開啟最新互動版儀表板
"""

import requests
import os
import logging
from datetime import datetime
import pytz

logger = logging.getLogger(__name__)

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8743919766:AAG6z6YPW7Gqt7rF2KY2xC9mvbm2Ge31tjQ")
TELEGRAM_CHAT_ID   = os.environ.get("TELEGRAM_CHAT_ID", "2117347781")
DASHBOARD_URL      = os.environ.get("DASHBOARD_URL", "https://ma-hunter-macro-dashboard.onrender.com")


def _warmup_render(timeout: int = 60) -> bool:
    """預熱 Render：先打一次 /api/refresh，等冷啟動完成"""
    try:
        logger.info("Warming up Render...")
        resp = requests.get(f"{DASHBOARD_URL}/api/refresh", timeout=timeout)
        ok = resp.status_code == 200
        logger.info(f"Render warmup {'OK' if ok else 'failed'}: {resp.status_code}")
        return ok
    except Exception as e:
        logger.warning(f"Render warmup timeout/error: {e}")
        return False


def _build_message(data: dict) -> str:
    """組裝 Telegram 推送文字：日期 + 摘要 + 連結按鈕"""
    taipei_tz = pytz.timezone("Asia/Taipei")
    now = datetime.now(taipei_tz)
    weekday_map = {0:"一",1:"二",2:"三",3:"四",4:"五",5:"六",6:"日"}
    today_str  = now.strftime("%Y/%m/%d")
    weekday    = weekday_map[now.weekday()]
    market_date = data.get("market_date", today_str)

    # 指數摘要（S&P + Nasdaq + 台股）
    indices = {i["name"]: i for i in data.get("indices", [])}
    sp  = indices.get("S&P 500", {})
    nq  = indices.get("Nasdaq", {})
    tw  = indices.get("加權指數", {})

    def fmt_idx(d):
        if not d or d.get("price") is None:
            return "N/A"
        c = d.get("change_pct", 0) or 0
        arrow = "▲" if c >= 0 else "▼"
        return f"{d['price_fmt']} {arrow}{abs(c):.2f}%"

    # 匯率摘要
    fx_map = {f["name"]: f for f in data.get("fx", [])}
    dxy  = fx_map.get("DXY 美元指數", {})
    twd  = fx_map.get("USD/TWD", {})

    # Fear & Greed
    fg = data.get("fear_greed", {})
    fg_score  = fg.get("score", "N/A")
    fg_rating = fg.get("rating", "N/A")
    fg_emoji  = "😱" if isinstance(fg_score, int) and fg_score < 25 else \
                "😨" if isinstance(fg_score, int) and fg_score < 45 else \
                "😐" if isinstance(fg_score, int) and fg_score < 55 else \
                "😄" if isinstance(fg_score, int) and fg_score < 75 else "🤑"

    # 三大法人
    tw_mkt = data.get("taiwan") or {}
    total_yi = tw_mkt.get("total_yi")
    tw_date  = tw_mkt.get("date", "")
    tw_line  = f"{'▲' if total_yi and total_yi >= 0 else '▼'}{abs(total_yi):.1f}億（{tw_date[:4]}/{tw_date[4:6]}/{tw_date[6:]}）" \
               if total_yi is not None else "N/A"

    # 央行利率
    cb = data.get("cb_rates", {})
    fed_rate = cb.get("Fed", {}).get("rate", "N/A")
    fed_next = cb.get("Fed", {}).get("next", "")

    msg = (
        f"📊 <b>M&A Hunter 總經儀表板</b>\n"
        f"<i>{today_str}（{weekday}）｜數據基準：{market_date}</i>\n"
        f"\n"
        f"📈 <b>全球指數</b>\n"
        f"  S&P 500　　{fmt_idx(sp)}\n"
        f"  Nasdaq　　{fmt_idx(nq)}\n"
        f"  加權指數　{fmt_idx(tw)}\n"
        f"\n"
        f"💵 <b>匯率</b>\n"
        f"  DXY {dxy.get('price_fmt','N/A')}　USD/TWD {twd.get('price_fmt','N/A')}\n"
        f"\n"
        f"{fg_emoji} <b>市場情緒</b>：{fg_score} — {fg_rating}\n"
        f"\n"
        f"🏦 <b>Fed</b>：{fed_rate}%　下次：{fed_next}\n"
        f"\n"
        f"🇹🇼 <b>三大法人合計</b>：{tw_line}\n"
        f"\n"
        f"🔗 <a href=\"{DASHBOARD_URL}\">點此開啟完整儀表板 →</a>"
    )
    return msg


def send_telegram_link(text: str) -> bool:
    """傳送含連結的 Telegram 訊息（含 inline 按鈕）"""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
        "reply_markup": {
            "inline_keyboard": [[
                {"text": "📊 開啟儀表板", "url": DASHBOARD_URL}
            ]]
        }
    }
    resp = requests.post(url, json=payload, timeout=15)
    result = resp.json()
    if not result.get("ok"):
        logger.error(f"Telegram send failed: {result}")
    return result.get("ok", False)


def push_daily() -> bool:
    """每日推送主流程：預熱 Render → 抓新數據 → 傳連結+摘要"""
    from data_fetcher import fetch_all, clear_cache

    # 1. 清快取、抓最新數據（本機）
    clear_cache()
    data = fetch_all()

    # 2. 預熱 Render（讓使用者點連結時無需等冷啟動）
    _warmup_render(timeout=60)

    # 3. 組裝訊息並推送
    msg = _build_message(data)
    success = send_telegram_link(msg)

    if success:
        logger.info("Telegram push succeeded")
    else:
        logger.error("Telegram push failed")
    return success


if __name__ == "__main__":
    push_daily()
