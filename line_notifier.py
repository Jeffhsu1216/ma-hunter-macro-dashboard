"""
LINE Messaging API 推送模組
每日推送：傳送 Flex Message 儀表板摘要卡片到 LINE 官方帳號

使用端點：POST https://api.line.me/v2/bot/message/push
文件參考：https://developers.line.biz/en/reference/messaging-api/#send-push-message
"""

import os
import logging
import requests
from datetime import datetime
import pytz

logger = logging.getLogger(__name__)

# ── 環境變數（優先讀取），否則使用預設占位符 ──────────────────────────────
LINE_CHANNEL_TOKEN = os.environ.get("LINE_CHANNEL_TOKEN", "4a40a77c00340aec99e0312b65b49d5f")
LINE_USER_ID       = os.environ.get("LINE_USER_ID",       "Uffcb0e0752c2425f941466f4ba72a20e")
DASHBOARD_URL      = os.environ.get("DASHBOARD_URL",      "https://ma-hunter-macro-dashboard.onrender.com")

# LINE Messaging API v2 推送端點
LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"

# ── 主題色彩（深色主題，金色點綴） ──────────────────────────────────────
COLOR_BG        = "#1a1a2e"   # 卡片背景（深藍黑）
COLOR_HEADER_BG = "#16213e"   # Header 背景
COLOR_GOLD      = "#d4af37"   # 金色重點
COLOR_WHITE     = "#ffffff"   # 主文字
COLOR_GRAY      = "#a0a0b0"   # 次要文字
COLOR_GREEN     = "#00c853"   # 上漲
COLOR_RED       = "#ff1744"   # 下跌
COLOR_NEUTRAL   = "#78909c"   # 持平


# ─────────────────────────────────────────────────────────────────────────────
#  工具函式
# ─────────────────────────────────────────────────────────────────────────────

def _taipei_now() -> datetime:
    """回傳台北時區的當前時間"""
    return datetime.now(pytz.timezone("Asia/Taipei"))


def _weekday_zh(dt: datetime) -> str:
    """將 datetime 轉換為中文星期"""
    return {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}[dt.weekday()]


def _change_color(pct) -> str:
    """依漲跌幅決定顯示顏色"""
    if pct is None:
        return COLOR_NEUTRAL
    return COLOR_GREEN if float(pct) >= 0 else COLOR_RED


def _fmt_change(pct) -> str:
    """格式化漲跌幅，附加箭頭符號"""
    if pct is None:
        return "N/A"
    arrow = "▲" if float(pct) >= 0 else "▼"
    return f"{arrow}{abs(float(pct)):.2f}%"


def _fear_greed_label(score) -> str:
    """將 Fear & Greed 數值轉為中文標籤"""
    if not isinstance(score, (int, float)):
        return "N/A"
    score = int(score)
    if score < 25:   return "極度恐慌"
    if score < 45:   return "恐慌"
    if score < 55:   return "中性"
    if score < 75:   return "貪婪"
    return "極度貪婪"


def _line_headers() -> dict:
    """組裝 LINE API 請求標頭"""
    return {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_TOKEN}",
    }


# ─────────────────────────────────────────────────────────────────────────────
#  Flex Message 組裝
# ─────────────────────────────────────────────────────────────────────────────

def _build_row(label: str, value: str, value_color: str = COLOR_WHITE) -> dict:
    """
    建立一列「標籤 ─ 數值」的 box 元件
    用於股市指數、匯率等資料列
    """
    return {
        "type": "box",
        "layout": "horizontal",
        "margin": "sm",
        "contents": [
            {
                "type": "text",
                "text": label,
                "size": "sm",
                "color": COLOR_GRAY,
                "flex": 4,
            },
            {
                "type": "text",
                "text": value,
                "size": "sm",
                "color": value_color,
                "flex": 6,
                "align": "end",
                "weight": "bold",
            },
        ],
    }


def _build_section_header(title: str) -> dict:
    """建立區塊標題列（金色文字）"""
    return {
        "type": "text",
        "text": title,
        "size": "xs",
        "color": COLOR_GOLD,
        "weight": "bold",
        "margin": "md",
    }


def _build_separator() -> dict:
    """建立淡色分隔線"""
    return {
        "type": "separator",
        "margin": "md",
        "color": "#2a2a4a",
    }


def _build_flex_bubble(data: dict) -> dict:
    """
    組裝完整的 Flex Message Bubble Card
    包含：Header / 全球股市 / 匯率 / 市場情緒 / 三大法人 / Footer 按鈕
    """
    now        = _taipei_now()
    today_str  = now.strftime("%Y/%m/%d")
    weekday    = _weekday_zh(now)

    # ── 解析各資料區塊 ──────────────────────────────────────────────────────

    # 全球股市：S&P 500 / Nasdaq / 加權指數
    indices    = {i["name"]: i for i in data.get("indices", [])}
    sp         = indices.get("S&P 500", {})
    nq         = indices.get("Nasdaq", {})
    tw         = indices.get("加權指數", {})

    def idx_value(d: dict) -> str:
        if not d or d.get("price") is None:
            return "N/A"
        price = d.get("price_fmt", str(d.get("price", "")))
        pct   = d.get("change_pct")
        return f"{price}  {_fmt_change(pct)}"

    # 匯率：DXY / USD/TWD
    fx_map     = {f["name"]: f for f in data.get("fx", [])}
    dxy        = fx_map.get("DXY 美元指數", {})
    twd        = fx_map.get("USD/TWD", {})

    dxy_val    = dxy.get("price_fmt", "N/A")
    twd_val    = twd.get("price_fmt", "N/A")
    dxy_pct    = dxy.get("change_pct")
    twd_pct    = twd.get("change_pct")

    # 市場情緒：VIX + Fear & Greed
    vix_data   = indices.get("VIX", {})
    vix_val    = vix_data.get("price_fmt", "N/A") if vix_data else "N/A"
    fg         = data.get("fear_greed", {})
    fg_score   = fg.get("score", "N/A")
    fg_label   = _fear_greed_label(fg_score)
    fg_display = f"{fg_score}  {fg_label}" if fg_score != "N/A" else "N/A"

    # 三大法人合計（億元）
    tw_mkt     = data.get("taiwan") or {}
    total_yi   = tw_mkt.get("total_yi")
    tw_date    = tw_mkt.get("date", "")
    if total_yi is not None:
        arrow      = "▲" if total_yi >= 0 else "▼"
        yi_display = f"{arrow}{abs(total_yi):.1f} 億"
        yi_color   = COLOR_GREEN if total_yi >= 0 else COLOR_RED
        date_fmt   = f"{tw_date[:4]}/{tw_date[4:6]}/{tw_date[6:]}" if len(tw_date) == 8 else tw_date
        yi_label   = f"三大法人合計（{date_fmt}）"
    else:
        yi_display = "N/A"
        yi_color   = COLOR_NEUTRAL
        yi_label   = "三大法人合計"

    # ── 組裝 Bubble 結構 ────────────────────────────────────────────────────
    bubble = {
        "type": "bubble",
        "size": "kilo",
        "styles": {
            "body": {"backgroundColor": COLOR_BG},
            "footer": {"backgroundColor": "#0f0f1e"},
        },

        # ── Header ──────────────────────────────────────────────────────────
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": COLOR_HEADER_BG,
            "paddingAll": "16px",
            "contents": [
                {
                    "type": "text",
                    "text": "📊 總經儀表板",
                    "size": "lg",
                    "weight": "bold",
                    "color": COLOR_GOLD,
                },
                {
                    "type": "text",
                    "text": f"{today_str}（週{weekday}）",
                    "size": "xs",
                    "color": COLOR_GRAY,
                    "margin": "sm",
                },
            ],
        },

        # ── Body ────────────────────────────────────────────────────────────
        "body": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "16px",
            "spacing": "none",
            "contents": [

                # 全球股市
                _build_section_header("📈 全球股市"),
                _build_row(
                    "S&P 500",
                    idx_value(sp),
                    _change_color(sp.get("change_pct")) if sp else COLOR_NEUTRAL,
                ),
                _build_row(
                    "Nasdaq",
                    idx_value(nq),
                    _change_color(nq.get("change_pct")) if nq else COLOR_NEUTRAL,
                ),
                _build_row(
                    "台股加權",
                    idx_value(tw),
                    _change_color(tw.get("change_pct")) if tw else COLOR_NEUTRAL,
                ),

                _build_separator(),

                # 匯率
                _build_section_header("💵 匯率"),
                _build_row(
                    "DXY 美元指數",
                    f"{dxy_val}  {_fmt_change(dxy_pct)}",
                    _change_color(dxy_pct),
                ),
                _build_row(
                    "USD/TWD",
                    f"{twd_val}  {_fmt_change(twd_pct)}",
                    _change_color(twd_pct),
                ),

                _build_separator(),

                # 市場情緒
                _build_section_header("🧠 市場情緒"),
                _build_row("VIX 恐慌指數", vix_val, COLOR_WHITE),
                _build_row("Fear & Greed", fg_display, COLOR_WHITE),

                _build_separator(),

                # 三大法人
                _build_section_header("🏦 三大法人"),
                _build_row(yi_label, yi_display, yi_color),
            ],
        },

        # ── Footer ──────────────────────────────────────────────────────────
        "footer": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "12px",
            "contents": [
                {
                    "type": "button",
                    "style": "primary",
                    "color": COLOR_GOLD,
                    "height": "sm",
                    "action": {
                        "type": "uri",
                        "label": "點此開啟完整儀表板",
                        "uri": DASHBOARD_URL,
                    },
                }
            ],
        },
    }

    return bubble


# ─────────────────────────────────────────────────────────────────────────────
#  公開函式
# ─────────────────────────────────────────────────────────────────────────────

def send_line_text(text: str) -> bool:
    """
    傳送純文字訊息到 LINE 使用者

    Args:
        text: 要傳送的純文字內容

    Returns:
        bool: 傳送成功回傳 True，否則 False
    """
    payload = {
        "to": LINE_USER_ID,
        "messages": [
            {
                "type": "text",
                "text": text,
            }
        ],
    }
    try:
        resp   = requests.post(LINE_PUSH_URL, headers=_line_headers(), json=payload, timeout=15)
        result = resp.json()
        if resp.status_code == 200:
            logger.info("LINE 純文字訊息傳送成功")
            return True
        else:
            logger.error(f"LINE 純文字訊息傳送失敗：{resp.status_code}  {result}")
            return False
    except Exception as e:
        logger.error(f"LINE 純文字訊息例外錯誤：{e}")
        return False


def send_line_flex(dashboard_data: dict) -> bool:
    """
    傳送 Flex Message 儀表板摘要卡片到 LINE 使用者
    深色主題（#1a1a2e 背景，#d4af37 金色點綴）

    Args:
        dashboard_data: 來自 data_fetcher.fetch_all() 的完整數據字典

    Returns:
        bool: 傳送成功回傳 True，否則 False
    """
    bubble  = _build_flex_bubble(dashboard_data)
    payload = {
        "to": LINE_USER_ID,
        "messages": [
            {
                "type": "flex",
                "altText": "📊 總經儀表板每日摘要",
                "contents": bubble,
            }
        ],
    }
    try:
        resp   = requests.post(LINE_PUSH_URL, headers=_line_headers(), json=payload, timeout=15)
        result = resp.json()
        if resp.status_code == 200:
            logger.info("LINE Flex Message 傳送成功")
            return True
        else:
            logger.error(f"LINE Flex Message 傳送失敗：{resp.status_code}  {result}")
            return False
    except Exception as e:
        logger.error(f"LINE Flex Message 例外錯誤：{e}")
        return False


def push_daily_line(dashboard_data: dict) -> bool:
    """
    每日 LINE 推送主入口
    策略：優先嘗試 Flex Message，若失敗則降級為純文字

    Args:
        dashboard_data: 來自 data_fetcher.fetch_all() 的完整數據字典

    Returns:
        bool: 任一方式傳送成功即回傳 True
    """
    # ── 1. 嘗試 Flex Message ─────────────────────────────────────────────
    logger.info("嘗試傳送 LINE Flex Message...")
    if send_line_flex(dashboard_data):
        return True

    # ── 2. 降級：組裝純文字摘要 ──────────────────────────────────────────
    logger.warning("Flex Message 失敗，降級為純文字推送")
    now       = _taipei_now()
    today_str = now.strftime("%Y/%m/%d")
    weekday   = _weekday_zh(now)

    indices   = {i["name"]: i for i in dashboard_data.get("indices", [])}
    sp        = indices.get("S&P 500", {})
    nq        = indices.get("Nasdaq", {})
    tw        = indices.get("加權指數", {})
    fx_map    = {f["name"]: f for f in dashboard_data.get("fx", [])}
    dxy       = fx_map.get("DXY 美元指數", {})
    twd       = fx_map.get("USD/TWD", {})
    fg        = dashboard_data.get("fear_greed", {})
    tw_mkt    = dashboard_data.get("taiwan") or {}
    total_yi  = tw_mkt.get("total_yi")
    tw_date   = tw_mkt.get("date", "")

    def idx_txt(d):
        if not d or d.get("price") is None:
            return "N/A"
        return f"{d.get('price_fmt','N/A')}  {_fmt_change(d.get('change_pct'))}"

    yi_txt = "N/A"
    if total_yi is not None:
        arrow  = "▲" if total_yi >= 0 else "▼"
        d_fmt  = f"{tw_date[:4]}/{tw_date[4:6]}/{tw_date[6:]}" if len(tw_date) == 8 else tw_date
        yi_txt = f"{arrow}{abs(total_yi):.1f} 億（{d_fmt}）"

    text = (
        f"📊 總經儀表板  {today_str}（週{weekday}）\n"
        f"\n"
        f"📈 全球股市\n"
        f"  S&P 500    {idx_txt(sp)}\n"
        f"  Nasdaq     {idx_txt(nq)}\n"
        f"  台股加權   {idx_txt(tw)}\n"
        f"\n"
        f"💵 匯率\n"
        f"  DXY    {dxy.get('price_fmt','N/A')}  {_fmt_change(dxy.get('change_pct'))}\n"
        f"  TWD    {twd.get('price_fmt','N/A')}  {_fmt_change(twd.get('change_pct'))}\n"
        f"\n"
        f"🧠 Fear & Greed：{fg.get('score','N/A')}  {_fear_greed_label(fg.get('score'))}\n"
        f"\n"
        f"🏦 三大法人合計：{yi_txt}\n"
        f"\n"
        f"🔗 {DASHBOARD_URL}"
    )
    return send_line_text(text)


# ─────────────────────────────────────────────────────────────────────────────
#  工具：取得自己的 LINE User ID
# ─────────────────────────────────────────────────────────────────────────────

def get_my_user_id() -> None:
    """
    印出取得 LINE User ID 的操作說明

    方法一：Webhook 抓取（推薦，可在 Render 部署後使用）
      1. 在 LINE Developers Console 設定 Webhook URL：
         https://<your-render-app>.onrender.com/webhook/line
      2. 在對應 LINE 官方帳號傳送任意訊息給自己
      3. Webhook 收到事件後，從 event.source.userId 取得 User ID
      4. 將該 ID 設定為環境變數 LINE_USER_ID

    方法二：LINE Official Account Manager
      1. 登入 https://manager.line.biz/
      2. 選擇對應官方帳號 → 聊天
      3. 點擊使用者頭像 → 查看詳細資料，即可看到 User ID

    方法三：curl 測試（需已知 User ID 時驗證用）
      curl -X POST https://api.line.me/v2/bot/message/push \\
           -H "Authorization: Bearer <YOUR_CHANNEL_TOKEN>" \\
           -H "Content-Type: application/json" \\
           -d '{"to":"<USER_ID>","messages":[{"type":"text","text":"test"}]}'
    """
    print("=" * 60)
    print("取得 LINE User ID 的方式")
    print("=" * 60)
    print()
    print("【方法一：Webhook（推薦）】")
    print("  1. 至 LINE Developers Console 設定 Webhook URL")
    print("     → https://<app>.onrender.com/webhook/line")
    print("  2. 對 LINE 官方帳號傳送任意訊息")
    print("  3. 從 Webhook payload 的 event.source.userId 取得")
    print()
    print("【方法二：LINE Official Account Manager】")
    print("  1. 登入 https://manager.line.biz/")
    print("  2. 官方帳號 → 聊天 → 點擊使用者頭像 → 詳細資料")
    print()
    print("【方法三：curl 驗證】")
    print("  curl -X POST https://api.line.me/v2/bot/message/push \\")
    print("       -H 'Authorization: Bearer YOUR_TOKEN' \\")
    print("       -H 'Content-Type: application/json' \\")
    print("       -d '{\"to\":\"USER_ID\",\"messages\":[{\"type\":\"text\",\"text\":\"test\"}]}'")
    print()
    print(f"目前設定的 LINE_USER_ID = {LINE_USER_ID!r}")
    print(f"目前設定的 LINE_CHANNEL_TOKEN = {'(已設定)' if LINE_CHANNEL_TOKEN != 'YOUR_LINE_CHANNEL_TOKEN' else '(尚未設定)'}")


# ─────────────────────────────────────────────────────────────────────────────
#  直接執行時：測試推送
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import json
    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")

    if len(sys.argv) > 1 and sys.argv[1] == "user-id":
        # python line_notifier.py user-id → 印出 User ID 取得說明
        get_my_user_id()
    else:
        # 正常推送流程：清快取 → 抓數據 → 推送
        try:
            from data_fetcher import fetch_all, clear_cache
            clear_cache()
            data = fetch_all()
        except ImportError:
            logger.warning("data_fetcher 不可用，使用空字典測試")
            data = {}

        success = push_daily_line(data)
        sys.exit(0 if success else 1)
