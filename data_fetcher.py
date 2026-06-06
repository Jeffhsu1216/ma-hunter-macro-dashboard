"""
總經儀表板 — 數據抓取模組 v3
M&A Hunter Macro Dashboard Data Fetcher

數據來源：
  匯率/指數/原物料/VIX — Yahoo Finance (yfinance) 批次下載
  殖利率曲線           — FRED (DGS2/DGS5/DGS10/DGS30)
  央行利率             — FRED (DFEDTARL/DFEDTARU) + FOMC/CBC 日程自動選
  Fear & Greed         — CNN API + VIX 備援
  經濟日曆             — ForexFactory JSON
  三大法人             — TWSE BFI82U

v3 更新：
  - yfinance 改批次 download()（一次 HTTP 呼叫）
  - 各區塊新增 commentary 解釋文字
  - 指數按漲跌幅排序
"""

import yfinance as yf
import requests
import pytz
from datetime import datetime, timedelta, date
import json, os, logging, re, urllib.request

logger = logging.getLogger(__name__)

# ============================================================
# 路徑
# ============================================================
CACHE_FILE       = os.path.join(os.path.dirname(__file__), "cache_data.json")
CALENDAR_BACKUP  = os.path.join(os.path.dirname(__file__), "calendar_backup.json")
TAIWAN_BACKUP    = os.path.join(os.path.dirname(__file__), "taiwan_backup.json")
GEOPOLITICS_FILE = os.path.join(os.path.dirname(__file__), "geopolitics.json")
CACHE_TTL_HOURS  = 1
TAIPEI_TZ        = pytz.timezone("Asia/Taipei")

# ============================================================
# 數據定義
# ============================================================

# (name, ticker, decimals, mode)
# mode:
#   "dxy"        — 美元指數（不轉 TWD，直接顯示原值）
#   "usd_twd"    — USD/TWD 本體（ticker 即 TWD=X，raw 就是 USD/TWD）
#   "usd_base"   — yfinance 給 USD/X（如 CNY=X / JPY=X），轉成 X/TWD = (USD/TWD) / (USD/X)
#   "quote_base" — yfinance 給 X/USD（如 EURUSD=X），轉成 X/TWD = (X/USD) × (USD/TWD)
# 全部匯率從 TWD 視角呈現：X/TWD ↑ = X 對 TWD 升值（TWD 貶值）
FX_TICKERS = [
    ("美元指數 (DXY)",       "DX-Y.NYB", 2, "dxy"),
    ("美元 (USD/TWD)",       "TWD=X",    3, "usd_twd"),
    ("人民幣 (CNY/TWD)",     "CNY=X",    3, "usd_base"),
    ("歐元 (EUR/TWD)",       "EURUSD=X", 3, "quote_base"),
    ("英鎊 (GBP/TWD)",       "GBPUSD=X", 3, "quote_base"),
    ("瑞士法郎 (CHF/TWD)",   "CHF=X",    3, "usd_base"),
    ("日圓 (JPY/TWD)",       "JPY=X",    4, "usd_base"),
    ("韓元 (KRW/TWD)",       "KRW=X",    5, "usd_base"),
    ("澳幣 (AUD/TWD)",       "AUDUSD=X", 3, "quote_base"),
]

# 殖利率用 FRED（官方來源，精準）
YIELD_FRED_IDS = [
    ("2Y",  "DGS2"),
    ("5Y",  "DGS5"),
    ("10Y", "DGS10"),
    ("30Y", "DGS30"),
]

INDEX_TICKERS = [
    ("標普 500 (S&P 500)",        "^GSPC",     "🇺🇸"),
    ("那斯達克 (Nasdaq)",          "^IXIC",     "🇺🇸"),
    ("費城半導體 (SOX)",           "^SOX",      "🇺🇸"),
    ("英國富時 100 (FTSE 100)",    "^FTSE",     "🇬🇧"),
    ("德國 DAX (DAX)",             "^GDAXI",    "🇩🇪"),
    ("歐洲 STOXX 600 (STOXX 600)", "^STOXX",    "🇪🇺"),
    ("日經 225 (Nikkei 225)",      "^N225",     "🇯🇵"),
    ("南韓綜合 (KOSPI)",            "^KS11",     "🇰🇷"),
    ("恆生指數 (HSI)",              "^HSI",      "🇭🇰"),
    ("上證綜合 (SSEC)",             "000001.SS", "🇨🇳"),
    ("印度 Nifty 50 (Nifty 50)",   "^NSEI",     "🇮🇳"),
    ("加權指數 (TAIEX)",            "^TWII",     "🇹🇼"),
]

# (name, ticker, decimals, currency_symbol)
COMMODITY_TICKERS = [
    ("WTI 原油",   "CL=F",    2, "$"),
    ("布蘭特原油", "BZ=F",    2, "$"),
    ("天然氣",     "NG=F",    3, "$"),
    ("黃金",       "GC=F",    0, "$"),
    ("白銀",       "SI=F",    2, "$"),
    ("銅",         "HG=F",    3, "$"),
]

# 加密貨幣（CoinGecko）
CRYPTO_LIST = [
    ("BTC", "bitcoin",  0),
    ("ETH", "ethereum", 2),
]

VIX_TICKER = "^VIX"

# 2026 全年 FOMC 會議日程（結束日 / release date，源：federalreserve.gov）
FOMC_DATES_2026 = [
    date(2026, 1, 28), date(2026, 3, 18), date(2026, 4, 29),
    date(2026, 6, 17), date(2026, 7, 29), date(2026, 9, 16),
    date(2026, 10, 28), date(2026, 12, 9),
]

# 2026 全年 ECB 理事會決策會議日程（決議日，源：ecb.europa.eu）
ECB_DATES_2026 = [
    date(2026, 1, 29), date(2026, 3, 19), date(2026, 4, 30),
    date(2026, 6, 11), date(2026, 7, 23), date(2026, 9, 10),
    date(2026, 10, 29), date(2026, 12, 17),
]

# 2026 全年 BOJ 金融政策決定會合日程（結束日）
BOJ_DATES_2026 = [
    date(2026, 1, 23), date(2026, 3, 19), date(2026, 4, 28),
    date(2026, 6, 16), date(2026, 7, 31), date(2026, 9, 18),
    date(2026, 10, 30), date(2026, 12, 18),
]

# 2026 全年 BoE MPC 會議日程（決議日，源：bankofengland.co.uk）
BOE_DATES_2026 = [
    date(2026, 2, 5),  date(2026, 3, 19), date(2026, 5, 7),
    date(2026, 6, 18), date(2026, 8, 6),  date(2026, 9, 17),
    date(2026, 11, 5), date(2026, 12, 17),
]

# 2026 全年 CBC（台灣央行）理監事會日程
CBC_DATES_2026 = [
    date(2026, 3, 19), date(2026, 6, 18),
    date(2026, 9, 17), date(2026, 12, 17),
]

# PBOC 沒有固定會議日程；LPR 每月 20 日例行公布（遇假日順延）
# OMO（7 天逆回購）為主要政策利率，與 LPR 同步公布或央行臨時調整


# ============================================================
# 核心工具
# ============================================================

def _get_quote(ticker_symbol: str, retries: int = 2) -> dict:
    """Yahoo Finance v8 Chart API（直接 HTTP，不用 yfinance 避免 rate limit）

    修正：優先使用 meta.regularMarketPrice 當現價（避免節假日/時區造成 closes[-1] 落後）
         前收盤優先用 meta.chartPreviousClose，fallback 到 closes[-2]
    """
    import urllib.request as _urlr, urllib.parse as _urlp, time
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{_urlp.quote(ticker_symbol)}?interval=1d&range=1d")
    for attempt in range(retries + 1):
        try:
            req = _urlr.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with _urlr.urlopen(req, timeout=10) as r:
                d = json.loads(r.read())
            result = d["chart"]["result"][0]
            meta   = result["meta"]

            # 現價：優先用 regularMarketPrice（即時/最新），避免節假日 None 問題
            p = meta.get("regularMarketPrice")

            # 前收盤：chartPreviousClose 是最後一個完整 session 的收盤
            prev = meta.get("chartPreviousClose")

            # fallback：若 meta 欄位缺失，才回頭用 closes 陣列
            if p is None or prev is None:
                closes = result["indicators"]["quote"][0]["close"]
                closes = [x for x in closes if x is not None]
                if not closes:
                    raise ValueError("no valid closes")
                if p    is None: p    = closes[-1]
                if prev is None: prev = closes[-2] if len(closes) > 1 else closes[-1]

            chg  = p - prev
            chgp = chg / prev * 100 if prev else 0
            return {
                "price":      round(float(p),    6),
                "prev_close": round(float(prev),  6),
                "change_pct": round(float(chgp),  2),
            }
        except Exception as e:
            logger.warning(f"_get_quote {ticker_symbol} attempt {attempt+1}: {e}")
            if attempt < retries:
                time.sleep(0.4)
    return {"price": None, "prev_close": None, "change_pct": None}


def _get_history(ticker_symbol: str, rng: str = "1mo") -> list:
    """抓 Yahoo Chart 日線收盤序列，回傳 [{date: 'MM/DD', value: float}, ...]"""
    import urllib.request as _urlr, urllib.parse as _urlp, datetime as _dt
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{_urlp.quote(ticker_symbol)}?interval=1d&range={rng}")
    try:
        req = _urlr.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with _urlr.urlopen(req, timeout=10) as r:
            d = json.loads(r.read())
        result = d["chart"]["result"][0]
        ts     = result["timestamp"]
        closes = result["indicators"]["quote"][0]["close"]
        out = []
        for t, c in zip(ts, closes):
            if c is None:
                continue
            out.append({"date": _dt.datetime.fromtimestamp(t).strftime("%m/%d"),
                        "value": round(float(c), 2)})
        return out
    except Exception as e:
        logger.warning(f"_get_history {ticker_symbol} {rng} failed: {e}")
        return []


def _fmt(price, decimals=2):
    if price is None:
        return "N/A"
    if price >= 10000:
        return f"{price:,.0f}"
    return f"{price:,.{decimals}f}"


def _get_fred_csv(series_id: str, n_rows: int = 10, retries: int = 2) -> list:
    """從 FRED 抓 CSV，回傳最後 n 筆 [(date_str, value_float), ...]；含 retry"""
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
    last_err = None
    for attempt in range(retries + 1):
        try:
            resp = requests.get(url, timeout=(10, 30), headers={"User-Agent": "Mozilla/5.0"})
            if resp.status_code != 200:
                last_err = f"HTTP {resp.status_code}"
                continue
            lines = resp.text.strip().split("\n")[1:]
            results = []
            for line in lines[-n_rows:]:
                parts = line.split(",")
                if len(parts) == 2 and parts[1] != ".":
                    results.append((parts[0], float(parts[1])))
            if results:
                return results
        except Exception as e:
            last_err = str(e)
    logger.warning(f"FRED {series_id} failed after {retries+1} tries: {last_err}")
    return []


def _read_last_cb_rate(name: str):
    """從上次 cache_data.json 讀取已成功抓到的央行利率（fallback 用）。
    回傳 (rate_str, source_str) 或 (None, None)。"""
    try:
        if not os.path.exists(CACHE_FILE):
            return (None, None)
        with open(CACHE_FILE, "r") as f:
            cache = json.load(f)
        rate = cache.get("data", {}).get("cb_rates", {}).get(name, {}).get("rate")
        ts = cache.get("timestamp", "")
        return (rate, ts[:10] if ts else None)
    except Exception as e:
        logger.warning(f"Read last {name} rate from cache failed: {e}")
        return (None, None)


def _scrape_fed_official() -> str:
    """爬 federalreserve.gov 抓最新 FOMC 目標區間（FRED fallback）"""
    try:
        import re as _re
        resp = requests.get(
            "https://www.federalreserve.gov/monetarypolicy/openmarket.htm",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        # 找所有 "X.XX-X.XX" 格式的目標區間，page 由近至遠排序，第一個即最新
        m = _re.search(r'<td[^>]*>(\d+\.\d+)-(\d+\.\d+)</td>', resp.text)
        if m:
            return f"{float(m.group(1)):.2f}–{float(m.group(2)):.2f}"
    except Exception as e:
        logger.warning(f"Fed official scrape failed: {e}")
    return ""


def _scrape_ecb_official() -> str:
    """爬 ecb.europa.eu 抓最新 Deposit facility rate（FRED fallback）"""
    try:
        import re as _re
        resp = requests.get(
            "https://www.ecb.europa.eu/stats/policy_and_exchange_rates/key_ecb_interest_rates/html/index.en.html",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        # 抓 Deposit facility 表格的第一個 <tr>
        m = _re.search(r'Deposit facility.*?</thead>(.*?)</table>', resp.text, _re.DOTALL)
        if m:
            first_row = _re.search(r'<tr[^>]*>(.*?)</tr>', m.group(1), _re.DOTALL)
            if first_row:
                cells = _re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', first_row.group(1), _re.DOTALL)
                cells = [_re.sub(r'<[^>]+>', ' ', c).strip() for c in cells]
                # 表格欄位：年、日期、Deposit、Main refi (low)、-、Main refi (high)
                for c in cells:
                    if _re.match(r'^\d+\.\d+$', c):
                        return f"{float(c):.2f}"
    except Exception as e:
        logger.warning(f"ECB official scrape failed: {e}")
    return ""


def _scrape_boj_official() -> str:
    """爬 BOJ 政策利率（FRED fallback）— 從 trading economics meta description 抓最新值"""
    try:
        import re as _re
        resp = requests.get(
            "https://tradingeconomics.com/japan/interest-rate",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        # 主：<meta name="description"> 內 "interest rate in Japan was last recorded at X.XX percent"
        m = _re.search(
            r'<meta[^>]*name="description"[^>]*content="[^"]*?(\d+\.\d+)\s*percent',
            resp.text, _re.IGNORECASE)
        if m:
            return f"{float(m.group(1)):.2f}"
    except Exception as e:
        logger.warning(f"BOJ official scrape failed: {e}")
    return ""


def _scrape_boe_official() -> str:
    """爬 BoE 官網首頁 Bank Rate（主要政策利率）"""
    try:
        import re as _re
        resp = requests.get(
            "https://www.bankofengland.co.uk/",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        # 首頁顯示 "Bank Rate X.XX%"（不同 HTML 結構，用較寬鬆的 pattern）
        m = _re.search(r'Bank Rate[^0-9%]{1,80}(\d+\.\d+)\s*%', resp.text, _re.DOTALL)
        if m:
            return f"{float(m.group(1)):.2f}"
    except Exception as e:
        logger.warning(f"BoE official scrape failed: {e}")
    return ""


def _scrape_boe_tradingeconomics() -> str:
    """BoE 備援：tradingeconomics meta description"""
    try:
        import re as _re
        resp = requests.get(
            "https://tradingeconomics.com/united-kingdom/interest-rate",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        m = _re.search(
            r'<meta[^>]*name="description"[^>]*content="[^"]*?(\d+\.\d+)\s*percent',
            resp.text, _re.IGNORECASE)
        if m:
            return f"{float(m.group(1)):.2f}"
    except Exception as e:
        logger.warning(f"BoE TE scrape failed: {e}")
    return ""


def _scrape_pboc_omo() -> str:
    """爬 PBOC 7 天逆回購利率（央行政策利率，對標 Fed Funds / ECB DFR / BoE Bank Rate）
    來源：tradingeconomics（中國央行政策利率）"""
    try:
        import re as _re
        resp = requests.get(
            "https://tradingeconomics.com/china/interest-rate",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        m = _re.search(
            r'<meta[^>]*name="description"[^>]*content="[^"]*?(\d+\.\d+)\s*percent',
            resp.text, _re.IGNORECASE)
        if m:
            return f"{float(m.group(1)):.2f}"
    except Exception as e:
        logger.warning(f"PBOC OMO scrape failed: {e}")
    return ""


def _next_pboc_lpr_date() -> str:
    """PBOC LPR 每月 20 日公布（遇假日順延），用此當作下次發布日"""
    today = date.today()
    if today.day < 20:
        target = today.replace(day=20)
    else:
        if today.month == 12:
            target = date(today.year + 1, 1, 20)
        else:
            target = date(today.year, today.month + 1, 20)
    return target.strftime("%Y/%m/%d")


# ============================================================
# 備援抓取：TradingView → FRED（ForexFactory 未回填實際值時）
# ============================================================

def _fetch_tv_actuals(date_from: str, date_to: str) -> dict:
    """
    TradingView 經濟日曆 API（主備援）
    回傳：{ normalized_title: actual_str }
    端點：https://economic-calendar.tradingview.com/events
    """
    try:
        resp = requests.get(
            "https://economic-calendar.tradingview.com/events",
            headers={"User-Agent": "Mozilla/5.0", "Origin": "https://www.tradingview.com"},
            params={"from": date_from, "to": date_to,
                    "countries": "US,CN,EU,JP,TW"},
            timeout=10,
        )
        if resp.status_code != 200:
            return {}
        events = resp.json()
        if not isinstance(events, list):
            events = events.get("result", events.get("events", []))
        result = {}
        for e in events:
            actual = e.get("actual")
            if actual is not None:
                title_key = e.get("title", "").lower().strip()
                result[title_key] = str(actual)
        return result
    except Exception as ex:
        logger.warning(f"TradingView actuals failed: {ex}")
        return {}


def _match_tv_actual(ff_title: str, tv_actuals: dict) -> str:
    """
    把 ForexFactory 的事件標題對應到 TradingView 的實際值
    ⚠️ 關鍵：MoM / YoY / QoQ 期間必須嚴格匹配，不可跨期間互相汙染
    （例：FF "PPI m/m" 不可匹配 TV "PPI YoY"）
    """
    import re

    def _normalize(t: str):
        """回傳 (base, period) 兩部分。period ∈ {mom, yoy, qoq, none}"""
        s = t.lower().strip()
        # 拿掉修訂詞（不影響期間辨識）
        s = re.sub(r'\s*(final|preliminary|prelim|flash|revised|advance)\s*$', '', s).strip()
        # 偵測期間後綴並拿掉
        m = re.search(r'\s*(m/m|mom|q/q|qoq|y/y|yoy)\s*$', s)
        if m:
            tag = m.group(1).replace('/', '')
            period = {'mm': 'mom', 'qq': 'qoq', 'yy': 'yoy'}.get(tag, tag)
            base = s[:m.start()].strip()
        else:
            period, base = 'none', s
        return base, period

    ff_base, ff_period = _normalize(ff_title)

    # 1. 直接比對（期間必須相同）
    for tv_title, val in tv_actuals.items():
        tv_base, tv_period = _normalize(tv_title)
        if ff_base == tv_base and ff_period == tv_period:
            return val

    # 2. 關鍵詞包含比對（同期間、取最長匹配）
    best_match, best_len = "", 0
    for tv_title, val in tv_actuals.items():
        tv_base, tv_period = _normalize(tv_title)
        if ff_period != tv_period:   # 期間不同絕對不匹配
            continue
        if ff_base and (ff_base in tv_base or tv_base in ff_base):
            if len(tv_base) > best_len:
                best_match, best_len = val, len(tv_base)
    return best_match


# FRED 備援（TradingView 找不到時的第二道防線）
FRED_BACKUP_MAP = [
    # title 關鍵字（lower）          FRED series ID          格式
    ("core pce",                   "PCEPILFE",             "mom_pct"),
    ("pce price index",            "PCEPI",                "mom_pct"),
    ("core cpi",                   "CPILFESL",             "mom_pct"),
    ("cpi m/m",                    "CPIAUCSL",             "mom_pct"),
    ("cpi y/y",                    "CPIAUCSL",             "yoy_pct"),
    ("final gdp",                  "A191RL1Q225SBEA",      "level_1dp_pct"),
    ("gdp q/q",                    "A191RL1Q225SBEA",      "level_1dp_pct"),
    ("nonfarm payrolls",           "PAYEMS",               "mom_abs_k"),
    ("non-farm payrolls",          "PAYEMS",               "mom_abs_k"),
    ("unemployment rate",          "UNRATE",               "level_1dp_pct"),
]


def _fetch_actual_from_fred(title: str) -> str:
    """FRED 備援（第二層）：針對標準宏觀指標"""
    title_l = title.lower()
    for keyword, series_id, fmt in FRED_BACKUP_MAP:
        if keyword in title_l:
            try:
                n = 14 if fmt == "yoy_pct" else 3
                rows = _get_fred_csv(series_id, n_rows=n)
                if not rows:
                    return ""
                v = rows[-1][1]
                if fmt == "level_1dp":
                    return f"{v:.1f}"
                elif fmt == "level_1dp_pct":
                    return f"{v:.1f}%"
                elif fmt == "mom_pct":
                    if len(rows) >= 2:
                        pct = (v - rows[-2][1]) / rows[-2][1] * 100 if rows[-2][1] else 0
                        return f"{pct:+.1f}%"
                elif fmt == "yoy_pct":
                    if len(rows) >= 13:
                        pct = (v - rows[-13][1]) / rows[-13][1] * 100 if rows[-13][1] else 0
                        return f"{pct:+.1f}%"
                elif fmt == "mom_abs_k":
                    if len(rows) >= 2:
                        return f"{(v - rows[-2][1]) * 1000:+,.0f}K"
            except Exception as e:
                logger.warning(f"FRED backup {series_id} failed: {e}")
            break
    return ""


def _next_meeting(schedule: list) -> str:
    today = date.today()
    for d in schedule:
        if d >= today:
            return d.strftime("%Y/%m/%d")
    return schedule[-1].strftime("%Y/%m/%d") if schedule else "N/A"


# ============================================================
# 解釋文字生成
# ============================================================

def _fx_commentary(fx_list: list) -> str:
    """從 TWD 視角解釋匯率變動：X/TWD ↑ = X 對 TWD 升值（TWD 貶值）；DXY ↑ = 美元走強。"""
    # 各幣對的中文名稱與漲/跌時的解讀（從 TWD 視角）
    _META = {
        "美元 (USD/TWD)":      ("美元",     "美元對台幣升值（外資匯出壓力增、進口成本上升）", "美元對台幣貶值（外資匯入支撐、進口成本下降）"),
        "人民幣 (CNY/TWD)":    ("人民幣",   "人民幣對台幣升值（中國經濟信心回升或政策引導）", "人民幣對台幣貶值（中國資本外流壓力或政策寬鬆）"),
        "歐元 (EUR/TWD)":      ("歐元",     "歐元對台幣升值（歐洲經濟優於預期或 ECB 鷹派）",   "歐元對台幣貶值（歐洲經濟疲弱或美元避險需求）"),
        "英鎊 (GBP/TWD)":      ("英鎊",     "英鎊對台幣升值（英國經濟韌性或 BoE 偏鷹）",       "英鎊對台幣貶值（英國經濟下行壓力）"),
        "瑞士法郎 (CHF/TWD)":  ("瑞郎",     "瑞郎對台幣升值（避險資金流入、SNB 偏鷹）",        "瑞郎對台幣貶值（全球風險偏好回升、避險消退）"),
        "日圓 (JPY/TWD)":      ("日圓",     "日圓對台幣升值（避險資金湧入或日銀調整政策）",     "日圓對台幣貶值（日本出口導向但進口通膨壓力）"),
        "韓元 (KRW/TWD)":      ("韓元",     "韓元對台幣升值（外資回流韓股）",                   "韓元對台幣貶值（外資流出壓力）"),
        "澳幣 (AUD/TWD)":      ("澳幣",     "澳幣對台幣升值（大宗商品需求回升或中國經濟改善）", "澳幣對台幣貶值（商品價格走弱或全球風險趨避）"),
    }
    parts = []

    for item in fx_list:
        name = item["name"]
        c    = item.get("change_pct")

        # ── DXY 總覽（永遠第一）──
        if "DXY" in name:
            if c is None: continue
            if   c >  0.5: parts.append(f"DXY 美元指數上漲 {c:+.2f}%，美元全面走強，非美貨幣承壓")
            elif c >  0.1: parts.append(f"DXY 小幅走強 {c:+.2f}%，美元偏強格局")
            elif c < -0.5: parts.append(f"DXY 下跌 {c:+.2f}%，美元走弱，非美貨幣反彈")
            elif c < -0.1: parts.append(f"DXY 微跌 {c:+.2f}%，美元稍弱")
            else:           parts.append("DXY 持平，匯市觀望")
            continue

        meta = _META.get(name)
        if not meta or c is None or abs(c) < 0.2:
            continue
        cname, up_reason, dn_reason = meta
        if c > 0:
            parts.append(f"{name} ↑ {abs(c):.2f}%：{up_reason}")
        else:
            parts.append(f"{name} ↓ {abs(c):.2f}%：{dn_reason}")

    return "<br>".join(p + "。" for p in parts) if parts else "匯率整體變動不大，市場觀望氣氛濃厚。"


def _yield_commentary(yields: dict) -> str:
    """根據殖利率數據生成解釋"""
    y2  = yields.get("2Y",  {}).get("yield")
    y10 = yields.get("10Y", {}).get("yield")
    spread = yields.get("spread_10y_2y")

    parts = []

    # 利差解讀
    if spread is not None:
        if spread < 0:
            parts.append(f"殖利率曲線倒掛（10Y-2Y 利差 {spread:+.0f}bps），市場隱含衰退預期")
        elif spread < 20:
            parts.append(f"殖利率曲線接近平坦（利差 {spread:.0f}bps），經濟前景不確定性高")
        elif spread < 80:
            parts.append(f"殖利率曲線正常化（利差 {spread:.0f}bps），市場預期經濟溫和成長")
        else:
            parts.append(f"殖利率曲線陡峭化（利差 {spread:.0f}bps），市場定價未來通膨或降息預期")

    # 10Y 絕對水位
    if y10 is not None:
        if y10 > 5.0:
            parts.append(f"10Y 殖利率高於 5%（{y10:.2f}%），長天期融資成本壓力大，股市估值承壓")
        elif y10 > 4.5:
            parts.append(f"10Y 殖利率處於 {y10:.2f}% 高位，成長股估值面臨壓力")
        elif y10 < 3.5:
            parts.append(f"10Y 殖利率降至 {y10:.2f}%，反映市場避險或降息預期")

    # 短端變動
    chg_2y = yields.get("2Y", {}).get("change_bps")
    if chg_2y is not None and abs(chg_2y) >= 5:
        direction = "上升" if chg_2y > 0 else "下降"
        parts.append(f"2Y 利率單日{direction} {abs(chg_2y):.0f}bps，反映市場對 Fed 政策預期調整")

    return "<br>".join(p + "。" for p in parts) if parts else "殖利率變動不大，市場等待新催化劑。"


def _sentiment_commentary(vix_price: float = None, vix_chg: float = None,
                           pc_data: dict = None) -> str:
    """根據 VIX + Put/Call Ratio 生成市場情緒解釋"""
    parts = []

    # ── VIX 解讀 ──
    if vix_price is not None:
        if vix_price > 40:
            parts.append(f"VIX 飆至 {vix_price:.1f}，已進入歷史前 5% 高位區間，"
                         f"隱含 S&P 500 未來 30 日年化波動率約 {vix_price:.0f}%，"
                         f"選擇權市場定價劇烈震盪")
        elif vix_price > 30:
            parts.append(f"VIX {vix_price:.1f}，高於長期均值（19-20）超過 50%，"
                         f"避險需求顯著升溫，市場預期未來一個月波動加劇")
        elif vix_price > 25:
            parts.append(f"VIX {vix_price:.1f}，高於長期均值，"
                         f"市場不確定性上升，避險成本增加")
        elif vix_price > 20:
            parts.append(f"VIX {vix_price:.1f}，略高於長期均值，市場存在一定不安情緒")
        elif vix_price > 15:
            parts.append(f"VIX {vix_price:.1f}，處於正常區間（長期均值 19-20），市場波動溫和")
        else:
            parts.append(f"VIX 僅 {vix_price:.1f}，處於歷史低位，市場極度自滿，"
                         f"低波動環境下需警惕突發事件衝擊")

        if vix_chg is not None and abs(vix_chg) >= 5:
            if vix_chg > 15:
                parts.append(f"VIX 單日暴漲 {vix_chg:+.1f}%，選擇權避險需求急遽升溫")
            elif vix_chg > 5:
                parts.append(f"VIX 單日上升 {vix_chg:+.1f}%，避險情緒明顯加重")
            elif vix_chg < -10:
                parts.append(f"VIX 單日大跌 {vix_chg:+.1f}%，恐慌快速消退")
            elif vix_chg < -5:
                parts.append(f"VIX 單日回落 {vix_chg:+.1f}%，市場緊張情緒緩解")

    # ── Put/Call Ratio 解讀 ──
    if pc_data and pc_data.get("current") is not None:
        pc = pc_data["current"]
        puts_per_100c = int(round(pc * 100))
        if pc >= 1.20:
            parts.append(f"Put/Call Ratio {pc:.2f}（每 100 口 call 對應 {puts_per_100c} 口 put），"
                         f"put 量遠高於 call，市場極度避險；歷史上此區間後常見技術性反彈")
        elif pc >= 1.00:
            parts.append(f"Put/Call Ratio {pc:.2f}（每 100 口 call 對應 {puts_per_100c} 口 put），"
                         f"put 主導，看空情緒偏濃")
        elif pc >= 0.80:
            parts.append(f"Put/Call Ratio {pc:.2f}（每 100 口 call 對應 {puts_per_100c} 口 put），"
                         f"略偏空但接近中性")
        elif pc >= 0.65:
            parts.append(f"Put/Call Ratio {pc:.2f}（每 100 口 call 對應 {puts_per_100c} 口 put），"
                         f"略偏多")
        elif pc >= 0.50:
            parts.append(f"Put/Call Ratio {pc:.2f}（每 100 口 call 對應 {puts_per_100c} 口 put），"
                         f"call 主導，看多情緒明確")
        else:
            parts.append(f"Put/Call Ratio {pc:.2f}（每 100 口 call 對應 {puts_per_100c} 口 put），"
                         f"call 量遠高於 put，市場極度貪婪；歷史上此區間後常見短線拉回")

        chg = pc_data.get("change")
        if chg is not None and abs(chg) >= 0.10:
            direction = "急升" if chg > 0 else "驟降"
            parts.append(f"P/C 比率較前日{direction} {chg:+.2f}，情緒轉折劇烈")

    return "<br>".join(p + "。" for p in parts) if parts else "市場情緒數據暫無法取得。"


def _tech_commentary(tech: dict) -> str:
    """根據美股技術面數據生成 SPX/NDQ 文字總結"""
    parts = []
    for key in ('spx', 'ndq'):
        t = tech.get(key, {})
        if not t.get('ok'):
            continue
        name = t['name']
        price, m20, m60, m200 = t['price'], t['ma20'], t['ma60'], t['ma200']
        rsi, pct_high = t['rsi'], t['pct_high']

        # 趨勢
        if price > m20 > m60 > m200:
            trend = f"<b>{name}</b> 多頭排列穩固（現價站上月＞季＞年線），趨勢健康"
        elif price > m60 and price > m200 and price > m20:
            trend = f"<b>{name}</b> 多頭趨勢（站上月、季、年線）"
        elif price > m60 and price > m200:
            trend = f"<b>{name}</b> 多頭趨勢（站上季線與年線，月線拉回中）"
        elif price < m60 and price < m200:
            trend = f"<b>{name}</b> 空頭趨勢（跌破季線與年線），動能轉弱"
        else:
            trend = f"<b>{name}</b> 進入盤整（均線糾結）"

        # 距 52W 高
        if pct_high >= -3:
            high_lbl = f"距 52W 高 {pct_high:+.1f}%（強勢區，逼近高點）"
        elif pct_high >= -10:
            high_lbl = f"距 52W 高 {pct_high:+.1f}%（強勢區）"
        elif pct_high >= -20:
            high_lbl = f"距 52W 高 {pct_high:+.1f}%（回檔區）"
        else:
            high_lbl = f"距 52W 高 {pct_high:+.1f}%（深度回檔）"

        # RSI
        if rsi >= 70:
            rsi_lbl = f"RSI {rsi} <b>超買</b>（短線需留意回檔）"
        elif rsi >= 60:
            rsi_lbl = f"RSI {rsi}（偏多動能強）"
        elif rsi >= 50:
            rsi_lbl = f"RSI {rsi}（中性偏多）"
        elif rsi >= 40:
            rsi_lbl = f"RSI {rsi}（中性偏弱）"
        elif rsi >= 30:
            rsi_lbl = f"RSI {rsi}（偏弱）"
        else:
            rsi_lbl = f"RSI {rsi} <b>超賣</b>（技術面可能反彈）"

        parts.append(f"{trend}；{high_lbl}；{rsi_lbl}")

    return "<br>".join(p + "。" for p in parts) if parts else ""


def _commodity_commentary(cm) -> str:
    """根據原物料+加密數據生成詳細解釋，敘述順序與顯示排序一致（按漲跌幅降冪）"""
    # 統一轉為排序後的 list
    if isinstance(cm, dict):
        cm_list = sorted(cm.values(), key=lambda x: x.get("change_pct") or -999, reverse=True)
    else:
        cm_list = list(cm)

    # 建 dict 供多品項比較邏輯使用
    cm_dict = {c["name"]: c for c in cm_list}
    parts = []

    for item in cm_list:
        name = item["name"]
        c = item.get("change_pct")
        p = item.get("price", 0) or 0

        # ── WTI 原油 ──
        if name == "WTI 原油":
            if c is None: continue
            if   c >  3: parts.append(f"WTI 原油大漲 {c:+.2f}% 至 ${p:.2f}，可能受地緣衝突升級（中東局勢/制裁）或 OPEC+ 減產等供給面因素推動")
            elif c >  1: parts.append(f"WTI 上漲 {c:+.2f}% 至 ${p:.2f}，供需面偏緊")
            elif c < -3: parts.append(f"WTI 重挫 {c:+.2f}% 至 ${p:.2f}，可能反映全球需求放緩預期、美國庫存意外增加或 OPEC+ 增產訊號")
            elif c < -1: parts.append(f"WTI 下跌 {c:+.2f}% 至 ${p:.2f}，需求面偏弱")
            else:        parts.append(f"WTI 原油 ${p:.2f}（{c:+.2f}%），價格窄幅震盪")
            if   p > 90: parts.append(f"油價處於 ${p:.0f} 高位，推升運輸與製造成本，通膨壓力可能延遲 Fed 降息時程")
            elif p < 60: parts.append(f"油價跌至 ${p:.0f} 低位，有利消費者但打擊能源類股獲利")

        # ── 布蘭特原油（主要顯示 Brent-WTI 價差）──
        elif name == "布蘭特原油":
            wti_p = cm_dict.get("WTI 原油", {}).get("price", 0) or 0
            if p and wti_p:
                spread = p - wti_p
                if spread > 8:
                    parts.append(f"Brent-WTI 價差擴至 ${spread:.1f}，反映國際市場供給緊於美國")

        # ── 天然氣 ──
        elif name == "天然氣":
            if c is not None and abs(c) > 1.5:
                direction = "上漲" if c > 0 else "下跌"
                reason = "冬季取暖需求或供給中斷" if c > 0 else "暖冬或庫存充足"
                parts.append(f"天然氣{direction} {abs(c):.2f}% 至 ${p:.3f}，{reason}，影響電力與工業成本")

        # ── 黃金 ──
        elif name == "黃金":
            if c is not None:
                if   c >  1.5: parts.append(f"黃金大漲 {c:+.2f}% 至 ${p:,.0f}，避險需求與央行購金雙重推動，實質利率下行預期支撐金價")
                elif c >  0.3: parts.append(f"黃金上漲 {c:+.2f}% 至 ${p:,.0f}，避險買盤持續")
                elif c < -1.5: parts.append(f"黃金下跌 {c:+.2f}%，美元走強或實質利率上升壓抑金價")
                elif c < -0.3: parts.append(f"黃金微跌 {c:+.2f}%，獲利了結賣壓")
            if p > 4000:
                parts.append(f"金價 ${p:,.0f} 處於歷史高位，反映去美元化趨勢與全球央行儲備多元化")

        # ── 白銀 ──
        elif name == "白銀":
            if c is not None and abs(c) > 1.5:
                direction = "上漲" if c > 0 else "下跌"
                parts.append(f"白銀{direction} {abs(c):.2f}%，兼具工業（光伏、電子）與貴金屬屬性，{'跟隨金價走強' if c > 0 else '工業需求疑慮拖累'}")

        # ── 銅 ──
        elif name == "銅":
            if c is not None:
                if abs(c) > 1.5:
                    direction = "上漲" if c > 0 else "下跌"
                    signal = "全球製造業 PMI 改善或中國基建需求回升" if c > 0 else "工業需求放緩，中國經濟復甦不如預期"
                    parts.append(f"銅價{direction} {abs(c):.2f}% 至 ${p:.3f}/磅，「銅博士」暗示{signal}")
                elif abs(c) > 0.5:
                    direction = "小漲" if c > 0 else "小跌"
                    parts.append(f"銅價{direction} {abs(c):.2f}%，反映工業活動溫和波動")

        # ── BTC ──
        elif name == "BTC":
            if c is None: continue
            if   c >  5: parts.append(f"BTC 大漲 {c:+.2f}% 至 ${p:,.0f}，機構資金流入或監管利多，風險偏好顯著回升")
            elif c >  2: parts.append(f"BTC 上漲 {c:+.2f}%，加密市場情緒偏多")
            elif c < -5: parts.append(f"BTC 重挫 {c:+.2f}%，流動性收緊或大戶拋售，風險資產全面承壓")
            elif c < -2: parts.append(f"BTC 下跌 {c:+.2f}%，加密市場風險偏好下降")
            else:        parts.append(f"BTC ${p:,.0f}（{c:+.2f}%），窄幅整理")

        # ── ETH（與 BTC 相對表現）──
        elif name == "ETH":
            btc_c = cm_dict.get("BTC", {}).get("change_pct")
            if c is not None and btc_c is not None:
                if   c - btc_c >  3: parts.append("ETH 明顯跑贏 BTC，山寨幣輪動行情啟動")
                elif btc_c - c >  3: parts.append("BTC 獨強、ETH 落後，資金集中流向比特幣避險")

    return "<br>".join(p + "。" for p in parts) if parts else "原物料整體變動不大，市場觀望。"


def _taiwan_commentary(tw: dict) -> str:
    """根據三大法人數據生成解釋"""
    if not tw:
        return ""
    parts = []
    f_yi = tw.get("foreign_yi", 0)
    i_yi = tw.get("inv_trust_yi", 0)
    d_yi = tw.get("dealer_yi", 0)
    t_yi = tw.get("total_yi", 0)

    # 合計方向
    if t_yi > 50:
        parts.append(f"三大法人合計買超 {t_yi:.1f} 億元，資金淨流入台股，多方氣勢強")
    elif t_yi > 0:
        parts.append(f"三大法人合計小幅買超 {t_yi:.1f} 億元")
    elif t_yi > -50:
        parts.append(f"三大法人合計小幅賣超 {abs(t_yi):.1f} 億元")
    elif t_yi > -200:
        parts.append(f"三大法人合計賣超 {abs(t_yi):.1f} 億元，資金持續流出")
    else:
        parts.append(f"三大法人大幅賣超 {abs(t_yi):.1f} 億元，市場承受沉重賣壓")

    # 外資動向（通常是主力）
    if abs(f_yi) > 100:
        direction = "買超" if f_yi > 0 else "賣超"
        parts.append(f"外資單日{direction} {abs(f_yi):.1f} 億元，為今日主要{'推升' if f_yi > 0 else '壓抑'}力量")
    elif abs(f_yi) > 30:
        direction = "買超" if f_yi > 0 else "賣超"
        parts.append(f"外資{direction} {abs(f_yi):.1f} 億元")

    # 投信動向（反映內資法人看法）
    if abs(i_yi) > 20:
        direction = "買超" if i_yi > 0 else "賣超"
        signal = "內資法人持續佈局" if i_yi > 0 else "內資法人減碼"
        parts.append(f"投信{direction} {abs(i_yi):.1f} 億元，{signal}")

    # 外資 vs 投信 背離
    if f_yi < -50 and i_yi > 10:
        parts.append("外資賣、投信買，內外資出現背離，關注後續誰主導行情")
    elif f_yi > 50 and i_yi < -10:
        parts.append("外資買、投信賣，法人態度分歧")

    return "<br>".join(p + "。" for p in parts) if parts else ""


def _index_commentary(indices: list) -> str:
    """根據全球股市漲跌生成解釋"""
    non_vix = [i for i in indices if not i.get("is_vix")]
    if not non_vix:
        return ""
    parts = []

    valid = [i for i in non_vix if i.get("change_pct") is not None]
    if not valid:
        return ""

    best  = max(valid, key=lambda x: x["change_pct"])
    worst = min(valid, key=lambda x: x["change_pct"])

    # 整體氛圍
    up_count   = sum(1 for i in valid if i["change_pct"] > 0)
    down_count = sum(1 for i in valid if i["change_pct"] < 0)
    total      = len(valid)

    if up_count >= total * 0.8:
        parts.append(f"全球股市普漲，{up_count}/{total} 大市場上揚，風險偏好回升")
    elif down_count >= total * 0.8:
        parts.append(f"全球股市齊跌，{down_count}/{total} 大市場收跌，避險情緒主導")
    elif up_count > down_count:
        parts.append(f"股市多頭略占優勢（{up_count} 漲 / {down_count} 跌）")
    else:
        parts.append(f"股市偏弱（{down_count} 跌 / {up_count} 漲）")

    # 最強 / 最弱
    if best["change_pct"] > 0.5:
        parts.append(f"領漲：{best['name']} {best['change_pct']:+.2f}%")
    if worst["change_pct"] < -0.5:
        parts.append(f"領跌：{worst['name']} {worst['change_pct']:+.2f}%")

    # 台股
    taiex = next((i for i in valid if "TAIEX" in i["name"]), None)
    if taiex and taiex.get("change_pct") is not None:
        c = taiex["change_pct"]
        if c < -2:
            parts.append(f"台股重挫 {c:.2f}%，需留意外資動向與半導體類股壓力")
        elif c < -1:
            parts.append(f"台股下跌 {c:.2f}%，表現弱於全球均值")
        elif c > 2:
            parts.append(f"台股強漲 {c:.2f}%，表現明顯優於全球")
        elif c > 0.5:
            parts.append(f"台股上漲 {c:.2f}%，跟隨全球多頭")

    # US vs Asia 背離
    sp = next((i for i in valid if "S&P 500" in i["name"]), None)
    nk = next((i for i in valid if "Nikkei" in i["name"] or "日經" in i["name"]), None)
    if sp and nk and sp.get("change_pct") is not None and nk.get("change_pct") is not None:
        diff = sp["change_pct"] - nk["change_pct"]
        if diff > 2:
            parts.append("美股明顯強於亞股，資金偏向美國市場")
        elif diff < -2:
            parts.append("亞股相對強勢，美股承壓")

    return "<br>".join(p + "。" for p in parts) if parts else ""


def _calendar_commentary(cal: list) -> str:
    """根據本週經濟日曆生成總結，已公布事件逐項解讀"""
    if not cal:
        return ""
    parts = []

    published = [e for e in cal if e.get("published")]
    pending   = [e for e in cal if not e.get("published") and not e.get("is_past")]
    past_unconfirmed = [e for e in cal if e.get("is_past")]

    # ── 已公布：逐項解讀 ──
    for e in published:
        title = e.get("title", "")
        actual = e.get("actual", "")
        forecast = e.get("forecast", "—")
        bi = e.get("beat_indicator", "")
        title_l = title.lower()

        try:
            act_num = float(actual.replace("%","").replace("K","").replace(",","").strip())
            fc_num  = float(forecast.replace("%","").replace("K","").replace(",","").strip()) if forecast != "—" else None
        except:
            act_num = None; fc_num = None

        # 反向指標（數字越低越好）：▼=優於預期，▲=遜於預期
        # 同時含中英文關鍵字（標題經 _translate_cal_title 後可能已轉中文）
        INVERSE_INDICATORS = (
            # 英文（FF/TV 原始）
            "jobless claims", "initial claims", "continuing claims",
            "unemployment rate",
            "cpi", "inflation rate", "core inflation",
            "pce", "ppi",
            # 中文（_translate_cal_title 翻譯後）
            "失業金", "失業救濟", "失業率",
            "通膨", "核心 cpi",
        )
        is_inverse = any(k in title_l for k in INVERSE_INDICATORS)
        if is_inverse:
            beat_word = "優於預期" if bi == "▼" else ("遜於預期" if bi == "▲" else "符合預期")
        else:
            beat_word = "優於預期" if bi == "▲" else ("遜於預期" if bi == "▼" else "符合預期")
        unit = "%" if "%" in actual else ""

        if "ism services" in title_l or "non manufacturing pmi" in title_l:
            trend = "服務業景氣擴張（>50）" if act_num and act_num > 50 else "服務業景氣收縮（<50）"
            parts.append(f"ISM 服務業 PMI {actual}（{beat_word}，預期 {forecast}），{trend}，{'需注意新訂單與就業分項走向' if bi == '▼' else '支撐美國服務業佔比 80% 的 GDP'}")
        elif "cpi" in title_l or "inflation rate" in title_l or "core inflation" in title_l:
            parts.append(f"{title} 實際 {actual}（{beat_word}，預期 {forecast}），{'通膨高於預期，Fed 降息壓力增大' if bi == '▲' else ('通膨降溫，市場降息預期升溫' if bi == '▼' else '通膨符合 Fed 預測路徑')}")
        elif "pce" in title_l:
            parts.append(f"{title} {actual}（{beat_word}，預期 {forecast}），{'PCE 為 Fed 核心通膨指標，高於預期將延後降息時程' if bi == '▲' else 'PCE 回落支撐 Fed 降息預期'}")
        elif "gdp" in title_l:
            parts.append(f"GDP {actual}（{beat_word}，預期 {forecast}），{'經濟成長動能優於市場預期' if bi == '▲' else '經濟成長放緩，衰退疑慮升溫'}")
        elif "durable goods" in title_l:
            parts.append(f"耐久財訂單 {actual}（{beat_word}，預期 {forecast}），{'製造業資本支出需求增強' if bi == '▲' else '企業資本投資趨保守'}")
        elif "jobless claims" in title_l or "initial claims" in title_l:
            parts.append(f"首次申請失業救濟金 {actual}（{beat_word}，預期 {forecast}），{'申請人數增加，勞動市場出現鬆弛' if bi == '▲' else '勞動市場仍緊俏，支撐消費'}")
        elif "personal income" in title_l:
            parts.append(f"個人收入 {actual}（{beat_word}，預期 {forecast}），{'薪資成長支撐消費動能' if bi == '▲' else '薪資增速放緩，消費潛力承壓'}")
        elif "personal spending" in title_l:
            parts.append(f"個人支出 {actual}（{beat_word}，預期 {forecast}），{'消費擴張是美國 GDP 最大支柱' if bi == '▲' else '消費降溫，需觀察信心指標'}")
        elif "ppi" in title_l:
            parts.append(f"PPI {actual}（{beat_word}，預期 {forecast}），{'生產者物價上漲，通膨上游壓力升溫' if bi == '▲' else '上游通膨壓力緩解，有利終端物價穩定'}")
        elif "consumer confidence" in title_l or "michigan" in title_l:
            parts.append(f"{title} {actual}（{beat_word}，預期 {forecast}），{'消費信心回升，支撐零售與服務業需求' if bi == '▲' else '消費信心走弱，留意未來消費支出'}")
        elif "retail sales" in title_l:
            parts.append(f"零售銷售 {actual}（{beat_word}，預期 {forecast}），{'消費端仍有活力' if bi == '▲' else '消費降溫，需關注信心指標'}")
        elif beat_word != "符合預期":
            parts.append(f"{title} {actual}（{beat_word}，預期 {forecast}）")

    # ── 未確認（時間已過但無資料）──
    if past_unconfirmed:
        items = "、".join(e["title"] for e in past_unconfirmed[:2])
        parts.append(f"⚠️ 待確認：{items}（時間已過，尚未取得實際值）")

    # ── 本週待公布重點 ──
    KEY_KEYWORDS = ["cpi","pce","gdp","payroll","nonfarm","jobless","pmi","ppi","consumer confidence","michigan","retail"]
    key_pending = [e for e in pending if any(k in e.get("title","").lower() for k in KEY_KEYWORDS)]
    if key_pending:
        upcoming = "、".join(e["title"] for e in key_pending[:4])
        parts.append(f"本週待公布重點：{upcoming}")

    return "<br>".join(p + "。" for p in parts) if parts else ""


# ============================================================
# 各區塊抓取函式
# ============================================================

def fetch_fx_data() -> dict:
    """抓匯率，全部從 TWD 視角呈現（X/TWD），DXY 保留為錨指標。

    精確算法：
      X/TWD_today = USD/TWD_today / (USD/X)_today        (usd_base)
      X/TWD_today = (X/USD)_today × USD/TWD_today        (quote_base)
    再用 today / prev_close 各自值算 change_pct，避免 log-linear 近似誤差。
    """
    # Step 1：先取 USD/TWD spot + prev（其他換算都要用）
    twd_q = _get_quote("TWD=X")
    usdtwd_now  = twd_q["price"]
    usdtwd_prev = twd_q["prev_close"]

    results = []
    for name, ticker, dec, mode in FX_TICKERS:
        if mode == "dxy":
            q = _get_quote(ticker)
            price   = q["price"]
            chg_pct = q["change_pct"]

        elif mode == "usd_twd":
            price   = usdtwd_now
            chg_pct = twd_q["change_pct"]

        else:
            q = _get_quote(ticker)
            now, prev = q["price"], q["prev_close"]
            if now is None or prev is None or usdtwd_now is None or usdtwd_prev is None:
                price, chg_pct = None, None
            else:
                if mode == "usd_base":     # ticker 是 USD/X
                    x_twd_now  = usdtwd_now  / now
                    x_twd_prev = usdtwd_prev / prev
                else:                      # quote_base：ticker 是 X/USD
                    x_twd_now  = now  * usdtwd_now
                    x_twd_prev = prev * usdtwd_prev
                price   = round(x_twd_now, 6)
                chg_pct = round((x_twd_now - x_twd_prev) / x_twd_prev * 100, 2) if x_twd_prev else None

        results.append({
            "name": name,
            "price": price,
            "price_fmt": _fmt(price, dec) if price is not None else "N/A",
            "change_pct": chg_pct,
        })

    return {
        "items": results,
        "commentary": _fx_commentary(results),
    }


def fetch_yield_curve() -> dict:
    """從 FRED 抓殖利率 + 生成解釋"""
    result = {}
    for label, fred_id in YIELD_FRED_IDS:
        rows = _get_fred_csv(fred_id, n_rows=5)
        if len(rows) >= 2:
            curr = rows[-1][1]
            prev = rows[-2][1]
            chg_bps = round((curr - prev) * 100, 1)
            result[label] = {"yield": round(curr, 3), "change_bps": chg_bps, "date": rows[-1][0]}
        elif len(rows) == 1:
            result[label] = {"yield": round(rows[0][1], 3), "change_bps": 0, "date": rows[0][0]}
        else:
            result[label] = {"yield": None, "change_bps": None, "date": None}

    y10 = result.get("10Y", {}).get("yield")
    y2  = result.get("2Y",  {}).get("yield")
    result["spread_10y_2y"] = round((y10 - y2) * 100, 0) if y10 and y2 else None
    result["commentary"] = _yield_commentary(result)
    return result


def fetch_index_data() -> list:
    """抓指數，按漲跌幅排序（漲→跌），VIX 獨立置底"""
    results = []
    for name, ticker, flag in INDEX_TICKERS:
        q = _get_quote(ticker)
        results.append({
            "name": name, "flag": flag,
            "price": q["price"],
            "price_fmt": _fmt(q["price"]),
            "change_pct": q["change_pct"],
        })

    results.sort(key=lambda x: x["change_pct"] if x["change_pct"] is not None else -999, reverse=True)

    vix_q = _get_quote(VIX_TICKER)
    results.append({
        "name": "VIX", "flag": "📉",
        "price": vix_q["price"],
        "price_fmt": _fmt(vix_q["price"]),
        "change_pct": vix_q["change_pct"],
        "is_vix": True,
    })
    return results


def fetch_commodity_data() -> dict:
    """抓原物料，按漲跌幅排序 + 解釋"""
    results = []
    for name, ticker, dec, sym in COMMODITY_TICKERS:
        q = _get_quote(ticker)
        results.append({
            "name": name, "symbol": sym,
            "price": q["price"],
            "price_fmt": _fmt(q["price"], dec),
            "change_pct": q["change_pct"],
        })

    results.sort(key=lambda x: x["change_pct"] if x["change_pct"] is not None else -999, reverse=True)

    return {
        "items": results,
        "commentary": _commodity_commentary(results),
    }


def fetch_crypto_data() -> list:
    """從 Binance 抓加密貨幣（免費、無需 Key、24h 即時變動）
    端點：GET https://api.binance.com/api/v3/ticker/24hr?symbols=["BTCUSDT","ETHUSDT","SOLUSDT"]
    """
    BINANCE_SYMBOLS = [
        ("BTC", "BTCUSDT", 0),
        ("ETH", "ETHUSDT", 2),
        ("SOL", "SOLUSDT", 2),
    ]
    try:
        import json as _json
        syms = _json.dumps([s[1] for s in BINANCE_SYMBOLS], separators=(',', ':'))
        url  = f"https://api.binance.com/api/v3/ticker/24hr?symbols={requests.utils.quote(syms)}"
        resp = requests.get(url, headers={"Accept": "application/json"}, timeout=10)
        data = {item["symbol"]: item for item in resp.json()}

        results = []
        for name, sym, dec in BINANCE_SYMBOLS:
            item       = data.get(sym, {})
            price      = float(item["lastPrice"])      if item.get("lastPrice")      else None
            change_pct = float(item["priceChangePercent"]) if item.get("priceChangePercent") else None
            vol        = float(item.get("volume", 0)) * price if price else 0
            results.append({
                "name": name, "symbol": "$",
                "price":      round(price, dec)      if price      else None,
                "price_fmt":  _fmt(price, dec)       if price      else "N/A",
                "change_pct": round(change_pct, 2)   if change_pct else None,
                "volume_24h": vol,
                "source":     "binance",
            })
        return results

    except Exception as e:
        logger.warning(f"Binance failed: {e}, falling back to yfinance")
        fallback_tickers = [("BTC", "BTC-USD", 0), ("ETH", "ETH-USD", 2), ("SOL", "SOL-USD", 2)]
        results = []
        for name, ticker, dec in fallback_tickers:
            q = _get_quote(ticker)
            results.append({
                "name": name, "symbol": "$",
                "price":      q["price"],
                "price_fmt":  _fmt(q["price"], dec),
                "change_pct": q["change_pct"],
                "source":     "yfinance_fallback",
            })
        return results


def fetch_geopolitics() -> dict:
    """讀取 geopolitics.json（由 Claude WebSearch 生成）"""
    try:
        if os.path.exists(GEOPOLITICS_FILE):
            with open(GEOPOLITICS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        logger.warning(f"Geopolitics file failed: {e}")
    return None


def fetch_cb_rates() -> dict:
    """央行利率：四家每日自動抓，每家三層 fallback（主來源 → 備援爬蟲 → 上次 cache → hardcoded）"""
    result = {}

    def _resolve(name: str, primary_fn, secondary_fn, hardcoded: str) -> str:
        """三層 fallback：primary → secondary → 上次 cache → hardcoded"""
        # Tier 1: 主來源（FRED 或官網）
        try:
            v = primary_fn()
            if v:
                logger.info(f"{name} rate from primary: {v}")
                return v
        except Exception as e:
            logger.warning(f"{name} primary failed: {e}")
        # Tier 2: 備援來源
        if secondary_fn:
            try:
                v = secondary_fn()
                if v:
                    logger.info(f"{name} rate from secondary: {v}")
                    return v
            except Exception as e:
                logger.warning(f"{name} secondary failed: {e}")
        # Tier 3: 上次 cache（last-known-good）
        cached, cached_ts = _read_last_cb_rate(name)
        if cached:
            logger.warning(f"{name} using last-known-good from cache ({cached_ts}): {cached}")
            return cached
        # Tier 4: hardcoded final fallback
        logger.warning(f"{name} all sources failed, using hardcoded: {hardcoded}")
        return hardcoded

    # ── Fed ──
    def _fed_fred():
        l = _get_fred_csv("DFEDTARL", 3)
        u = _get_fred_csv("DFEDTARU", 3)
        return f"{l[-1][1]:.2f}–{u[-1][1]:.2f}" if (l and u) else ""

    fed_rate = _resolve("聯準會 (Fed)", _fed_fred, _scrape_fed_official, "3.50–3.75")

    # ── ECB ──
    def _ecb_fred():
        r = _get_fred_csv("ECBDFR", 3)
        return f"{r[-1][1]:.2f}" if r else ""

    ecb_rate = _resolve("歐洲央行 (ECB)", _ecb_fred, _scrape_ecb_official, "2.00")

    # ── BOJ ──
    def _boj_fred():
        r = _get_fred_csv("IRSTJPN156N", 3)
        return f"{r[-1][1]:.2f}" if r else ""

    boj_rate = _resolve("日本央行 (BOJ)", _boj_fred, _scrape_boj_official, "0.50")

    # ── BoE（英國央行 Bank Rate）──
    boe_rate = _resolve("英國央行 (BoE)", _scrape_boe_official, _scrape_boe_tradingeconomics, "4.25")

    # ── PBOC（中國央行 7 天逆回購）──
    pboc_rate = _resolve("中國央行 (PBOC)", _scrape_pboc_omo, None, "1.40")

    # ── CBC（台灣央行重貼現率）──
    def _cbc_official():
        import re as _re
        resp = requests.get(
            'https://www.cbc.gov.tw/tw/cp-534-4088-F0CAF-2.html',
            headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
        m = _re.search(r'重貼現率.*?<em>([\d.]+)%</em>', resp.text, _re.DOTALL)
        return f"{float(m.group(1)):.2f}" if m else ""

    cbc_rate = _resolve("中央銀行 (CBC)", _cbc_official, None, "2.00")

    # 顯示順序：Fed → ECB → BoE → BOJ → PBOC → CBC（依重要性與地理）
    result["聯準會 (Fed)"]    = {"rate": fed_rate,  "next": _next_meeting(FOMC_DATES_2026)}
    result["歐洲央行 (ECB)"]  = {"rate": ecb_rate,  "next": _next_meeting(ECB_DATES_2026)}
    result["英國央行 (BoE)"]  = {"rate": boe_rate,  "next": _next_meeting(BOE_DATES_2026)}
    result["日本央行 (BOJ)"]  = {"rate": boj_rate,  "next": _next_meeting(BOJ_DATES_2026)}
    result["中國央行 (PBOC)"] = {"rate": pboc_rate, "next": _next_pboc_lpr_date() + "（LPR）"}
    result["中央銀行 (CBC)"]  = {"rate": cbc_rate,  "next": _next_meeting(CBC_DATES_2026)}

    return result


def fetch_spx_technical() -> dict:
    """美股技術面（針對月度投資週期）：S&P 500 + Nasdaq
    5 個指標：MA20（月線）/ MA60（季線）/ MA200（年線）/ RSI(21) / 距 52W 高
    統計學原理：觀察窗 ≈ 投資週期 時訊號雜訊比最佳；月度持倉用 21 天 RSI 比 14 天乾淨"""
    def _rsi(closes, period=21):
        deltas = [closes[i] - closes[i-1] for i in range(1, len(closes))]
        gains  = [max(d, 0)      for d in deltas[-period:]]
        losses = [abs(min(d, 0)) for d in deltas[-period:]]
        avg_g = sum(gains)  / period
        avg_l = sum(losses) / period
        if avg_l == 0:
            return 100.0
        return round(100 - 100 / (1 + avg_g / avg_l), 1)

    def _rsi_lbl(v):
        if v < 30: return '超賣'
        if v < 40: return '偏弱'
        if v < 50: return '中性偏弱'
        if v < 60: return '中性偏多'
        if v < 70: return '偏多'
        return '超買'

    def _trend_lbl(price, ma20, ma60, ma200):
        """月度視角的趨勢判讀（站上幾條均線）"""
        if price > ma20 > ma60 > ma200:
            return ('多頭排列', True)        # 完美多頭：站上月＞季＞年
        if price > ma200 and price > ma60:
            return ('多頭趨勢', True)        # 站上季線與年線
        if price < ma200 and price < ma60:
            return ('空頭趨勢', False)       # 跌破季線與年線
        return ('盤整中', None)

    def _high_lbl(pct_from_high):
        """距 52W 高的月度安全邊際判讀"""
        if pct_from_high >= -3:   return '創新高 / 強勢區'
        if pct_from_high >= -10:  return '強勢區'
        if pct_from_high >= -20:  return '修正中'
        if pct_from_high >= -30:  return '熊市邊緣'
        return '深度熊市'

    import urllib.request as _urlr, urllib.parse as _urlp

    def _hist(symbol):
        url = f'https://query1.finance.yahoo.com/v8/finance/chart/{_urlp.quote(symbol)}?interval=1d&range=1y'
        req = _urlr.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with _urlr.urlopen(req, timeout=12) as r:
            d = json.loads(r.read())
        closes = d['chart']['result'][0]['indicators']['quote'][0]['close']
        return [x for x in closes if x is not None]

    result = {}
    for key, sym, name in [('spx', '^GSPC', '標普 500 (S&P 500)'), ('ndq', '^IXIC', '那斯達克 (Nasdaq)')]:
        try:
            closes = _hist(sym)
            if len(closes) < 200:
                result[key] = {'ok': False, 'name': name}
                continue
            price   = closes[-1]
            ma20    = sum(closes[-20:])  / 20
            ma60    = sum(closes[-60:])  / 60
            ma200   = sum(closes[-200:]) / 200
            rsi21   = _rsi(closes, 21)
            window52 = closes[-252:] if len(closes) >= 252 else closes
            high52w = max(window52)
            low52w  = min(window52)
            pct_high = (price - high52w) / high52w * 100
            # 價格軸定位圖：現價在 52W 區間的相對位置 (0~100%)
            range_pos = (price - low52w) / (high52w - low52w) * 100 if high52w > low52w else 50
            trend, trend_ok = _trend_lbl(price, ma20, ma60, ma200)
            result[key] = {
                'ok':       True,
                'name':     name,
                'price':    round(price, 2),
                'ma20':     round(ma20,  2),
                'ma60':     round(ma60,  2),
                'ma200':    round(ma200, 2),
                'pct20':    round((price - ma20)  / ma20  * 100, 1),
                'pct60':    round((price - ma60)  / ma60  * 100, 1),
                'pct200':   round((price - ma200) / ma200 * 100, 1),
                'rsi':      rsi21,
                'rsi_lbl':  _rsi_lbl(rsi21),
                'high52w':  round(high52w, 2),
                'low52w':   round(low52w,  2),
                'pct_high': round(pct_high, 1),
                'high_lbl': _high_lbl(pct_high),
                'range_pos': round(range_pos, 1),  # 現價在 52W 區間的位置 0~100
                'trend':    trend,
                'trend_ok': trend_ok,
            }
        except Exception as e:
            logger.warning(f"fetch_spx_technical {sym}: {e}")
            result[key] = {'ok': False, 'name': name}
    return result


def fetch_fear_greed() -> dict:
    """Alternative.me Fear & Greed（取代 CNN，無反爬限制，Render 可用）"""
    try:
        resp = requests.get(
            "https://api.alternative.me/fng/?limit=2",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=8
        )
        data = resp.json()["data"]
        score  = int(data[0]["value"])
        rating = data[0]["value_classification"]
        prev   = int(data[1]["value"])
        return {"score": score, "rating": rating, "change": score - prev}
    except Exception as e:
        logger.warning(f"Alternative.me F&G failed: {e}")

    # Fallback: VIX 推算
    try:
        vix_q = _get_quote(VIX_TICKER)
        v = vix_q.get("price")
        if v:
            if v < 12:   score, label = 85, "Extreme Greed"
            elif v < 15: score, label = 70, "Greed"
            elif v < 20: score, label = 55, "Neutral"
            elif v < 25: score, label = 38, "Fear"
            elif v < 30: score, label = 25, "Fear"
            else:        score, label = 12, "Extreme Fear"
            return {"score": score, "rating": f"{label} (VIX={v:.1f})", "change": None, "source": "vix_proxy"}
    except Exception:
        pass

    return {"score": None, "rating": "N/A", "change": None}


# ── CBOE Put/Call Ratio（從 ycharts 抓當前值，自建歷史累積）─────────────
PC_RATIO_HISTORY = os.path.join(os.path.dirname(__file__), "pc_ratio_history.json")

def fetch_put_call_ratio() -> dict:
    """CBOE Equity Put/Call Ratio：scrape ycharts 取當前值 + 自建 30 天歷史

    回傳：{
        "current": float,           # 最新值
        "change": float,            # vs 前日
        "history": [{"date","value"},...],  # 最近 30 個交易日
        "rating": str,              # 多空判讀
        "rating_color": str,        # #hex
    }
    高於 1.0 = 偏空（put 多於 call），低於 0.7 = 偏多（call 多於 put）
    """
    current = None
    # ── 主源：ycharts ──
    try:
        url = 'https://ycharts.com/indicators/cboe_equity_put_call_ratio'
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
        })
        html = urllib.request.urlopen(req, timeout=10).read().decode(errors='replace')
        m = re.search(r'"last_value"[:\s]*"?([0-9]\.[0-9]+)', html) or \
            re.search(r'([0-9]\.[0-9]+)\s*for\s*\w+\s+\d+\s+\d{4}', html)
        if m:
            current = float(m.group(1))
    except Exception as e:
        logger.warning(f"ycharts P/C failed: {e}")

    # ── 載入歷史 ──
    history = []
    try:
        if os.path.exists(PC_RATIO_HISTORY):
            with open(PC_RATIO_HISTORY, "r") as f:
                history = json.load(f).get("history", [])
    except Exception:
        history = []

    # ── append 今日值（同日只存最新一次）──
    if current is not None:
        today = datetime.now(TAIPEI_TZ).strftime("%Y-%m-%d")
        history = [h for h in history if h.get("date") != today]
        history.append({"date": today, "value": current})
        # 只保留最近 60 天，前端只用 30 天
        history = history[-60:]
        try:
            with open(PC_RATIO_HISTORY, "w") as f:
                json.dump({"history": history}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"P/C history save failed: {e}")

    # ── 變化（vs 前一個交易日）──
    change = None
    if len(history) >= 2:
        prev = history[-2]["value"]
        if current is not None:
            change = round(current - prev, 3)

    # ── 多空判讀 ──
    if current is None:
        rating, color = "N/A", "#6b7280"
    elif current >= 1.20:
        rating, color = "極度恐慌（put 飆升）", "#ef4444"
    elif current >= 1.00:
        rating, color = "偏空", "#f97316"
    elif current >= 0.80:
        rating, color = "中性偏空", "#eab308"
    elif current >= 0.65:
        rating, color = "中性偏多", "#84cc16"
    elif current >= 0.50:
        rating, color = "偏多", "#22c55e"
    else:
        rating, color = "極度貪婪（call 主導）", "#16a34a"

    return {
        "current": current,
        "change": change,
        "history": history[-30:],   # 給前端畫 sparkline
        "rating": rating,
        "rating_color": color,
    }


# ── 日曆事件：標題中文對照表 ──────────────────────────────────────────────────
_CAL_TITLE_ZH = {
    # 就業
    "Non Farm Payrolls":                   "非農就業人數",
    "Nonfarm Payrolls":                    "非農就業人數",
    "Unemployment Rate":                   "失業率",
    "Initial Jobless Claims":              "初領失業金人數",
    "Continuing Jobless Claims":           "持續申請失業救濟",
    "ADP Nonfarm Employment Change":       "ADP 民間就業",
    "Average Hourly Earnings MoM":         "平均時薪月增",
    # 通膨（TradingView 用 Inflation Rate；FF 用 CPI）
    "Inflation Rate YoY":                  "CPI 通膨年增率",
    "Inflation Rate MoM":                  "CPI 通膨月增率",
    "Core Inflation Rate YoY":             "核心 CPI 年增率",
    "Core Inflation Rate MoM":             "核心 CPI 月增率",
    "CPI MoM":                             "CPI 月增率",
    "CPI YoY":                             "CPI 年增率",
    "Core CPI MoM":                        "核心 CPI 月增率",
    "Core CPI YoY":                        "核心 CPI 年增率",
    "PPI MoM":                             "PPI 月增率",
    "PPI YoY":                             "PPI 年增率",
    "Core PPI MoM":                        "核心 PPI 月增率",
    "Core PPI YoY":                        "核心 PPI 年增率",
    "PCE Price Index MoM":                 "PCE 物價月增率",
    "PCE Price Index YoY":                 "PCE 物價年增率",
    "Core PCE Price Index MoM":            "核心 PCE 月增率",
    "Core PCE Price Index YoY":            "核心 PCE 年增率",
    # GDP / 生產
    "GDP Growth Rate QoQ":                 "GDP 季增率",
    "GDP Growth Rate QoQ 2nd Est":         "GDP 季增率(二估)",
    "GDP Growth Rate QoQ Final":           "GDP 季增率(終值)",
    "GDP Growth Rate YoY":                 "GDP 年增率",
    "GDP Growth Rate":                     "GDP 成長率",
    # 消費 / 零售
    "Retail Sales MoM":                    "零售銷售月增率",
    "Core Retail Sales MoM":               "核心零售銷售月增率",
    "Retail Sales YoY":                    "零售銷售年增率",
    "Consumer Confidence":                 "消費者信心指數",
    "Michigan Consumer Sentiment":         "密西根消費者信心",
    "Michigan Consumer Sentiment Final":   "密西根消費者信心(終值)",
    "Michigan Consumer Sentiment Prel":    "密西根消費者信心(初值)",
    # PMI / 景氣
    "ISM Manufacturing PMI":               "ISM 製造業 PMI",
    "ISM Non-Manufacturing PMI":           "ISM 服務業 PMI",
    "ISM Services PMI":                    "ISM 服務業 PMI",
    "S&P Global Manufacturing PMI":        "S&P 製造業 PMI",
    "S&P Global Services PMI":             "S&P 服務業 PMI",
    "Manufacturing PMI":                   "製造業 PMI",
    "Services PMI":                        "服務業 PMI",
    # Fed
    "Federal Funds Rate":                  "聯邦基金利率",
    "FOMC Meeting Minutes":                "FOMC 會議紀要",
    "Fed Interest Rate Decision":          "Fed 利率決策",
    # 台灣
    "Export Orders YoY":                   "外銷訂單年增率",
    "Exports YoY":                         "出口年增率",
    "Imports YoY":                         "進口年增率",
    "Balance of Trade":                    "貿易差額",
    "Industrial Production YoY":           "工業生產年增率",
    "M2 Money Supply YoY":                 "M2 貨幣供給年增率",
}

# ── 白名單：只放這些事件，其他過濾掉 ──────────────────────────────────
_US_IMPORTANT_TITLES = {
    # 通膨
    "Inflation Rate YoY", "Inflation Rate MoM",
    "Core Inflation Rate YoY", "Core Inflation Rate MoM",
    "CPI YoY", "CPI MoM", "Core CPI YoY", "Core CPI MoM",
    "PPI YoY", "PPI MoM", "Core PPI MoM", "Core PPI YoY",
    "PCE Price Index YoY", "PCE Price Index MoM",
    "Core PCE Price Index YoY", "Core PCE Price Index MoM",
    # 就業
    "Non Farm Payrolls", "Nonfarm Payrolls",
    "Unemployment Rate", "ADP Nonfarm Employment Change",
    "Initial Jobless Claims",
    "Average Hourly Earnings MoM",
    # GDP
    "GDP Growth Rate QoQ", "GDP Growth Rate QoQ 2nd Est", "GDP Growth Rate QoQ Final",
    # 景氣
    "ISM Manufacturing PMI", "ISM Non-Manufacturing PMI", "ISM Services PMI",
    # Fed
    "Fed Interest Rate Decision", "Federal Funds Rate", "FOMC Meeting Minutes",
    # 消費
    "Retail Sales MoM", "Core Retail Sales MoM",
    "Consumer Confidence",
    "Michigan Consumer Sentiment", "Michigan Consumer Sentiment Final",
    "Michigan Consumer Sentiment Prel",
}

# 台灣排除利率類；保留 GDP/通膨/貿易/生產/就業/PMI
_TW_IMPORTANT_TITLES = {
    # GDP
    "GDP Growth Rate YoY", "GDP Growth Rate",
    # 通膨
    "Inflation Rate YoY", "CPI YoY",
    # 貿易
    "Export Orders YoY", "Exports YoY", "Imports YoY", "Balance of Trade",
    # 生產 / 就業
    "Industrial Production YoY", "Unemployment Rate",
    # PMI
    "Manufacturing PMI", "S&P Global Manufacturing PMI",
    # 貨幣供給
    "M2 Money Supply YoY",
}


def _is_important_event(country: str, title: str) -> bool:
    """白名單過濾：只保留高關注度事件"""
    title = (title or "").strip()
    if country == "USD":
        return title in _US_IMPORTANT_TITLES
    if country == "TWD":
        # 台灣排除利率（央行利率決策、貼放款利率等）
        if "interest rate" in title.lower() or "rate decision" in title.lower():
            return False
        return title in _TW_IMPORTANT_TITLES
    return False


def _translate_cal_title(title: str) -> str:
    """回傳中文標題；找不到對照則回傳英文"""
    zh = _CAL_TITLE_ZH.get(title)
    if zh:
        return zh
    # 模糊匹配（title 包含 key 或 key 包含 title）
    for k, v in _CAL_TITLE_ZH.items():
        if k.lower() in title.lower() or title.lower() in k.lower():
            return v
    return title

def _fmt_cal_value(v: str) -> str:
    """格式化日曆數值：>1M 換算為 K/M/B/T，保留原有 % 符號"""
    if not v or v in ("—", "N/A", ""):
        return v
    v = v.strip()
    # 已有單位後綴（%、K、M、B、T）→ 直接回傳
    if v and v[-1].upper() in ('K', 'M', 'B', 'T', '%'):
        return v
    # 嘗試解析純數字（支援負號、小數、逗號千分位）
    try:
        num = float(v.replace(',', ''))
        abs_num = abs(num)
        sign = "-" if num < 0 else ""
        if   abs_num >= 1_000_000_000_000: return f"{sign}{abs_num/1_000_000_000_000:.2f}T"
        elif abs_num >= 1_000_000_000:     return f"{sign}{abs_num/1_000_000_000:.2f}B"
        elif abs_num >= 1_000_000:         return f"{sign}{abs_num/1_000_000:.2f}M"
        elif abs_num >= 10_000:            return f"{sign}{abs_num/1_000:.1f}K"
        else:                              return v
    except:
        return v


def _parse_calendar_events(events_raw: list, tv_actuals: dict = None) -> list:
    """
    解析 ForexFactory 日曆事件
    規則：
      - 僅保留 High impact、目標國家、且有 forecast 數值的事件（移除演講/聲明類）
      - 時間已過且無 actual → is_past=True，前端顯示 ⚠️ 待確認
      - 時間未到 → is_past=False，前端顯示 ⏳
      - 已有 actual → published=True，顯示 ✅ + beat indicator
    """
    TARGET_COUNTRIES = {"USD", "TWD"}
    now_utc = datetime.now(pytz.utc)

    # ── 本週視窗（Taipei 週一 00:00 ~ 下週一 00:00）──
    _now_taipei = now_utc.astimezone(TAIPEI_TZ)
    _week_start_taipei = (_now_taipei - timedelta(days=_now_taipei.weekday())).replace(
        hour=0, minute=0, second=0, microsecond=0)
    _week_end_taipei = _week_start_taipei + timedelta(days=7)
    _week_start_utc = _week_start_taipei.astimezone(pytz.utc)
    _week_end_utc = _week_end_taipei.astimezone(pytz.utc)

    def _in_this_week(ev):
        try:
            dt = datetime.fromisoformat(ev["date"].replace("Z", "+00:00"))
            return _week_start_utc <= dt < _week_end_utc
        except:
            return False

    high = [e for e in events_raw
            if e.get("impact") in ("High", "Medium")
            and e.get("country") in TARGET_COUNTRIES
            and (e.get("forecast") or "").strip()
            and _is_important_event(e.get("country", ""), e.get("title", ""))   # 白名單篩選
            and _in_this_week(e)]                                                # 僅本週（Taipei 週一→週日）

    def _sort_key(e):
        try:
            return datetime.fromisoformat(e["date"].replace("Z", "+00:00"))
        except:
            return datetime.min.replace(tzinfo=pytz.utc)
    high.sort(key=_sort_key)

    results = []
    for e in high:
        if len(results) >= 15:
            break

        dt_str = e.get("date", "")
        is_past = False
        try:
            dt_utc = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
            dt_taipei = dt_utc.astimezone(TAIPEI_TZ)
            date_fmt = dt_taipei.strftime("%m/%d %H:%M")
            is_past = dt_utc < now_utc
        except:
            date_fmt = dt_str[:10]

        actual   = (e.get("actual")   or "").strip()
        # ForexFactory 未回填時：① TradingView → ② FRED
        if is_past and not actual:
            title = e.get("title", "")
            if tv_actuals:
                actual = _match_tv_actual(title, tv_actuals)
            if not actual:
                actual = _fetch_actual_from_fred(title)

        forecast = (e.get("forecast") or "").strip() or "—"
        previous = (e.get("previous") or "").strip() or "—"

        beat_indicator = ""
        if actual:
            try:
                def _to_num(s):
                    return float(s.replace("%","").replace("K","000").replace("M","000000").strip())
                if forecast != "—":
                    beat_indicator = "▲" if _to_num(actual) > _to_num(forecast) else \
                                     ("▼" if _to_num(actual) < _to_num(forecast) else "")
            except:
                pass

        results.append({
            "date":           date_fmt,
            "country":        e.get("country", ""),
            "title":          _translate_cal_title(e.get("title", "")),
            "forecast":       _fmt_cal_value(forecast),
            "previous":       _fmt_cal_value(previous),
            "actual":         _fmt_cal_value(actual),
            "beat_indicator": beat_indicator,
            "published":      bool(actual),
            "is_past":        is_past and not bool(actual),
        })

    return results


def _fetch_tv_calendar_full(tv_from: str, tv_to: str) -> list:
    """
    從 TradingView 抓完整日曆事件清單（含 forecast + actual）
    用於 ForexFactory 失敗時的備援，回傳已格式化的事件 list
    """
    TARGET_COUNTRIES = {"USD", "TWD"}
    COUNTRY_MAP = {"US": "USD", "TW": "TWD"}
    try:
        resp = requests.get(
            "https://economic-calendar.tradingview.com/events",
            headers={"User-Agent": "Mozilla/5.0", "Origin": "https://www.tradingview.com"},
            params={"from": tv_from, "to": tv_to, "countries": "US,TW"},
            timeout=10,
        )
        if resp.status_code != 200:
            return []
        data = resp.json()
        raw_events = data.get("result", data) if isinstance(data, dict) else data
        if not isinstance(raw_events, list):
            return []

        # 轉換成 _parse_calendar_events 所需的 FF 格式
        now_utc = datetime.now(pytz.utc)
        pseudo_raw = []
        for e in raw_events:
            country_code = COUNTRY_MAP.get(e.get("country", "").upper(), "")
            if country_code not in TARGET_COUNTRIES:
                continue
            if e.get("importance", -1) < 0:
                continue
            forecast_raw = e.get("forecastRaw")
            if forecast_raw is None:
                continue  # 無預測數值
            pseudo_raw.append({
                "title":    e.get("title", ""),
                "country":  country_code,
                "date":     e.get("date", ""),
                "impact":   "High",
                "forecast": str(forecast_raw),
                "previous": str(e.get("previousRaw") or ""),
                "actual":   str(e.get("actualRaw")) if e.get("actualRaw") is not None else "",
            })

        # 直接呼叫解析（tv_actuals 已內嵌在 actual 欄位裡，不需再補）
        return _parse_calendar_events(pseudo_raw, tv_actuals={})
    except Exception as ex:
        logger.warning(f"TradingView full calendar failed: {ex}")
        return []


def fetch_economic_calendar() -> list:
    now_utc = datetime.now(pytz.utc)
    tv_from = (now_utc - timedelta(days=7)).strftime("%Y-%m-%dT00:00:00.000Z")
    tv_to   = (now_utc + timedelta(days=7)).strftime("%Y-%m-%dT23:59:59.000Z")

    # ── 主來源：TradingView（actual 直接內建，不需 title 配對）──
    tv_results = _fetch_tv_calendar_full(tv_from, tv_to)
    if tv_results:
        logger.info(f"TradingView primary returned {len(tv_results)} events")
        try:
            with open(CALENDAR_BACKUP, "w", encoding="utf-8") as f:
                json.dump(tv_results, f, ensure_ascii=False)
        except:
            pass
        return tv_results

    # ── 備援一：ForexFactory + 借 TV actuals 補值 ──
    logger.info("TradingView failed, fallback to ForexFactory")
    tv_actuals = _fetch_tv_actuals(tv_from, tv_to)
    for week in ["thisweek", "nextweek"]:
        try:
            url = f"https://nfs.faireconomy.media/ff_calendar_{week}.json"
            resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
            if resp.status_code != 200:
                continue
            events_raw = resp.json()
            if not events_raw:
                continue
            results = _parse_calendar_events(events_raw, tv_actuals=tv_actuals)
            if results:
                try:
                    with open(CALENDAR_BACKUP, "w", encoding="utf-8") as f:
                        json.dump(results, f, ensure_ascii=False)
                except:
                    pass
                return results
        except Exception as e:
            logger.warning(f"Calendar FF {week} failed: {e}")
            continue

    # ── 備援二：calendar_backup.json（同週才用）──
    try:
        if os.path.exists(CALENDAR_BACKUP):
            with open(CALENDAR_BACKUP, "r", encoding="utf-8") as f:
                cached = json.load(f)
            # 確認備份是本週資料（第一筆日期要在 5 天內）
            if cached:
                first_date = cached[0].get("date", "")
                try:
                    from datetime import datetime as _dt
                    cached_dt = _dt.strptime(first_date, "%m/%d %H:%M")
                    cached_dt = cached_dt.replace(year=now_utc.year)
                    if abs((now_utc.replace(tzinfo=None) - cached_dt).days) <= 7:
                        return cached
                except:
                    pass
    except:
        pass
    return []


def fetch_taiwan_market() -> dict:
    """
    三大法人買賣超（億元）
    策略：TWSE BFI82U 不帶日期 → 自動回最近交易日
    注意：TWSE 對海外 IP（如 Render）可能封鎖，fallback 使用 taiwan_backup.json
          backup 由本機 macro_dashboard_runner.py 每日更新後 push 至 GitHub
    """
    def parse_amt(s):
        try: return int(str(s).replace(",", ""))
        except: return 0

    def yi(val):
        return round(val / 1e8, 1)

    try:
        # 嘗試直接抓 TWSE（台灣 IP 可用，海外可能 timeout）
        url = "https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=json&type=day"
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=6)
        data = resp.json()
        if data.get("stat") != "OK" or not data.get("data"):
            raise ValueError("TWSE returned no data")

        rows = data["data"]
        foreign = inv_trust = dealer = total = 0
        for row in rows:
            name = row[0]
            net  = parse_amt(row[3])
            if name.startswith("外資及陸資"):
                foreign = net
            elif name == "投信":
                inv_trust = net
            elif "自營商" in name:
                dealer += net
            elif name == "合計":
                total = net

        result = {
            "foreign": foreign, "foreign_yi": yi(foreign),
            "inv_trust": inv_trust, "inv_trust_yi": yi(inv_trust),
            "dealer": dealer, "dealer_yi": yi(dealer),
            "total": total, "total_yi": yi(total),
            "date": data.get("date", ""),
            "unit": "億元",
            "source": "live",
        }
        # 成功就更新 backup
        try:
            with open(TAIWAN_BACKUP, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False)
        except:
            pass
        return result

    except Exception as e:
        logger.warning(f"Taiwan TWSE live failed ({e})，改用 backup")
        # fallback：用 github push 進來的 backup（由本機每日更新）
        try:
            if os.path.exists(TAIWAN_BACKUP):
                with open(TAIWAN_BACKUP, "r", encoding="utf-8") as f:
                    result = json.load(f)
                result["source"] = "backup"
                return result
        except:
            pass
        return None


# ============================================================
# fetch_all + cache
# ============================================================

def fetch_all() -> dict:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r") as f:
                cache = json.load(f)
            cache_time = datetime.fromisoformat(cache["timestamp"])
            if (datetime.now() - cache_time).total_seconds() < CACHE_TTL_HOURS * 3600:
                logger.info("Returning cached data")
                return cache["data"]
        except:
            pass

    logger.info("Fetching fresh data...")

    fx_data     = fetch_fx_data();         fx_ts    = _now_taipei()
    tech_data   = fetch_spx_technical();  tech_ts   = _now_taipei()
    cb_data     = fetch_cb_rates();        cb_ts    = _now_taipei()
    indices     = fetch_index_data();   indices_ts  = _now_taipei()
    fg          = fetch_fear_greed();   fg_ts       = _now_taipei()
    pc_ratio    = fetch_put_call_ratio()
    comm_data   = fetch_commodity_data(); comm_ts   = _now_taipei()
    crypto_data = fetch_crypto_data(); crypto_ts    = _now_taipei()
    tw_data     = fetch_taiwan_market(); tw_ts      = _now_taipei()
    cal_data    = fetch_economic_calendar(); cal_ts  = _now_taipei()

    # 提取 VIX 給情緒解釋用
    vix_item = next((i for i in indices if i.get("is_vix")), None)
    vix_price = vix_item["price"] if vix_item else None
    vix_chg   = vix_item["change_pct"] if vix_item else None

    # VIX 近 1 個月走勢（取代原 Put/Call 近 N 日走勢圖）
    vix_history = _get_history(VIX_TICKER, "1mo")

    # 合併原物料 + 加密貨幣，按漲跌幅排序後傳入 commentary
    all_comm = comm_data["items"] + crypto_data
    all_comm_sorted = sorted(all_comm, key=lambda x: x.get("change_pct") or -999, reverse=True)
    combined_commentary = _commodity_commentary(all_comm_sorted)

    # 台灣三大法人：用 TWSE 資料日期
    if tw_data and tw_data.get("date"):
        d = tw_data["date"]
        try:
            tw_ts = f"{d[:4]}/{d[4:6]}/{d[6:]}（TWSE）"
        except:
            pass

    data = {
        "fx":            fx_data["items"],
        "fx_commentary": fx_data["commentary"],
        "fx_updated_at": fx_ts,
        "spx_tech":          tech_data,
        "tech_commentary":   _tech_commentary(tech_data),
        "spx_tech_updated_at": tech_ts,
        "cb_rates":      cb_data,
        "cb_rates_updated_at": cb_ts,
        "indices":            indices,
        "indices_commentary": _index_commentary(indices),
        "indices_updated_at": indices_ts,
        "commodities":        comm_data["items"],
        "crypto":             crypto_data,
        "commodities_commentary": combined_commentary,
        "commodities_updated_at": comm_ts,
        "crypto_updated_at":  crypto_ts,
        "fear_greed":    fg,
        "put_call_ratio": pc_ratio,
        "vix_history":   vix_history,
        "sentiment_commentary": _sentiment_commentary(vix_price, vix_chg, pc_ratio),
        "fg_updated_at": fg_ts,
        "calendar":             cal_data,
        "calendar_commentary":  _calendar_commentary(cal_data),
        "calendar_updated_at":  cal_ts,
        "taiwan":             tw_data,
        "taiwan_commentary":  _taiwan_commentary(tw_data),
        "taiwan_updated_at":  tw_ts,
        "geopolitics":   fetch_geopolitics(),
        "updated_at":    _now_taipei(),
        "market_date":   _get_last_trading_day(),
    }

    try:
        with open(CACHE_FILE, "w") as f:
            json.dump({"timestamp": datetime.now().isoformat(), "data": data}, f, ensure_ascii=False)
    except:
        pass

    return data


def _now_taipei() -> str:
    """回傳當前台北時間字串，格式：YYYY/MM/DD HH:MM"""
    return datetime.now(TAIPEI_TZ).strftime("%Y/%m/%d %H:%M")


def _get_last_trading_day() -> str:
    now = datetime.now()
    wd = now.weekday()
    if wd == 5:
        last = now - timedelta(days=1)
    elif wd == 6:
        last = now - timedelta(days=2)
    else:
        last = now
    return last.strftime("%Y/%m/%d")


def clear_cache():
    if os.path.exists(CACHE_FILE):
        os.remove(CACHE_FILE)
