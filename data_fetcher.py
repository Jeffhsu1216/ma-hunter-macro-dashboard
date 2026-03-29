"""
總經儀表板 — 數據抓取模組
M&A Hunter Macro Dashboard Data Fetcher

數據來源：Yahoo Finance (yfinance), CNN Fear & Greed, ForexFactory, TWSE
"""

import yfinance as yf
import requests
from datetime import datetime, timedelta, date
import json, os, logging

logger = logging.getLogger(__name__)

# ============================================================
# 數據定義
# ============================================================

# (name, ticker, decimals)
FX_TICKERS = [
    ("DXY 美元指數", "DX-Y.NYB", 2),
    ("USD/TWD",     "TWD=X",    2),
    ("USD/JPY",     "JPY=X",    2),
    ("USD/CNY",     "CNY=X",    4),
    ("EUR/USD",     "EURUSD=X", 4),
    ("GBP/USD",     "GBPUSD=X", 4),
    ("AUD/USD",     "AUDUSD=X", 4),
    ("USD/KRW",     "KRW=X",    0),
]

YIELD_TICKERS = [
    ("2Y",  "^IRX"),
    ("5Y",  "^FVX"),
    ("10Y", "^TNX"),
    ("30Y", "^TYX"),
]

# CBC hardcoded — update when policy changes
_CB_STATIC = {
    "CBC": {"rate": "2.00", "next": "2026/06/19"},
}

INDEX_TICKERS = [
    ("S&P 500",   "^GSPC",     "🇺🇸"),
    ("Nasdaq",    "^IXIC",     "🇺🇸"),
    ("FTSE 100",  "^FTSE",     "🇬🇧"),
    ("DAX",       "^GDAXI",    "🇩🇪"),
    ("STOXX 600", "^STOXX",    "🇪🇺"),
    ("日經 225",  "^N225",     "🇯🇵"),
    ("KOSPI",     "^KS11",     "🇰🇷"),
    ("恆生指數",  "^HSI",      "🇭🇰"),
    ("上證綜合",  "000001.SS", "🇨🇳"),
    ("Nifty 50",  "^NSEI",     "🇮🇳"),
    ("加權指數",  "^TWII",     "🇹🇼"),
]

# (name, ticker, decimals, currency_symbol)
COMMODITY_TICKERS = [
    ("WTI 原油",   "CL=F",    2, "$"),
    ("布蘭特原油", "BZ=F",    2, "$"),
    ("天然氣",     "NG=F",    3, "$"),
    ("黃金",       "GC=F",    0, "$"),
    ("白銀",       "SI=F",    2, "$"),
    ("銅",         "HG=F",    3, "$"),
    ("BTC",        "BTC-USD", 0, "$"),
    ("ETH",        "ETH-USD", 2, "$"),
]

VIX_TICKER = "^VIX"
CACHE_FILE          = os.path.join(os.path.dirname(__file__), "cache_data.json")
CALENDAR_BACKUP     = os.path.join(os.path.dirname(__file__), "calendar_backup.json")
CACHE_TTL_HOURS = 1


# ============================================================
# 核心抓取
# ============================================================

def _get_quote(ticker_symbol: str) -> dict:
    """取得單一標的即時報價"""
    try:
        tk = yf.Ticker(ticker_symbol)
        info = tk.fast_info
        price = getattr(info, "last_price", None)
        prev_close = getattr(info, "previous_close", None)

        if price is None or prev_close is None:
            hist = tk.history(period="5d")
            if len(hist) >= 2:
                price = float(hist["Close"].iloc[-1])
                prev_close = float(hist["Close"].iloc[-2])
            elif len(hist) == 1:
                price = float(hist["Close"].iloc[-1])
                prev_close = price
            else:
                return {"price": None, "change_pct": None}

        change_pct = ((price - prev_close) / prev_close) * 100 if prev_close else 0
        return {"price": round(float(price), 6), "change_pct": round(change_pct, 2)}
    except Exception as e:
        logger.warning(f"Failed {ticker_symbol}: {e}")
        return {"price": None, "change_pct": None}


def _fmt(price, decimals=2):
    if price is None:
        return "N/A"
    if price >= 10000:
        return f"{price:,.0f}"
    return f"{price:,.{decimals}f}"


# ============================================================
# 各區塊抓取函式
# ============================================================

def fetch_fx_data() -> list:
    results = []
    for name, ticker, dec in FX_TICKERS:
        q = _get_quote(ticker)
        results.append({
            "name": name,
            "price": q["price"],
            "price_fmt": _fmt(q["price"], dec),
            "change_pct": q["change_pct"],
        })
    return results


def fetch_yield_curve() -> dict:
    result = {}
    for label, ticker in YIELD_TICKERS:
        q = _get_quote(ticker)
        p = q["price"]
        chg_pct = q["change_pct"]
        chg_bps = round(chg_pct * p / 100, 1) if p and chg_pct else None
        result[label] = {"yield": p, "change_bps": chg_bps}

    y10 = result.get("10Y", {}).get("yield")
    y2  = result.get("2Y",  {}).get("yield")
    result["spread_10y_2y"] = round((y10 - y2) * 100, 0) if y10 and y2 else None
    return result


def fetch_index_data() -> list:
    results = []
    for name, ticker, flag in INDEX_TICKERS:
        q = _get_quote(ticker)
        results.append({
            "name": name, "flag": flag,
            "price": q["price"],
            "price_fmt": _fmt(q["price"]),
            "change_pct": q["change_pct"],
        })
    vix = _get_quote(VIX_TICKER)
    results.append({
        "name": "VIX", "flag": "📉",
        "price": vix["price"],
        "price_fmt": _fmt(vix["price"]),
        "change_pct": vix["change_pct"],
        "is_vix": True,
    })
    return results


def fetch_commodity_data() -> list:
    results = []
    for name, ticker, dec, sym in COMMODITY_TICKERS:
        q = _get_quote(ticker)
        results.append({
            "name": name, "symbol": sym,
            "price": q["price"],
            "price_fmt": _fmt(q["price"], dec),
            "change_pct": q["change_pct"],
        })
    return results


def fetch_cb_rates() -> dict:
    """抓取央行利率：Fed 動態從 FRED，其餘 hardcoded"""
    result = {}
    # Fed：從 FRED 抓目標區間上下限
    try:
        lower = requests.get(
            "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DFEDTARL",
            timeout=8
        ).text.strip().split("\n")[-1].split(",")[1]
        upper = requests.get(
            "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DFEDTARU",
            timeout=8
        ).text.strip().split("\n")[-1].split(",")[1]
        fed_rate = f"{float(lower):.2f}–{float(upper):.2f}"
    except Exception as e:
        logger.warning(f"FRED Fed rate failed: {e}")
        fed_rate = "N/A"
    result["Fed"] = {"rate": fed_rate, "next": "2026/05/07"}
    result.update(_CB_STATIC)
    return result


def fetch_fear_greed() -> dict:
    """抓取 CNN Fear & Greed 指數（多重 headers 嘗試）"""
    headers_list = [
        {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
            "Referer": "https://edition.cnn.com/markets/fear-and-greed",
            "Accept": "application/json, text/plain, */*",
        },
        {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
    ]
    for hdrs in headers_list:
        try:
            resp = requests.get(
                "https://production.dataviz.cnn.io/index/fearandgreed/graphdata",
                headers=hdrs, timeout=8
            )
            if resp.status_code == 200:
                fg = resp.json()["fear_and_greed"]
                score = round(fg["score"])
                rating = fg["rating"].replace("_", " ").title()
                prev = fg.get("previous_close", fg["score"])
                return {"score": score, "rating": rating, "change": round(fg["score"] - prev, 1)}
        except Exception:
            continue

    # Fallback: 根據 VIX 推算情緒（VIX < 15 極貪婪, > 30 極恐懼）
    try:
        vix = _get_quote(VIX_TICKER)
        v = vix["price"]
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


def _parse_calendar_events(events_raw: list) -> list:
    """解析 ForexFactory JSON → 高影響力事件列表"""
    high = [e for e in events_raw
            if e.get("impact") == "High"
            and e.get("country") in ("USD", "CNY", "EUR", "JPY", "TWD")]
    results = []
    for e in high[:12]:
        dt_str = e.get("date", "")
        try:
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
            date_fmt = dt.strftime("%m/%d %H:%M")
        except:
            date_fmt = dt_str[:10]
        results.append({
            "date": date_fmt,
            "country": e.get("country", ""),
            "title": e.get("title", ""),
            "forecast": e.get("forecast") or "—",
            "previous": e.get("previous") or "—",
        })
    return results


def fetch_economic_calendar() -> list:
    """抓本週經濟日曆；失敗時用備援快取（週末常見）"""
    for week in ["thisweek", "nextweek"]:
        try:
            url = f"https://nfs.faireconomy.media/ff_calendar_{week}.json"
            resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
            if resp.status_code != 200:
                continue
            events_raw = resp.json()
            if not events_raw:
                continue
            results = _parse_calendar_events(events_raw)
            if results:
                # 成功 → 寫備援快取
                try:
                    with open(CALENDAR_BACKUP, "w", encoding="utf-8") as f:
                        json.dump(results, f, ensure_ascii=False)
                except:
                    pass
                return results
        except Exception as e:
            logger.warning(f"Calendar {week} failed: {e}")
            continue

    # 所有來源失敗（週末）→ 讀上次備援
    try:
        if os.path.exists(CALENDAR_BACKUP):
            with open(CALENDAR_BACKUP, "r", encoding="utf-8") as f:
                cached = json.load(f)
            logger.info("Calendar: using backup cache")
            return cached
    except:
        pass

    return []


def fetch_taiwan_market() -> dict:
    try:
        url = "https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=json&type=day"
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
        data = resp.json()
        if data.get("stat") != "OK" or not data.get("data"):
            return None

        rows = data["data"]

        def parse_amt(s):
            try:
                return int(str(s).replace(",", ""))
            except:
                return 0

        def yi(val):  # 轉億元，保留1位小數
            return round(val / 1e8, 1)

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

        return {
            "foreign":   foreign,   "foreign_yi":   yi(foreign),
            "inv_trust": inv_trust, "inv_trust_yi": yi(inv_trust),
            "dealer":    dealer,    "dealer_yi":    yi(dealer),
            "total":     total,     "total_yi":     yi(total),
            "date":      data.get("date", ""),
            "unit":      "億元",
        }
    except Exception as e:
        logger.warning(f"Taiwan market failed: {e}")
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
    data = {
        "fx":            fetch_fx_data(),
        "yields":        fetch_yield_curve(),
        "cb_rates":      fetch_cb_rates(),
        "indices":       fetch_index_data(),
        "commodities":   fetch_commodity_data(),
        "fear_greed":    fetch_fear_greed(),
        "calendar":      fetch_economic_calendar(),
        "taiwan":        fetch_taiwan_market(),
        "updated_at":    datetime.now().strftime("%Y/%m/%d %H:%M"),
        "market_date":   _get_last_trading_day(),
    }

    try:
        with open(CACHE_FILE, "w") as f:
            json.dump({"timestamp": datetime.now().isoformat(), "data": data}, f, ensure_ascii=False)
    except:
        pass

    return data


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
