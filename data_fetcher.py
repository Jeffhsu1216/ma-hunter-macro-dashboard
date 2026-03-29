"""
總經儀表板 — 數據抓取模組
M&A Hunter Macro Dashboard Data Fetcher

數據來源：Yahoo Finance (yfinance)
更新頻率：每日台北時間 08:00
"""

import yfinance as yf
from datetime import datetime, timedelta
import json
import os
import logging

logger = logging.getLogger(__name__)

# ============================================================
# 數據定義
# ============================================================

FX_TICKERS = {
    "USD/TWD": "TWD=X",
    "USD/JPY": "JPY=X",
    "USD/CNY": "CNY=X",
    "EUR/USD": "EURUSD=X",
    "USD/KRW": "KRW=X",
}

RATE_TICKERS = {
    "US 10Y": "^TNX",
    "US 2Y": "^IRX",  # 13-week T-Bill as proxy; will try 2Y separately
}

INDEX_TICKERS = {
    "S&P 500": {"ticker": "^GSPC", "flag": "\U0001f1fa\U0001f1f8"},
    "Nasdaq": {"ticker": "^IXIC", "flag": "\U0001f1fa\U0001f1f8"},
    "上證綜合": {"ticker": "000001.SS", "flag": "\U0001f1e8\U0001f1f3"},
    "STOXX 600": {"ticker": "^STOXX", "flag": "\U0001f1ea\U0001f1fa"},
    "日經 225": {"ticker": "^N225", "flag": "\U0001f1ef\U0001f1f5"},
    "KOSPI": {"ticker": "^KS11", "flag": "\U0001f1f0\U0001f1f7"},
    "加權指數": {"ticker": "^TWII", "flag": "\U0001f1f9\U0001f1fc"},
}

VIX_TICKER = "^VIX"

COMMODITY_TICKERS = {
    "WTI 原油": "CL=F",
    "布蘭特原油": "BZ=F",
    "黃金": "GC=F",
    "白銀": "SI=F",
    "BTC": "BTC-USD",
}

CACHE_FILE = os.path.join(os.path.dirname(__file__), "cache_data.json")
CACHE_TTL_HOURS = 1  # 快取有效期（小時）


def _get_quote(ticker_symbol: str) -> dict:
    """取得單一標的即時報價"""
    try:
        tk = yf.Ticker(ticker_symbol)
        info = tk.fast_info
        price = getattr(info, "last_price", None)
        prev_close = getattr(info, "previous_close", None)

        if price is None or prev_close is None:
            # fallback: 用 history
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
        return {"price": round(float(price), 4), "change_pct": round(change_pct, 2)}
    except Exception as e:
        logger.warning(f"Failed to fetch {ticker_symbol}: {e}")
        return {"price": None, "change_pct": None}


def _format_price(price, decimals=2):
    """格式化價格"""
    if price is None:
        return "N/A"
    if price >= 10000:
        return f"{price:,.0f}"
    elif price >= 100:
        return f"{price:,.{decimals}f}"
    else:
        return f"{price:.{decimals}f}"


def fetch_fx_data() -> list:
    """抓取匯率數據"""
    results = []
    for name, ticker in FX_TICKERS.items():
        data = _get_quote(ticker)
        # EUR/USD 需要反轉顯示邏輯（yfinance 回傳的是 EURUSD）
        if name == "EUR/USD" and data["price"]:
            pass  # EURUSD=X already gives EUR/USD

        decimals = 4 if name == "EUR/USD" else 2
        if name == "USD/KRW" and data["price"]:
            decimals = 0

        results.append({
            "name": name,
            "price": data["price"],
            "price_fmt": _format_price(data["price"], decimals),
            "change_pct": data["change_pct"],
        })
    return results


def fetch_rate_data() -> dict:
    """抓取利率數據"""
    # US 10Y
    tnx = _get_quote("^TNX")
    ten_y = tnx["price"]  # ^TNX is in percentage points
    ten_y_chg = tnx["change_pct"]

    # US 2Y — use ^IRX (13-week T-Bill) as proxy, or try direct history
    two_y_data = {"price": None, "change_pct": None}
    for two_y_ticker in ["^IRX"]:
        try:
            tk = yf.Ticker(two_y_ticker)
            hist = tk.history(period="5d")
            if len(hist) >= 2:
                two_y_val = float(hist["Close"].iloc[-1])
                two_y_prev = float(hist["Close"].iloc[-2])
                two_y_chg_val = ((two_y_val - two_y_prev) / two_y_prev * 100) if two_y_prev else 0
                two_y_data = {"price": round(two_y_val, 2), "change_pct": round(two_y_chg_val, 2)}
                break
            elif len(hist) == 1:
                two_y_data = {"price": round(float(hist["Close"].iloc[-1]), 2), "change_pct": 0}
                break
        except:
            continue

    two_y = two_y_data["price"]
    two_y_chg = two_y_data["change_pct"]

    # 利差
    spread = None
    if ten_y and two_y:
        spread = round((ten_y - two_y) * 100, 0)  # bps

    return {
        "fed_rate": "3.50 – 3.75",
        "ten_y": ten_y,
        "ten_y_chg_bps": round(ten_y_chg * ten_y / 100, 1) if ten_y and ten_y_chg else None,
        "two_y": two_y,
        "two_y_chg_bps": round(two_y_chg * two_y / 100, 1) if two_y and two_y_chg else None,
        "spread_bps": spread,
    }


def fetch_index_data() -> list:
    """抓取全球股市指數"""
    results = []
    for name, info in INDEX_TICKERS.items():
        data = _get_quote(info["ticker"])
        results.append({
            "name": name,
            "flag": info["flag"],
            "price": data["price"],
            "price_fmt": _format_price(data["price"]),
            "change_pct": data["change_pct"],
        })

    # VIX
    vix = _get_quote(VIX_TICKER)
    results.append({
        "name": "VIX",
        "flag": "\U0001f4c9",
        "price": vix["price"],
        "price_fmt": _format_price(vix["price"]),
        "change_pct": vix["change_pct"],
        "is_vix": True,
    })
    return results


def fetch_commodity_data() -> list:
    """抓取原物料與加密貨幣"""
    results = []
    for name, ticker in COMMODITY_TICKERS.items():
        data = _get_quote(ticker)
        decimals = 0 if name == "BTC" else 2
        results.append({
            "name": name,
            "price": data["price"],
            "price_fmt": _format_price(data["price"], decimals),
            "change_pct": data["change_pct"],
        })
    return results


def fetch_all() -> dict:
    """抓取所有數據，帶快取"""
    # 檢查快取
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
        "fx": fetch_fx_data(),
        "rates": fetch_rate_data(),
        "indices": fetch_index_data(),
        "commodities": fetch_commodity_data(),
        "updated_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
        "market_date": _get_last_trading_day(),
    }

    # 寫入快取
    try:
        with open(CACHE_FILE, "w") as f:
            json.dump({"timestamp": datetime.now().isoformat(), "data": data}, f, ensure_ascii=False)
    except:
        pass

    return data


def _get_last_trading_day() -> str:
    """取得最近交易日"""
    now = datetime.now()
    weekday = now.weekday()
    if weekday == 5:  # Saturday
        last = now - timedelta(days=1)
    elif weekday == 6:  # Sunday
        last = now - timedelta(days=2)
    else:
        last = now
    return last.strftime("%Y/%m/%d")


def clear_cache():
    """清除快取"""
    if os.path.exists(CACHE_FILE):
        os.remove(CACHE_FILE)
