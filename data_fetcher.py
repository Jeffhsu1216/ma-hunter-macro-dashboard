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
import json, os, logging

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

# (name, ticker, decimals, invert)
# invert=True：yfinance 給 USD/XXX，需倒轉為 XXX/USD（price=1/p，change_pct 取反）
FX_TICKERS = [
    ("DXY 美元指數", "DX-Y.NYB", 2, False),
    ("TWD/USD",     "TWD=X",    4, True),
    ("JPY/USD",     "JPY=X",    6, True),
    ("CNY/USD",     "CNY=X",    4, True),
    ("EUR/USD",     "EURUSD=X", 4, False),
    ("GBP/USD",     "GBPUSD=X", 4, False),
    ("AUD/USD",     "AUDUSD=X", 4, False),
    ("KRW/USD",     "KRW=X",    6, True),
]

# 殖利率用 FRED（官方來源，精準）
YIELD_FRED_IDS = [
    ("2Y",  "DGS2"),
    ("5Y",  "DGS5"),
    ("10Y", "DGS10"),
    ("30Y", "DGS30"),
]

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
]

# 加密貨幣（CoinGecko）
CRYPTO_LIST = [
    ("BTC", "bitcoin",  0),
    ("ETH", "ethereum", 2),
]

VIX_TICKER = "^VIX"

# 2026 全年 FOMC 會議日程（結束日）
FOMC_DATES_2026 = [
    date(2026, 1, 29), date(2026, 3, 19), date(2026, 5,  7),
    date(2026, 6, 18), date(2026, 7, 30), date(2026, 9, 17),
    date(2026, 11, 5), date(2026, 12, 17),
]

# 2026 全年 CBC（台灣央行）理監事會日程
CBC_DATES_2026 = [
    date(2026, 3, 19), date(2026, 6, 18),
    date(2026, 9, 17), date(2026, 12, 17),
]


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
           f"{_urlp.quote(ticker_symbol)}?interval=1d&range=5d")
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


def _fmt(price, decimals=2):
    if price is None:
        return "N/A"
    if price >= 10000:
        return f"{price:,.0f}"
    return f"{price:,.{decimals}f}"


def _get_fred_csv(series_id: str, n_rows: int = 10) -> list:
    """從 FRED 抓 CSV，回傳最後 n 筆 [(date_str, value_float), ...]"""
    try:
        url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
        resp = requests.get(url, timeout=(5, 25), headers={"User-Agent": "Mozilla/5.0"})
        lines = resp.text.strip().split("\n")[1:]
        results = []
        for line in lines[-n_rows:]:
            parts = line.split(",")
            if len(parts) == 2 and parts[1] != ".":
                results.append((parts[0], float(parts[1])))
        return results
    except Exception as e:
        logger.warning(f"FRED {series_id} failed: {e}")
        return []


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
    策略：先完全比對（lower），再關鍵詞比對
    """
    title_l = ff_title.lower().strip()

    # 1. 直接比對
    if title_l in tv_actuals:
        return tv_actuals[title_l]

    # 2. 去除常見後綴再比對（m/m, q/q, y/y, final, preliminary, flash）
    import re
    clean = re.sub(r'\s*(m/m|q/q|y/y|mom|qoq|yoy|final|preliminary|flash|revised)\s*$',
                   '', title_l).strip()
    for tv_title, val in tv_actuals.items():
        tv_clean = re.sub(r'\s*(m/m|q/q|y/y|mom|qoq|yoy|final|preliminary|flash|revised)\s*$',
                          '', tv_title).strip()
        if clean == tv_clean:
            return val

    # 3. 關鍵詞包含比對（取最長匹配）
    best_match, best_len = "", 0
    for tv_title, val in tv_actuals.items():
        if clean in tv_title or tv_title in clean:
            if len(tv_title) > best_len:
                best_match, best_len = val, len(tv_title)
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
    """統一 XXX/USD 格式解釋匯率變動，敘述順序與顯示排序一致（DXY 第一，其餘按漲跌幅）
    XXX/USD ↑ = XXX 走強、美元走弱
    DXY ↑ = 美元走強
    """
    # 各幣對的中文名稱與漲/跌時的解讀
    _META = {
        "TWD/USD": ("新台幣", "外資匯入支撐、進口成本下降",       "外資匯出壓力增、進口成本上升"),
        "JPY/USD": ("日圓",   "避險資金湧入日圓、日銀可能調整政策","日本出口競爭力增但進口通膨壓力大"),
        "CNY/USD": ("人民幣", "中國經濟信心回升或政策引導升值",   "中國資本外流壓力或政策寬鬆預期"),
        "KRW/USD": ("韓元",   "外資回流韓股、韓元走強",           "韓國出口導向受益但外資流出壓力"),
        "EUR/USD": ("歐元",   "歐洲經濟數據優於預期或 ECB 鷹派", "歐洲經濟疲弱或美元避險需求上升"),
        "GBP/USD": ("英鎊",   "英國經濟韌性或 BOE 偏鷹",         "英國經濟下行壓力或脫歐後續影響"),
        "AUD/USD": ("澳幣",   "大宗商品需求回升、中國經濟改善預期","商品價格走弱或全球風險趨避"),
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

        # ── 其餘幣對（按傳入順序，即漲跌幅排序）──
        meta = _META.get(name)
        if not meta or c is None or abs(c) < 0.2:
            continue
        cname, up_reason, dn_reason = meta
        if c > 0:
            parts.append(f"{cname}升值 {abs(c):.2f}%（{name} ↑），{up_reason}")
        else:
            parts.append(f"{cname}貶值 {abs(c):.2f}%（{name} ↓），{dn_reason}")

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


def _sentiment_commentary(fg: dict, vix_price: float = None, vix_chg: float = None) -> str:
    """根據 Fear & Greed + VIX 生成精準市場情緒解釋"""
    parts = []
    score = fg.get("score")
    change = fg.get("change")

    # CNN Fear & Greed — 精準分級
    if score is not None:
        if score <= 10:
            parts.append(f"CNN Fear & Greed 指數跌至 {score}，觸及歷史極端低位。"
                         f"該指數綜合動能、股價強度、Put/Call 比率、避險需求等七項指標，"
                         f"當前讀數顯示市場全面性恐慌")
        elif score <= 20:
            parts.append(f"CNN Fear & Greed 指數 {score}（極度恐懼），"
                         f"Put/Call 比率偏高、市場波動加劇，投資人大幅降低風險敞口")
        elif score <= 35:
            parts.append(f"CNN Fear & Greed 指數 {score}（恐懼），"
                         f"市場風險偏好下降，避險資產需求上升")
        elif score <= 50:
            parts.append(f"CNN Fear & Greed 指數 {score}（偏空中性），"
                         f"市場觀望氣氛濃，缺乏明確方向")
        elif score <= 65:
            parts.append(f"CNN Fear & Greed 指數 {score}（偏多中性），"
                         f"風險偏好溫和，資金持續進場")
        elif score <= 80:
            parts.append(f"CNN Fear & Greed 指數 {score}（貪婪），"
                         f"市場追漲意願強，估值可能偏高")
        else:
            parts.append(f"CNN Fear & Greed 指數 {score}（極度貪婪），"
                         f"市場過度樂觀，歷史上此水位後常見 5-10% 修正")

    # VIX — 精準分級
    if vix_price is not None:
        if vix_price > 40:
            parts.append(f"VIX 恐慌指數飆至 {vix_price:.1f}，已進入歷史前 5% 高位區間，"
                         f"隱含 S&P 500 未來 30 日年化波動率約 {vix_price:.0f}%，"
                         f"選擇權市場定價劇烈震盪")
        elif vix_price > 30:
            parts.append(f"VIX {vix_price:.1f}，高於長期均值（約 19-20）超過 50%，"
                         f"選擇權隱含波動率顯著擴大，市場預期未來一個月波動加劇")
        elif vix_price > 25:
            parts.append(f"VIX {vix_price:.1f}，高於長期均值，"
                         f"市場不確定性上升，避險成本增加")
        elif vix_price > 20:
            parts.append(f"VIX {vix_price:.1f}，略高於長期均值，市場存在一定不安情緒")
        elif vix_price > 15:
            parts.append(f"VIX {vix_price:.1f}，處於正常區間，市場波動溫和")
        else:
            parts.append(f"VIX 僅 {vix_price:.1f}，處於歷史低位，市場極度自滿，"
                         f"低波動環境下需警惕突發事件衝擊")

        # VIX 日變動
        if vix_chg is not None and abs(vix_chg) >= 5:
            if vix_chg > 15:
                parts.append(f"VIX 單日暴漲 {vix_chg:+.1f}%，選擇權避險需求急遽升溫")
            elif vix_chg > 5:
                parts.append(f"VIX 單日上升 {vix_chg:+.1f}%，避險情緒明顯加重")
            elif vix_chg < -10:
                parts.append(f"VIX 單日大跌 {vix_chg:+.1f}%，恐慌快速消退")
            elif vix_chg < -5:
                parts.append(f"VIX 單日回落 {vix_chg:+.1f}%，市場緊張情緒緩解")

    # CNN 日變動
    if change is not None:
        if abs(change) >= 10:
            direction = "急升" if change > 0 else "暴跌"
            parts.append(f"CNN 指數較前日{direction} {change:+.1f} 點，情緒面出現劇烈轉折")
        elif abs(change) >= 5:
            direction = "回升" if change > 0 else "惡化"
            parts.append(f"CNN 指數較前日{direction} {change:+.1f} 點")

    return "<br>".join(p + "。" for p in parts) if parts else "市場情緒數據暫無法取得。"


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
    taiex = next((i for i in valid if i["name"] == "加權指數"), None)
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
    sp = next((i for i in valid if i["name"] == "S&P 500"), None)
    nk = next((i for i in valid if i["name"] == "日經 225"), None)
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
    """抓匯率，統一 XXX/USD 格式，DXY 置頂，其餘按漲跌幅排序 + 解釋"""
    results = []
    for name, ticker, dec, inv in FX_TICKERS:
        q = _get_quote(ticker)
        price   = q["price"]
        chg_pct = q["change_pct"]
        if inv and price:
            price   = 1.0 / price
            chg_pct = (-chg_pct) if chg_pct is not None else None
        results.append({
            "name": name,
            "price": price,
            "price_fmt": _fmt(price, dec) if price is not None else "N/A",
            "change_pct": chg_pct,
        })

    # DXY 固定第一行，其餘按漲跌幅排序
    dxy = [r for r in results if "DXY" in r["name"]]
    others = [r for r in results if "DXY" not in r["name"]]
    others.sort(key=lambda x: x["change_pct"] if x["change_pct"] is not None else -999, reverse=True)
    results = dxy + others

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
    """央行利率：Fed 從 FRED 動態抓（失敗 fallback 4.25–4.50）+ ECB/BOJ hardcoded + CBC 會議日期自動算"""
    result = {}

    # ── Fed（FRED 動態抓，失敗 fallback hardcoded）──
    FED_FALLBACK = "4.25–4.50"  # 最後已知值，FRED 失敗時使用
    try:
        lower_rows = _get_fred_csv("DFEDTARL", 3)
        upper_rows = _get_fred_csv("DFEDTARU", 3)
        if lower_rows and upper_rows:
            fed_rate = f"{lower_rows[-1][1]:.2f}–{upper_rows[-1][1]:.2f}"
        else:
            fed_rate = FED_FALLBACK
            logger.warning("FRED Fed rate: empty rows, using fallback")
    except Exception as e:
        logger.warning(f"FRED Fed rate failed: {e}, using fallback")
        fed_rate = FED_FALLBACK

    result["Fed"] = {
        "rate": fed_rate,
        "next": _next_meeting(FOMC_DATES_2026),
    }

    # ── ECB（FRED ECBDFR 動態抓，失敗 fallback）──
    ECB_FALLBACK = "2.50"  # 2026-04 確認值（FRED ECBDFR 成功時回傳 2.50）
    try:
        ecb_rows = _get_fred_csv("ECBDFR", 3)
        ecb_rate = f"{ecb_rows[-1][1]:.2f}" if ecb_rows else ECB_FALLBACK
    except Exception as e:
        logger.warning(f"FRED ECB rate failed: {e}, using fallback")
        ecb_rate = ECB_FALLBACK
    result["ECB"] = {
        "rate": ecb_rate,
        "next": "手動更新",
    }

    # ── BOJ（FRED IRSTJPN156N 動態抓，失敗 fallback）──
    BOJ_FALLBACK = "0.50"  # 2026-04 確認值（2025-01 升至 0.50%）
    try:
        boj_rows = _get_fred_csv("IRSTJPN156N", 3)
        boj_rate = f"{boj_rows[-1][1]:.2f}" if boj_rows else BOJ_FALLBACK
    except Exception as e:
        logger.warning(f"FRED BOJ rate failed: {e}, using fallback")
        boj_rate = BOJ_FALLBACK
    result["BOJ"] = {
        "rate": boj_rate,
        "next": "手動更新",
    }

    # ── CBC（台灣央行重貼現率 — CBC 官網自動更新）──
    CBC_FALLBACK = "2.00"  # 2024-03-22 起確認值
    try:
        import re as _re
        _cbc_resp = requests.get(
            'https://www.cbc.gov.tw/tw/cp-534-4088-F0CAF-2.html',
            headers={'User-Agent': 'Mozilla/5.0'}, timeout=10)
        _m = _re.search(r'重貼現率.*?<em>([\d.]+)%</em>', _cbc_resp.text, _re.DOTALL)
        cbc_rate = _m.group(1) if _m else CBC_FALLBACK
    except Exception as e:
        logger.warning(f"CBC rate scrape failed: {e}, using fallback")
        cbc_rate = CBC_FALLBACK
    result["CBC"] = {
        "rate": cbc_rate,
        "next": _next_meeting(CBC_DATES_2026),
    }

    return result


def fetch_spx_technical() -> dict:
    """美股技術面：S&P 500 + Nasdaq  MA50/MA200/RSI(14)"""
    def _rsi(closes, period=14):
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

    import urllib.request as _urlr, urllib.parse as _urlp

    def _hist(symbol):
        url = f'https://query1.finance.yahoo.com/v8/finance/chart/{_urlp.quote(symbol)}?interval=1d&range=1y'
        req = _urlr.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with _urlr.urlopen(req, timeout=12) as r:
            d = json.loads(r.read())
        closes = d['chart']['result'][0]['indicators']['quote'][0]['close']
        return [x for x in closes if x is not None]

    result = {}
    for key, sym, name in [('spx', '^GSPC', 'S&P 500'), ('ndq', '^IXIC', 'Nasdaq')]:
        try:
            closes = _hist(sym)
            if len(closes) < 200:
                result[key] = {'ok': False, 'name': name}
                continue
            price  = closes[-1]
            ma50   = sum(closes[-50:])  / 50
            ma200  = sum(closes[-200:]) / 200
            rsi    = _rsi(closes)
            result[key] = {
                'ok':    True,
                'name':  name,
                'price': round(price, 2),
                'ma50':  round(ma50,  2),
                'ma200': round(ma200, 2),
                'pct50':  round((price - ma50)  / ma50  * 100, 1),
                'pct200': round((price - ma200) / ma200 * 100, 1),
                'rsi':     rsi,
                'rsi_lbl': _rsi_lbl(rsi),
                'cross':    '黃金交叉' if ma50 > ma200 else '死亡交叉',
                'cross_ok':  ma50 > ma200,
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


# ── 日曆事件：標題中文對照表 ──────────────────────────────────────────────────
_CAL_TITLE_ZH = {
    # 就業
    "Non Farm Payrolls":                   "非農就業人數",
    "Nonfarm Payrolls":                    "非農就業人數",
    "Unemployment Rate":                   "失業率",
    "Initial Jobless Claims":              "初次申請失業救濟",
    "Continuing Jobless Claims":           "持續申請失業救濟",
    "ADP Nonfarm Employment Change":       "ADP非農就業",
    "Average Hourly Earnings MoM":         "平均時薪 (月)",
    # 通膨
    "CPI MoM":                             "CPI (月)",
    "CPI YoY":                             "CPI (年)",
    "Core CPI MoM":                        "核心CPI (月)",
    "Core CPI YoY":                        "核心CPI (年)",
    "PPI MoM":                             "PPI (月)",
    "PPI YoY":                             "PPI (年)",
    "Core PPI MoM":                        "核心PPI (月)",
    "PCE Price Index MoM":                 "PCE物價指數 (月)",
    "Core PCE Price Index MoM":            "核心PCE (月)",
    "Import Prices MoM":                   "進口物價 (月)",
    "Export Prices MoM":                   "出口物價 (月)",
    # GDP / 生產
    "GDP Growth Rate QoQ":                 "GDP (季)",
    "GDP Growth Rate QoQ 2nd Est":         "GDP 二估 (季)",
    "GDP Growth Rate QoQ Final":           "GDP 終值 (季)",
    "Industrial Production MoM":           "工業生產 (月)",
    "Capacity Utilization Rate":           "產能利用率",
    "Manufacturing Production MoM":        "製造業產出 (月)",
    # 消費 / 零售
    "Retail Sales MoM":                    "零售銷售 (月)",
    "Core Retail Sales MoM":               "核心零售銷售 (月)",
    "Consumer Confidence":                 "消費者信心",
    "Michigan Consumer Sentiment":         "密西根消費者信心",
    "Michigan Consumer Sentiment Final":   "密西根消費者信心終值",
    # 房市
    "Existing Home Sales":                 "成屋銷售",
    "New Home Sales":                      "新屋銷售",
    "Housing Starts":                      "新屋開工",
    "Building Permits":                    "建築許可",
    "NAHB Housing Market Index":           "NAHB房市指數",
    "Pending Home Sales MoM":              "成屋簽約 (月)",
    # PMI / 景氣
    "ISM Manufacturing PMI":               "ISM製造業PMI",
    "ISM Non-Manufacturing PMI":           "ISM服務業PMI",
    "ISM Services PMI":                    "ISM服務業PMI",
    "S&P Global Manufacturing PMI":        "S&P製造業PMI",
    "Philadelphia Fed Manufacturing Index":"費城聯儲製造業",
    "NY Empire State Manufacturing Index": "紐約製造業指數",
    "Chicago PMI":                         "芝加哥PMI",
    "Dallas Fed Manufacturing Index":      "達拉斯製造業指數",
    # Fed / 利率
    "Federal Funds Rate":                  "聯邦基金利率",
    "FOMC Meeting Minutes":                "FOMC會議紀要",
    "Fed Interest Rate Decision":          "Fed利率決策",
    # 貿易 / 資本
    "Trade Balance":                       "貿易差額",
    "Balance of Trade":                    "貿易差額",
    "Net Long-term TIC Flows":             "長期資本淨流入",
    "Current Account":                     "經常帳",
    # 耐久財
    "Durable Goods Orders MoM":            "耐久財訂單 (月)",
    "Core Durable Goods Orders MoM":       "核心耐久財 (月)",
    # 台灣
    "GDP Growth Rate":                     "GDP成長率",
    "Export Orders YoY":                   "出口訂單 (年)",
    "Industrial Production YoY":           "工業生產 (年)",
    "Unemployment Rate":                   "失業率",
}

def _translate_cal_title(title: str) -> str:
    """回傳 '英文（中文）' 格式；找不到對照則只回傳英文"""
    zh = _CAL_TITLE_ZH.get(title)
    if zh:
        return f"{title}（{zh}）"
    # 模糊匹配（title 包含 key 或 key 包含 title）
    for k, v in _CAL_TITLE_ZH.items():
        if k.lower() in title.lower() or title.lower() in k.lower():
            return f"{title}（{v}）"
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

    high = [e for e in events_raw
            if e.get("impact") in ("High", "Medium")
            and e.get("country") in TARGET_COUNTRIES
            and (e.get("forecast") or "").strip()]   # 無預測數值（演講/聲明）直接過濾

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
    tv_from = (now_utc - timedelta(days=3)).strftime("%Y-%m-%dT00:00:00.000Z")
    tv_to   = (now_utc + timedelta(days=7)).strftime("%Y-%m-%dT23:59:59.000Z")

    # 同時抓 TV actuals（供 FF 事件補值用）
    tv_actuals = _fetch_tv_actuals(tv_from, tv_to)
    logger.info(f"TradingView actuals fetched: {len(tv_actuals)} events")

    # ── 主來源：ForexFactory ──
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

    # ── 備援一：TradingView 完整日曆 ──
    logger.info("FF failed, trying TradingView full calendar...")
    tv_results = _fetch_tv_calendar_full(tv_from, tv_to)
    if tv_results:
        try:
            with open(CALENDAR_BACKUP, "w", encoding="utf-8") as f:
                json.dump(tv_results, f, ensure_ascii=False)
        except:
            pass
        return tv_results

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
    comm_data   = fetch_commodity_data(); comm_ts   = _now_taipei()
    crypto_data = fetch_crypto_data(); crypto_ts    = _now_taipei()
    tw_data     = fetch_taiwan_market(); tw_ts      = _now_taipei()
    cal_data    = fetch_economic_calendar(); cal_ts  = _now_taipei()

    # 提取 VIX 給情緒解釋用
    vix_item = next((i for i in indices if i.get("is_vix")), None)
    vix_price = vix_item["price"] if vix_item else None
    vix_chg   = vix_item["change_pct"] if vix_item else None

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
        "sentiment_commentary": _sentiment_commentary(fg, vix_price, vix_chg),
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
