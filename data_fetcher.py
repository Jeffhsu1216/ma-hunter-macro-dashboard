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
    """Yahoo Finance v8 Chart API（直接 HTTP，不用 yfinance 避免 rate limit）"""
    import urllib.request as _urlr, urllib.parse as _urlp, time
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{_urlp.quote(ticker_symbol)}?interval=1d&range=5d")
    for attempt in range(retries + 1):
        try:
            req = _urlr.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with _urlr.urlopen(req, timeout=10) as r:
                d = json.loads(r.read())
            closes = d["chart"]["result"][0]["indicators"]["quote"][0]["close"]
            closes = [x for x in closes if x is not None]
            p    = closes[-1]
            prev = closes[-2] if len(closes) > 1 else p
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
        resp = requests.get(url, timeout=10)
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
    """以美元為基礎解釋匯率變動

    報價邏輯（判讀方式）：
      DXY         ↑ = 美元走強
      USD/XXX 系列 ↑ = 美元走強、該幣走弱（TWD, JPY, CNY, KRW）
      XXX/USD 系列 ↑ = 該幣走強、美元走弱（EUR, GBP, AUD）
    """
    fm = {f["name"]: f for f in fx_list}
    parts = []

    # ── DXY 總覽 ──
    dxy = fm.get("DXY 美元指數", {})
    dc = dxy.get("change_pct")
    if dc is not None:
        if dc > 0.5:
            parts.append(f"DXY 美元指數上漲 {dc:+.2f}%，美元全面走強，非美貨幣承壓")
        elif dc > 0.1:
            parts.append(f"DXY 小幅走強 {dc:+.2f}%，美元偏強格局")
        elif dc < -0.5:
            parts.append(f"DXY 下跌 {dc:+.2f}%，美元走弱，非美貨幣反彈")
        elif dc < -0.1:
            parts.append(f"DXY 微跌 {dc:+.2f}%，美元稍弱")
        else:
            parts.append("DXY 持平，匯市觀望")

    # ── 逐幣解釋（只解釋顯著變動的）──
    # USD/XXX 系列：數值↑ = 該幣貶值（對台灣投資人：TWD 貶 = 進口成本增加）
    usd_xxx = [
        ("USD/TWD", "新台幣", "外資匯出壓力增、進口成本上升", "外資匯入支撐、進口成本下降"),
        ("USD/JPY", "日圓",   "日本出口競爭力增但進口通膨壓力大", "避險資金湧入日圓、日銀可能調整政策"),
        ("USD/CNY", "人民幣", "中國資本外流壓力或政策寬鬆預期", "中國經濟信心回升或政策引導升值"),
        ("USD/KRW", "韓元",   "韓國出口導向受益但外資流出壓力", "外資回流韓股、韓元走強"),
    ]
    for pair, cname, up_reason, dn_reason in usd_xxx:
        item = fm.get(pair, {})
        c = item.get("change_pct")
        if c is not None and abs(c) >= 0.2:
            if c > 0:
                parts.append(f"{cname}貶值 {abs(c):.2f}%（{pair} ↑），{up_reason}")
            else:
                parts.append(f"{cname}升值 {abs(c):.2f}%（{pair} ↓），{dn_reason}")

    # XXX/USD 系列：數值↑ = 該幣升值 / 美元走弱
    xxx_usd = [
        ("EUR/USD", "歐元", "歐洲經濟數據優於預期或 ECB 鷹派", "歐洲經濟疲弱或美元避險需求上升"),
        ("GBP/USD", "英鎊", "英國經濟韌性或 BOE 偏鷹", "英國經濟下行壓力或脫歐後續影響"),
        ("AUD/USD", "澳幣", "大宗商品需求回升、中國經濟改善預期", "商品價格走弱或全球風險趨避"),
    ]
    for pair, cname, up_reason, dn_reason in xxx_usd:
        item = fm.get(pair, {})
        c = item.get("change_pct")
        if c is not None and abs(c) >= 0.2:
            if c > 0:
                parts.append(f"{cname}升值 {abs(c):.2f}%（{pair} ↑），{up_reason}")
            else:
                parts.append(f"{cname}貶值 {abs(c):.2f}%（{pair} ↓），{dn_reason}")

    return "。".join(parts) + "。" if parts else "匯率整體變動不大，市場觀望氣氛濃厚。"


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

    return "。".join(parts) + "。" if parts else "殖利率變動不大，市場等待新催化劑。"


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

    return "。".join(parts) + "。" if parts else "市場情緒數據暫無法取得。"


def _commodity_commentary(cm) -> str:
    """根據原物料+加密數據生成詳細解釋。cm 可以是 list 或 dict"""
    if isinstance(cm, list):
        cm = {c["name"]: c for c in cm}
    parts = []

    # ── 原油（全球通膨領先指標）──
    wti = cm.get("WTI 原油", {})
    brent = cm.get("布蘭特原油", {})
    wti_c = wti.get("change_pct")
    wti_p = wti.get("price", 0)
    brent_c = brent.get("change_pct")
    brent_p = brent.get("price", 0)

    if wti_c is not None:
        if wti_c > 3:
            parts.append(f"WTI 原油大漲 {wti_c:+.2f}% 至 ${wti_p:.2f}，"
                         f"可能受地緣衝突升級（中東局勢/制裁）或 OPEC+ 減產等供給面因素推動")
        elif wti_c > 1:
            parts.append(f"WTI 上漲 {wti_c:+.2f}% 至 ${wti_p:.2f}，供需面偏緊")
        elif wti_c < -3:
            parts.append(f"WTI 重挫 {wti_c:+.2f}% 至 ${wti_p:.2f}，"
                         f"可能反映全球需求放緩預期、美國庫存意外增加或 OPEC+ 增產訊號")
        elif wti_c < -1:
            parts.append(f"WTI 下跌 {wti_c:+.2f}% 至 ${wti_p:.2f}，需求面偏弱")
        else:
            parts.append(f"WTI 原油 ${wti_p:.2f}（{wti_c:+.2f}%），價格窄幅震盪")

    if wti_p and wti_p > 90:
        parts.append(f"油價處於 ${wti_p:.0f} 高位，推升運輸與製造成本，"
                     f"通膨壓力可能延遲 Fed 降息時程")
    elif wti_p and wti_p < 60:
        parts.append(f"油價跌至 ${wti_p:.0f} 低位，有利消費者但打擊能源類股獲利")

    # WTI vs Brent 價差
    if wti_p and brent_p and brent_p > 0:
        spread = brent_p - wti_p
        if spread > 8:
            parts.append(f"Brent-WTI 價差擴至 ${spread:.1f}，反映國際市場供給緊於美國")

    # ── 天然氣 ──
    ng = cm.get("天然氣", {})
    ng_c = ng.get("change_pct")
    ng_p = ng.get("price", 0)
    if ng_c is not None and abs(ng_c) > 1.5:
        direction = "上漲" if ng_c > 0 else "下跌"
        reason = "冬季取暖需求或供給中斷" if ng_c > 0 else "暖冬或庫存充足"
        parts.append(f"天然氣{direction} {abs(ng_c):.2f}% 至 ${ng_p:.3f}，{reason}，"
                     f"影響電力與工業成本")

    # ── 黃金（避險 + 通膨對沖）──
    gold = cm.get("黃金", {})
    gold_c = gold.get("change_pct")
    gold_p = gold.get("price", 0)
    if gold_c is not None:
        if gold_c > 1.5:
            parts.append(f"黃金大漲 {gold_c:+.2f}% 至 ${gold_p:,.0f}，"
                         f"避險需求與央行購金雙重推動，實質利率下行預期支撐金價")
        elif gold_c > 0.3:
            parts.append(f"黃金上漲 {gold_c:+.2f}% 至 ${gold_p:,.0f}，避險買盤持續")
        elif gold_c < -1.5:
            parts.append(f"黃金下跌 {gold_c:+.2f}%，美元走強或實質利率上升壓抑金價")
        elif gold_c < -0.3:
            parts.append(f"黃金微跌 {gold_c:+.2f}%，獲利了結賣壓")

    if gold_p and gold_p > 4000:
        parts.append(f"金價 ${gold_p:,.0f} 處於歷史高位，反映去美元化趨勢與全球央行儲備多元化")

    # ── 白銀（工業 + 貴金屬雙重屬性）──
    silver = cm.get("白銀", {})
    silver_c = silver.get("change_pct")
    if silver_c is not None and abs(silver_c) > 1.5:
        direction = "上漲" if silver_c > 0 else "下跌"
        parts.append(f"白銀{direction} {abs(silver_c):.2f}%，"
                     f"兼具工業（光伏、電子）與貴金屬屬性，"
                     f"{'跟隨金價走強' if silver_c > 0 else '工業需求疑慮拖累'}")

    # ── 銅（全球景氣晴雨表）──
    copper = cm.get("銅", {})
    copper_c = copper.get("change_pct")
    copper_p = copper.get("price", 0)
    if copper_c is not None:
        if abs(copper_c) > 1.5:
            direction = "上漲" if copper_c > 0 else "下跌"
            signal = ("全球製造業 PMI 改善或中國基建需求回升"
                      if copper_c > 0 else "工業需求放緩，中國經濟復甦不如預期")
            parts.append(f"銅價{direction} {abs(copper_c):.2f}% 至 ${copper_p:.3f}/磅，"
                         f"「銅博士」暗示{signal}")
        elif abs(copper_c) > 0.5:
            direction = "小漲" if copper_c > 0 else "小跌"
            parts.append(f"銅價{direction} {abs(copper_c):.2f}%，反映工業活動溫和波動")

    # ── BTC / ETH（風險偏好 + 流動性指標）──
    btc = cm.get("BTC", {})
    eth = cm.get("ETH", {})
    btc_c = btc.get("change_pct")
    btc_p = btc.get("price", 0)
    eth_c = eth.get("change_pct")

    if btc_c is not None:
        if btc_c > 5:
            parts.append(f"BTC 大漲 {btc_c:+.2f}% 至 ${btc_p:,.0f}，"
                         f"機構資金流入或監管利多，風險偏好顯著回升")
        elif btc_c > 2:
            parts.append(f"BTC 上漲 {btc_c:+.2f}%，加密市場情緒偏多")
        elif btc_c < -5:
            parts.append(f"BTC 重挫 {btc_c:+.2f}%，流動性收緊或大戶拋售，風險資產全面承壓")
        elif btc_c < -2:
            parts.append(f"BTC 下跌 {btc_c:+.2f}%，加密市場風險偏好下降")
        else:
            parts.append(f"BTC ${btc_p:,.0f}（{btc_c:+.2f}%），窄幅整理")

    if eth_c is not None and btc_c is not None:
        if eth_c - btc_c > 3:
            parts.append("ETH 明顯跑贏 BTC，山寨幣輪動行情啟動")
        elif btc_c - eth_c > 3:
            parts.append("BTC 獨強、ETH 落後，資金集中流向比特幣避險")

    return "。".join(parts) + "。" if parts else "原物料整體變動不大，市場觀望。"


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

    return "。".join(parts) + "。" if parts else ""


# ============================================================
# 各區塊抓取函式
# ============================================================

def fetch_fx_data() -> dict:
    """抓匯率，DXY 置頂，其餘按漲跌幅排序 + 解釋"""
    results = []
    for name, ticker, dec in FX_TICKERS:
        q = _get_quote(ticker)
        results.append({
            "name": name,
            "price": q["price"],
            "price_fmt": _fmt(q["price"], dec),
            "change_pct": q["change_pct"],
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
    """央行利率：Fed 從 FRED 動態抓 + 會議日期自動算"""
    result = {}

    # ── Fed ──
    try:
        lower_rows = _get_fred_csv("DFEDTARL", 3)
        upper_rows = _get_fred_csv("DFEDTARU", 3)
        if lower_rows and upper_rows:
            fed_rate = f"{lower_rows[-1][1]:.2f}–{upper_rows[-1][1]:.2f}"
        else:
            fed_rate = "N/A"
    except Exception as e:
        logger.warning(f"FRED Fed rate failed: {e}")
        fed_rate = "N/A"

    result["Fed"] = {
        "rate": fed_rate,
        "next": _next_meeting(FOMC_DATES_2026),
    }

    # ── CBC ──
    result["CBC"] = {
        "rate": "2.00",
        "next": _next_meeting(CBC_DATES_2026),
    }

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


def _parse_calendar_events(events_raw: list) -> list:
    high = [e for e in events_raw
            if e.get("impact") == "High"
            and e.get("country") in ("USD", "CNY", "EUR", "JPY", "TWD")]
    results = []
    for e in high[:20]:  # 多取一些，過濾後才限制數量
        actual = (e.get("actual") or "").strip()
        if not actual:
            continue  # 只顯示已公布的數據

        dt_str = e.get("date", "")
        try:
            dt_utc = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
            dt_taipei = dt_utc.astimezone(TAIPEI_TZ)
            date_fmt = dt_taipei.strftime("%m/%d %H:%M")
        except:
            date_fmt = dt_str[:10]

        forecast = (e.get("forecast") or "").strip() or "—"
        previous = (e.get("previous") or "").strip() or "—"

        # 判斷超預期 / 低於預期（純數字比較）
        beat_indicator = ""
        try:
            def _to_num(s):
                return float(s.replace("%", "").replace("K", "000").replace("M", "000000").strip())
            if forecast != "—":
                act_num = _to_num(actual)
                fct_num = _to_num(forecast)
                beat_indicator = "▲" if act_num > fct_num else ("▼" if act_num < fct_num else "")
        except:
            pass

        results.append({
            "date": date_fmt,
            "country": e.get("country", ""),
            "title": e.get("title", ""),
            "forecast": forecast,
            "previous": previous,
            "actual": actual,
            "beat_indicator": beat_indicator,
        })

        if len(results) >= 12:
            break

    return results


def fetch_economic_calendar() -> list:
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
                try:
                    with open(CALENDAR_BACKUP, "w", encoding="utf-8") as f:
                        json.dump(results, f, ensure_ascii=False)
                except:
                    pass
                return results
        except Exception as e:
            logger.warning(f"Calendar {week} failed: {e}")
            continue

    try:
        if os.path.exists(CALENDAR_BACKUP):
            with open(CALENDAR_BACKUP, "r", encoding="utf-8") as f:
                return json.load(f)
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

    fx_data     = fetch_fx_data();      fx_ts       = _now_taipei()
    yields_data = fetch_yield_curve();  yields_ts   = _now_taipei()
    cb_data     = fetch_cb_rates();     cb_ts       = _now_taipei()
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

    # 合併原物料 + 加密貨幣用於 commentary
    all_comm = comm_data["items"] + crypto_data
    combined_commentary = _commodity_commentary(
        {c["name"]: c for c in all_comm}  # pass as dict for lookup
    )

    # 殖利率：優先用 FRED 資料日期（格式 YYYY-MM-DD → YYYY/MM/DD）
    fred_date = yields_data.get("10Y", {}).get("date") or yields_data.get("2Y", {}).get("date")
    if fred_date:
        try:
            yields_ts = datetime.strptime(fred_date, "%Y-%m-%d").strftime("%Y/%m/%d") + "（FRED）"
        except:
            pass

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
        "yields":        yields_data,
        "yields_commentary": yields_data.pop("commentary", ""),
        "yields_updated_at": yields_ts,
        "cb_rates":      cb_data,
        "cb_rates_updated_at": cb_ts,
        "indices":       indices,
        "indices_updated_at": indices_ts,
        "commodities":        comm_data["items"],
        "crypto":             crypto_data,
        "commodities_commentary": combined_commentary,
        "commodities_updated_at": comm_ts,
        "crypto_updated_at":  crypto_ts,
        "fear_greed":    fg,
        "sentiment_commentary": _sentiment_commentary(fg, vix_price, vix_chg),
        "fg_updated_at": fg_ts,
        "calendar":      cal_data,
        "calendar_updated_at": cal_ts,
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
