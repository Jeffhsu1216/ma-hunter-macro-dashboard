#!/usr/bin/env python3
"""
meeting_prep.py — 拜訪前簡報包產生器
================================================
用法（CLI）:
  python3 meeting_prep.py <股票代號> [會議日期 YYYY-MM-DD]

  範例：
    python3 meeting_prep.py 3026
    python3 meeting_prep.py 3026 2026-04-15

用法（import）:
  from meeting_prep import run
  run("3026", "2026-04-15")

輸出：
  ~/Desktop/Claude/{代號}TT {公司名稱}/{代號}TT {公司名稱}_拜訪簡報包_{YYYYMMDD}.pdf
"""

import sys, os, json, datetime, urllib.request, textwrap, math
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

# ─── 顏色常數（圖靈金融配色）────────────────────────────────────────────────
NAVY    = (0.122, 0.306, 0.475)   # #1F4E79
GOLD    = (0.788, 0.659, 0.298)   # #C9A84C
LBLUE   = (0.851, 0.882, 0.949)   # #D9E1F2
GREEN   = (0.886, 0.937, 0.855)   # #E2EFDA
RED_BG  = (0.988, 0.894, 0.839)   # #FCE4D6
LGRAY   = (0.949, 0.949, 0.949)   # #F2F2F2
WHITE   = (1, 1, 1)
BLACK   = (0, 0, 0)
DGREEN  = (0.216, 0.337, 0.137)   # #375623
DRED    = (0.753, 0, 0)           # #C00000


# ─── Yahoo Finance 收盤價 ─────────────────────────────────────────────────────

def yf_close(code):
    for suffix in ['.TW', '.TWO']:
        try:
            url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
                   f"{code}{suffix}?range=3mo&interval=1d")
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            data = json.loads(urllib.request.urlopen(req, timeout=10).read())
            result = data['chart']['result'][0]
            closes = result['indicators']['quote'][0]['close']
            timestamps = result['timestamp']
            # 取最近有效值
            price = None
            for c in reversed(closes):
                if c is not None:
                    price = round(c, 2)
                    break
            # 3個月前第一個有效收盤價（計算漲跌幅用）
            first_close = None
            for c in closes:
                if c is not None:
                    first_close = round(c, 2)
                    break
            return price, first_close, suffix
        except Exception:
            continue
    return None, None, None


# ─── Five91 財務數據 ──────────────────────────────────────────────────────────

def fetch_metrics(code):
    try:
        url = f"https://five91.onrender.com/api/metrics?stock_id={code}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        data = json.loads(urllib.request.urlopen(req, timeout=15).read())
        # 格式1: {"stocks": {"3026": {...}}}
        if isinstance(data, dict) and 'stocks' in data:
            stocks = data['stocks']
            if code in stocks:
                return stocks[code]
            # 只有一筆時直接取第一個
            if len(stocks) == 1:
                return next(iter(stocks.values()))
        # 格式2: list
        if isinstance(data, list) and data:
            return data[0]
        # 格式3: 直接是 dict
        if isinstance(data, dict) and 'name' in data:
            return data
    except Exception as e:
        print(f"⚠️  five91 API 錯誤: {e}")
    return {}

def fetch_quarterly(code):
    """取近4季損益數據（保留舊介面）"""
    return []

def fetch_annual_financials(code, suffix='.TW'):
    """
    用 yfinance 取近3年年度財報
    回傳 list of dict：[{year, revenue, gross_margin, op_margin, net_income, eps}, ...]
    由新到舊排列
    """
    try:
        import yfinance as yf
        tk   = yf.Ticker(f"{code}{suffix}")
        fin  = tk.financials          # 年度損益（columns = datetime, index = 項目）
        info = tk.info or {}

        if fin is None or fin.empty:
            if suffix == '.TW':
                return fetch_annual_financials(code, '.TWO')
            return []

        rows = []
        for col in fin.columns[:3]:   # 最近3年
            year = col.year
            try:
                rev  = fin.loc['Total Revenue', col]       if 'Total Revenue'    in fin.index else None
                gp   = fin.loc['Gross Profit',  col]       if 'Gross Profit'     in fin.index else None
                oi   = fin.loc['Operating Income', col]    if 'Operating Income' in fin.index else None
                ni   = fin.loc['Net Income', col]          if 'Net Income'       in fin.index else None

                # 如果 Operating Income 不在，嘗試 EBIT
                if oi is None or (hasattr(oi, '__float__') and math.isnan(float(oi))):
                    oi = fin.loc['EBIT', col] if 'EBIT' in fin.index else None

                def safe(v):
                    if v is None: return None
                    try:
                        f = float(v)
                        return None if math.isnan(f) else f
                    except: return None

                rev_v = safe(rev)
                gp_v  = safe(gp)
                oi_v  = safe(oi)
                ni_v  = safe(ni)

                gm  = round(gp_v / rev_v * 100, 1) if rev_v and gp_v  else None
                om  = round(oi_v / rev_v * 100, 1) if rev_v and oi_v  else None
                rev_b = round(rev_v / 1e8, 1)       if rev_v          else None  # 億
                ni_b  = round(ni_v  / 1e8, 1)       if ni_v           else None

                rows.append({
                    'year': year,
                    'revenue':    rev_b,
                    'gross_margin': gm,
                    'op_margin':  om,
                    'net_income': ni_b,
                })
            except Exception:
                continue

        # YoY 計算
        for i in range(len(rows) - 1):
            cur = rows[i]['revenue']
            prv = rows[i+1]['revenue']
            if cur and prv and prv != 0:
                rows[i]['yoy'] = round((cur - prv) / prv * 100, 1)
            else:
                rows[i]['yoy'] = None
        if rows:
            rows[-1]['yoy'] = None

        return rows

    except Exception as e:
        print(f"⚠️  年度財報抓取失敗: {e}")
        if suffix == '.TW':
            return fetch_annual_financials(code, '.TWO')
        return []


# ─── 自動生成建議提問（三大面向）────────────────────────────────────────────

def generate_questions(m):
    """
    根據財務指標自動生成三大面向共 8 題提問。
    回傳：[('面向名稱', ['Q1', 'Q2', ...]), ...]
    """
    gm       = m.get('gross_margin_12q') or m.get('gross_margin') or 0
    om       = m.get('operating_margin_12q') or m.get('operating_margin') or 0
    debt_r   = m.get('debt_ratio') or 0
    free_cf  = m.get('free_cf')
    roe      = m.get('roe_12q') or 0
    category = m.get('category') or ''

    # ══ 面向一：財務體質（3題）══
    spread = (gm - om) if gm and om else 0
    if spread > 15:
        q_fin1 = f'毛利率 {gm:.0f}% vs 營益率 {om:.0f}%，費用率偏高（{spread:.0f}pp gap），降費的具體計畫與時程？'
    elif gm and gm < 20:
        q_fin1 = f'毛利率 {gm:.0f}% 偏低，改善毛利結構的具體做法（產品組合、漲價、降成本）？'
    elif gm and gm > 40:
        q_fin1 = f'毛利率 {gm:.0f}% 高於同業，維持此優勢的護城河與潛在風險點？'
    else:
        q_fin1 = f'毛利率 {gm:.0f}%，原物料 / 匯率對毛利的最新影響及下半年展望？'

    if roe and roe > 20:
        q_fin2 = f'ROE {roe:.0f}% 高於同業，高股東報酬的可持續性與主要驅動因子？'
    elif roe and roe < 0:
        q_fin2 = f'ROE {roe:.0f}% 為負，虧損主因及轉虧為盈的關鍵里程碑與時程？'
    elif roe and roe < 8:
        q_fin2 = f'ROE {roe:.0f}% 偏低，中期提升股東報酬的具體規劃？'
    else:
        q_fin2 = '今年 EPS 展望與主要拉動因子，法人預估與公司自評是否一致？'

    if debt_r > 50:
        q_fin3 = f'負債比率 {debt_r:.0f}% 偏高，短期到期債務分布與再融資計畫？'
    elif debt_r < 20:
        q_fin3 = f'負債比率 {debt_r:.0f}% 低，是否有加槓桿 / 策略收購的空間與規劃？'
    else:
        q_fin3 = f'負債比率 {debt_r:.0f}%，現金 / 銀行額度是否足以支應今年擴張需求？'

    # ══ 面向二：產業情況（3題）══
    if '半導體' in category or '電子' in category or '光電' in category:
        q_ind1 = 'AI / CoWoS / 先進封裝訂單能見度？客戶拉貨節奏與庫存調整現況？'
        q_ind2 = '關稅 / 美中貿易政策對供應鏈布局的影響？客戶是否要求在台 / 美增產？'
    elif '生技' in category or '醫療' in category or '醫藥' in category:
        q_ind1 = '主力產品在美 / 台 / 中法規申請進度，預計取證時程與關鍵里程碑？'
        q_ind2 = '競品上市 / 同類新藥研發動態，對定價權與市佔的實質影響？'
    elif '金融' in category or '銀行' in category or '保險' in category:
        q_ind1 = '利率走勢對 NIM / 投資收益的影響，及放款品質 NPL 最新狀況？'
        q_ind2 = '金融監管新規（IFRS 17 / 資本適足率）的衝擊與應對策略？'
    elif '能源' in category or '再生' in category or '電力' in category or '綠能' in category:
        q_ind1 = '離岸 / 陸域風電或太陽能案場的裝機量目標、簽約進度及 IRR 水準？'
        q_ind2 = '電力售價、躉購費率調整或 PPA 談判最新進展，對長期收益率的影響？'
    elif '零售' in category or '餐飲' in category or '消費' in category:
        q_ind1 = '同店成長率（SSSG）趨勢，通膨對消費力及客單價的實際影響？'
        q_ind2 = '新店展店計畫與租金 / 人力成本壓力，拓點速度是否調整？'
    else:
        q_ind1 = '今年產業需求最樂觀與最悲觀情境，公司目前判斷偏向哪一側？'
        q_ind2 = '在產業鏈的定位與主要競爭者的核心差異化優勢，最大威脅來自哪裡？'
    q_ind3 = '今年營收成長主要驅動力（產品 / 客戶 / 地區），能見度能看到哪一季？'

    # ══ 面向三：資本支出（2題）══
    if free_cf is not None and free_cf < 0:
        q_cap1 = f'FCF 為負（{free_cf:.1f}億），Capex 高峰預計何時結束？現金流轉正的前提條件與時程？'
    else:
        q_cap1 = '今年 Capex 預算金額與重點投資方向（擴產 / 自動化 / 研發 / 新市場）？'
    q_cap2 = 'Capex 投入的預期回報（產能增幅 / 效率提升 / 市佔擴張），達成目標 ROI 的時程？'

    return [
        ('財務體質', [q_fin1, q_fin2, q_fin3]),
        ('產業情況', [q_ind1, q_ind2, q_ind3]),
        ('資本支出', [q_cap1, q_cap2]),
    ]


# ─── PDF 生成 ─────────────────────────────────────────────────────────────────

def build_pdf(code, meeting_date_str, m, price, first_close, suffix, news_items, annual_rows, out_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm, mm
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # 字型設定（PingFang TC 優先，備用 Helvetica）
    # STHeiti：macOS 內建，支援繁體中文
    pdfmetrics.registerFont(TTFont('STHeitiL', '/System/Library/Fonts/STHeiti Light.ttc',  subfontIndex=0))
    pdfmetrics.registerFont(TTFont('STHeitiM', '/System/Library/Fonts/STHeiti Medium.ttc', subfontIndex=0))
    font_regular = 'STHeitiL'
    font_bold    = 'STHeitiM'

    W, H = A4  # 595 x 842 pt
    c = canvas.Canvas(out_path, pagesize=A4)

    name      = m.get('name', f'代號 {code}')
    category  = m.get('category', '—')
    market_cap= m.get('market_cap') or 0
    pe_val    = m.get('pe') or 0
    ev_ebitda = m.get('ev_ebitda') or 0
    gm        = m.get('gross_margin_12q') or m.get('gross_margin') or 0
    om        = m.get('operating_margin_12q') or m.get('operating_margin') or 0
    eps_val   = m.get('eps') or 0
    rev_ttm   = m.get('revenue_ttm') or 0
    net_inc   = m.get('net_income_ttm') or 0
    debt_r    = m.get('debt_ratio') or 0
    roe       = m.get('roe_12q') or 0
    pb        = m.get('pb') or 0
    div_yield = m.get('dividend_yield') or 0
    free_cf   = m.get('free_cf')

    # 3個月漲跌幅
    if price and first_close and first_close != 0:
        price_chg_3m = (price - first_close) / first_close * 100
    else:
        price_chg_3m = None

    market = '上市' if suffix == '.TW' else '上櫃'

    # ── 頁首橫幅 ──────────────────────────────────────────────────────────────
    c.setFillColorRGB(*NAVY)
    c.rect(0, H - 72, W, 72, fill=1, stroke=0)

    c.setFillColorRGB(*WHITE)
    c.setFont(font_bold, 18)
    c.drawString(20, H - 32, f"{code}  {name}")

    c.setFont(font_regular, 10)
    c.drawString(20, H - 52, f"{market} ｜ {category} ｜ 拜訪日期：{meeting_date_str}")

    # 右上角圖靈金融
    c.setFont(font_bold, 11)
    c.setFillColorRGB(*GOLD)
    c.drawRightString(W - 20, H - 32, "圖靈金融集團")
    c.setFont(font_regular, 8)
    c.setFillColorRGB(*WHITE)
    c.drawRightString(W - 20, H - 48, "Turing Financial Group")
    c.drawRightString(W - 20, H - 60, "jeff@turingfinancial.net")

    y = H - 85

    # ── Section 1：快速財務指標（6格）────────────────────────────────────────
    def draw_stat_box(x, y_top, w, h, label, value, sub='', bg=LBLUE, is_neg=False):
        c.setFillColorRGB(*bg)
        c.roundRect(x, y_top - h, w, h, 4, fill=1, stroke=0)
        # label
        c.setFillColorRGB(*NAVY)
        c.setFont(font_bold, 7.5)
        c.drawCentredString(x + w/2, y_top - 14, label)
        # value
        color = DRED if is_neg else DGREEN if 'is_pos' in str(is_neg) else BLACK
        c.setFillColorRGB(*BLACK)
        c.setFont(font_bold, 14)
        c.drawCentredString(x + w/2, y_top - 32, str(value))
        # sub
        c.setFont(font_regular, 7.5)
        c.setFillColorRGB(0.4, 0.4, 0.4)
        c.drawCentredString(x + w/2, y_top - 44, sub)

    box_w = (W - 40 - 10*5) / 6
    box_h = 54
    bx = 20

    # 股價
    price_str = f"NT${price:,.1f}" if price else '—'
    chg_str   = (f"3M: {price_chg_3m:+.1f}%" if price_chg_3m is not None else '')
    draw_stat_box(bx, y, box_w, box_h, '最新股價', price_str, chg_str)
    bx += box_w + 10

    # 市值（five91 已為億元單位）
    mc_str = f"NT${market_cap:.0f}億" if market_cap else '—'
    draw_stat_box(bx, y, box_w, box_h, '市值', mc_str)
    bx += box_w + 10

    # P/E
    pe_str = f"{pe_val:.1f}x" if pe_val and pe_val > 0 else '虧損'
    draw_stat_box(bx, y, box_w, box_h, 'P/E', pe_str)
    bx += box_w + 10

    # EV/EBITDA
    ev_str = f"{ev_ebitda:.1f}x" if ev_ebitda and ev_ebitda > 0 else '—'
    draw_stat_box(bx, y, box_w, box_h, 'EV/EBITDA', ev_str)
    bx += box_w + 10

    # 毛利率
    gm_str = f"{gm:.1f}%" if gm else '—'
    draw_stat_box(bx, y, box_w, box_h, '毛利率(TTM)', gm_str)
    bx += box_w + 10

    # 營益率
    om_str = f"{om:.1f}%" if om else '—'
    draw_stat_box(bx, y, box_w, box_h, '營益率(TTM)', om_str)

    y -= (box_h + 12)

    # ── 分隔線（向下畫，不覆蓋上方內容）────────────────────────────────────────
    HDR_H = 17
    def section_header(label, y_top):
        """y_top = 這個 section header 的頂部。向下畫 HDR_H pt，回傳底部 y。"""
        c.setFillColorRGB(*NAVY)
        c.rect(20, y_top - HDR_H, W - 40, HDR_H, fill=1, stroke=0)
        c.setFillColorRGB(*WHITE)
        c.setFont(font_bold, 9)
        c.drawString(25, y_top - HDR_H + 5, label)
        return y_top - HDR_H

    # ── Section 2：財務摘要（精簡4格，1列）───────────────────────────────────
    y -= 6
    y = section_header('📊  財務摘要', y)
    y -= 4

    fin_items = [
        ('EPS (TTM)',  f"NT${eps_val:.2f}" if eps_val else '—'),
        ('ROE (12Q)',  f"{roe:.1f}%"        if roe      else '—'),
        ('殖利率',     f"{div_yield:.1f}%"  if div_yield else '—'),
        ('負債比率',   f"{debt_r:.1f}%"     if debt_r   else '—'),
    ]

    col_w = (W - 40) / 4
    ROW_H = 17
    for i, (lbl, val) in enumerate(fin_items):
        fx = 20 + i * col_w
        bg = LGRAY if i % 2 == 0 else WHITE
        c.setFillColorRGB(*bg)
        c.rect(fx, y - ROW_H, col_w, ROW_H, fill=1, stroke=0)
        c.setFillColorRGB(*NAVY)
        c.setFont(font_regular, 8)
        c.drawString(fx + 5, y - ROW_H + 5, lbl)
        c.setFillColorRGB(*BLACK)
        c.setFont(font_bold, 9)
        c.drawRightString(fx + col_w - 6, y - ROW_H + 5, val)

    y -= (ROW_H + 4)

    # ── 頁尾輔助 ──────────────────────────────────────────────────────────────
    today_str = datetime.date.today().strftime('%Y/%m/%d')
    def draw_footer():
        c.setFillColorRGB(*NAVY)
        c.rect(0, 0, W, 28, fill=1, stroke=0)
        c.setFillColorRGB(*WHITE)
        c.setFont(font_regular, 7.5)
        c.drawString(20, 10, "Jeff Hsu, CFA｜Investment Manager｜Turing Financial Group")
        c.setFillColorRGB(*GOLD)
        c.drawRightString(W - 20, 10, f"製作日期：{today_str}  ｜  機密文件，僅供內部使用")

    # ── Section 3：近三年財報摘要 ────────────────────────────────────────────
    y -= 6
    y = section_header('📈  近三年財報摘要', y)
    y -= 4

    TBL_HDR_H = 15
    TBL_ROW_H = 14
    col_labels = ['年度', '營收（億）', '毛利率', '營益率', '稅後淨利（億）', 'YoY']
    ncols = len(col_labels)
    cw = (W - 40) / ncols

    if annual_rows:
        # 表頭
        for ci, lbl in enumerate(col_labels):
            fx = 20 + ci * cw
            c.setFillColorRGB(*LBLUE)
            c.rect(fx, y - TBL_HDR_H, cw, TBL_HDR_H, fill=1, stroke=0)
            c.setFillColorRGB(*NAVY)
            c.setFont(font_bold, 7.5)
            c.drawCentredString(fx + cw/2, y - TBL_HDR_H + 4, lbl)
        y -= TBL_HDR_H

        # 資料列
        for ri, row in enumerate(annual_rows):
            bg = LGRAY if ri % 2 == 0 else WHITE
            yoy = row.get('yoy')
            yoy_str   = f"{yoy:+.1f}%" if yoy is not None else '—'
            yoy_color = DGREEN if (yoy or 0) >= 0 else DRED
            vals = [
                str(row.get('year', '—')),
                f"{row['revenue']:.1f}"      if row.get('revenue')      else '—',
                f"{row['gross_margin']:.1f}%" if row.get('gross_margin') else '—',
                f"{row['op_margin']:.1f}%"    if row.get('op_margin')    else '—',
                f"{row['net_income']:.1f}"    if row.get('net_income')   else '—',
                yoy_str,
            ]
            for ci, val in enumerate(vals):
                fx = 20 + ci * cw
                c.setFillColorRGB(*bg)
                c.rect(fx, y - TBL_ROW_H, cw, TBL_ROW_H, fill=1, stroke=0)
                c.setFillColorRGB(*(yoy_color if ci == 5 and yoy is not None else BLACK))
                c.setFont(font_bold if ci == 0 else font_regular, 8)
                c.drawCentredString(fx + cw/2, y - TBL_ROW_H + 3, val)
            y -= TBL_ROW_H
    else:
        c.setFillColorRGB(0.5, 0.5, 0.5)
        c.setFont(font_regular, 8)
        c.drawString(20, y - 10, '（財報資料抓取失敗）')
        y -= 18

    # ── Section 4：近期重要訊息 ─────────────────────────────────────────────
    y -= 8   # 確保與上方表格有足夠間距
    y = section_header('📰  近期重要訊息（最近 90 天）', y)
    y -= 10

    if news_items:
        for idx, news in enumerate(news_items[:5]):
            title  = news.get('title', '')
            source = news.get('source', '')
            date_s = news.get('date', '')

            # 圓形序號
            c.setFillColorRGB(*NAVY)
            c.circle(28, y - 5, 6, fill=1, stroke=0)
            c.setFillColorRGB(*WHITE)
            c.setFont(font_bold, 7)
            c.drawCentredString(28, y - 8, str(idx + 1))

            # 標題
            c.setFillColorRGB(*BLACK)
            c.setFont(font_regular, 8.5)
            if len(title) > 56:
                c.drawString(42, y, title[:56])
                y -= 11
                c.drawString(42, y, title[56:112])
            else:
                c.drawString(42, y, title)

            # 來源+日期
            y -= 10
            c.setFillColorRGB(0.5, 0.5, 0.5)
            c.setFont(font_regular, 7)
            c.drawString(42, y, f"{source}  {date_s}".strip())
            y -= 12
    else:
        c.setFillColorRGB(0.5, 0.5, 0.5)
        c.setFont(font_regular, 8)
        c.drawString(25, y - 4, '（無近期新聞資料）')
        y -= 16

    # ── 頁尾 P1 ───────────────────────────────────────────────────────────────
    draw_footer()

    # ════════════════════════════════════════════════════
    # 第 2 頁：建議提問清單
    # ════════════════════════════════════════════════════
    c.showPage()

    # 第2頁頁首（簡版）
    c.setFillColorRGB(*NAVY)
    c.rect(0, H - 50, W, 50, fill=1, stroke=0)
    c.setFillColorRGB(*WHITE)
    c.setFont(font_bold, 14)
    c.drawString(20, H - 28, f"{code}  {name}  ｜  建議提問清單")
    c.setFont(font_regular, 9)
    c.drawString(20, H - 42, f"拜訪日期：{meeting_date_str}")
    c.setFont(font_bold, 10)
    c.setFillColorRGB(*GOLD)
    c.drawRightString(W - 20, H - 32, "圖靈金融集團")

    y2 = H - 68
    sections = generate_questions(m)

    # 面向對應色帶（左側色條用）
    SEC_COLORS = [NAVY, (0.13, 0.45, 0.30), (0.55, 0.27, 0.07)]
    q_global = 1

    for sec_idx, (sec_name, qs) in enumerate(sections):
        sec_color = SEC_COLORS[sec_idx % len(SEC_COLORS)]

        # ── 面向標題列 ──
        c.setFillColorRGB(*sec_color)
        c.roundRect(20, y2 - 18, W - 40, 20, 3, fill=1, stroke=0)
        c.setFillColorRGB(*WHITE)
        c.setFont(font_bold, 10)
        c.drawString(28, y2 - 12, f'【{sec_name}】')
        y2 -= 26

        for q in qs:
            ANS_H = 38   # 回答欄高度（緊湊版）
            Q_H   = 22   # 問題文字區高度

            # 左側色條
            c.setFillColorRGB(*sec_color)
            c.rect(20, y2 - Q_H - ANS_H - 4, 4, Q_H + ANS_H + 4, fill=1, stroke=0)

            # Q 編號徽章
            c.setFillColorRGB(*sec_color)
            c.circle(34, y2 - Q_H / 2 + 2, 8, fill=1, stroke=0)
            c.setFillColorRGB(*WHITE)
            c.setFont(font_bold, 7.5)
            c.drawCentredString(34, y2 - Q_H / 2 - 1, f'Q{q_global}')

            # 問題本文
            c.setFillColorRGB(*BLACK)
            c.setFont(font_regular, 9.5)
            lines = textwrap.wrap(q, 48)
            for li, line in enumerate(lines[:2]):
                c.drawString(48, y2 - 10 - li * 13, line)
            y2 -= Q_H

            # 回答欄（淺灰框）
            c.setFillColorRGB(*LGRAY)
            c.roundRect(26, y2 - ANS_H - 2, W - 46, ANS_H, 2, fill=1, stroke=0)
            c.setFillColorRGB(0.68, 0.68, 0.68)
            c.setFont(font_regular, 7.5)
            c.drawString(34, y2 - ANS_H + 4, '（會議紀錄）')
            y2 -= (ANS_H + 10)

            q_global += 1

        y2 -= 6   # 區塊間距

    # ── 頁尾 P2 ───────────────────────────────────────────────────────────────
    draw_footer()

    c.save()
    print(f"✅ PDF 已儲存：{out_path}")


# ─── WebSearch 新聞（透過 subprocess 呼叫 Claude 工具）────────────────────────

def fetch_news_via_websearch(code, name):
    """
    回傳格式：[{"title": ..., "source": ..., "date": ...}, ...]
    因為腳本本身無法呼叫 WebSearch，這裡回傳空清單，
    由 CLAUDE.md 流程中在 PDF 生成後由 Claude 補入。
    """
    return []


# ─── 主函式 ──────────────────────────────────────────────────────────────────

def run(code, meeting_date_str=None, news_items=None):
    """
    code             : str，股票代號（如 '3026'）
    meeting_date_str : str，會議日期 'YYYY-MM-DD'，預設今日
    news_items       : list，由 Claude 透過 WebSearch 抓取後傳入
                       格式：[{"title":"...","source":"...","date":"..."}]
    回傳 : str，輸出 PDF 路徑
    """
    if not meeting_date_str:
        meeting_date_str = datetime.date.today().strftime('%Y-%m-%d')

    print(f"\n🗂  拜訪前簡報包 — {code}｜{meeting_date_str}")
    print("━" * 50)

    # ── 抓財務數據 ──
    print("📡 從 five91 抓取財務指標...")
    m = fetch_metrics(code)
    name = m.get('name', '')
    # five91 找不到名稱時，從 TWSE OpenAPI 取中文簡稱作為 fallback
    if not name or name == code:
        try:
            twse_url = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
            req2 = urllib.request.Request(twse_url, headers={"User-Agent": "Mozilla/5.0"})
            twse_data = json.loads(urllib.request.urlopen(req2, timeout=10).read())
            for item in twse_data:
                if str(item.get('公司代號', '')) == str(code):
                    name = item.get('公司簡稱', '') or item.get('公司名稱', '')
                    break
        except Exception:
            pass
    if not name:
        name = code
    m['name'] = name   # 確保 build_pdf 能讀到正確公司名稱

    # ── yfinance fallback：five91 缺少財務指標時補齊 ──
    _needs_fallback = (not m.get('market_cap') and not m.get('gross_margin') and
                       not m.get('gross_margin_12q'))
    if _needs_fallback:
        try:
            import yfinance as yf
            info = yf.Ticker(f"{code}.TW").info
            TWD = 1e8  # 億
            def _pct(v): return round(v * 100, 2) if v else None

            if not m.get('market_cap') and info.get('marketCap'):
                m['market_cap'] = round(info['marketCap'] / TWD, 1)
            if not m.get('pe') and info.get('forwardPE'):
                m['pe'] = round(info['forwardPE'], 1)
            if not m.get('pb') and info.get('priceToBook'):
                m['pb'] = round(info['priceToBook'], 2)
            if not m.get('ev_ebitda') and info.get('enterpriseToEbitda'):
                m['ev_ebitda'] = round(info['enterpriseToEbitda'], 1)
            if not m.get('gross_margin_12q') and info.get('grossMargins'):
                m['gross_margin_12q'] = _pct(info['grossMargins'])
            if not m.get('operating_margin_12q') and info.get('operatingMargins'):
                m['operating_margin_12q'] = _pct(info['operatingMargins'])
            if not m.get('net_margin') and info.get('profitMargins'):
                m['net_margin'] = _pct(info['profitMargins'])
            if not m.get('roe_12q') and info.get('returnOnEquity'):
                m['roe_12q'] = _pct(info['returnOnEquity'])
            if not m.get('debt_ratio') and info.get('debtToEquity'):
                # debtToEquity = D/E，轉換為 D/(D+E) 負債比率
                de = info['debtToEquity'] / 100
                m['debt_ratio'] = round(de / (1 + de) * 100, 1)
            if not m.get('free_cf') and info.get('freeCashflow'):
                m['free_cf'] = round(info['freeCashflow'] / TWD, 2)
            if not m.get('revenue_ttm') and info.get('totalRevenue'):
                m['revenue_ttm'] = round(info['totalRevenue'] / TWD, 1)
            if not m.get('eps') and info.get('trailingEps'):
                m['eps'] = info['trailingEps']
            if not m.get('dividend_yield') and info.get('dividendYield'):
                m['dividend_yield'] = _pct(info['dividendYield'])
            print(f"   ⚡ yfinance fallback 已補齊財務指標")
        except Exception as e:
            print(f"   ⚠️  yfinance fallback 失敗: {e}")

    print(f"   公司：{name}｜產業：{m.get('category','—')}｜市值：NT${m.get('market_cap',0):.0f}億")

    # ── 抓股價 ──
    print("💹 從 Yahoo Finance 抓取近3個月股價...")
    price, first_close, suffix = yf_close(code)
    if price:
        chg = (price - first_close) / first_close * 100 if first_close else 0
        print(f"   現價：NT${price:,.1f}｜3M漲跌：{chg:+.1f}%")

    # ── 近三年年度財報 ──
    print("📊 從 yfinance 抓取近三年年度財報...")
    yf_suffix = suffix if suffix else '.TW'
    annual_rows = fetch_annual_financials(code, yf_suffix)
    if annual_rows:
        for r in annual_rows:
            yoy_str = f"YoY {r['yoy']:+.1f}%" if r.get('yoy') is not None else ''
            print(f"   {r['year']}  營收 NT${r.get('revenue','—')}億  毛利 {r.get('gross_margin','—')}%  {yoy_str}")

    # ── 新聞 ──
    if news_items is None:
        news_items = []

    # ── 確認輸出資料夾 ──
    base_dir  = os.path.expanduser('~/Desktop/Turing/上市櫃公司')
    folder    = os.path.join(base_dir, f"{code}TT {name}")
    os.makedirs(folder, exist_ok=True)

    date_compact = meeting_date_str.replace('-', '')
    filename  = f"{code}TT {name}_會前QA準備_{date_compact}.pdf"
    out_path  = os.path.join(folder, filename)

    # ── 生成 PDF ──
    print("📄 生成 PDF...")
    build_pdf(code, meeting_date_str, m, price, first_close, suffix, news_items, annual_rows, out_path)

    # ── 開啟 PDF ──
    subprocess.run(['open', out_path])
    print(f"🚀 已開啟：{out_path}")

    return out_path


# ─── CLI 入口 ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    code = sys.argv[1]
    date = sys.argv[2] if len(sys.argv) >= 3 else None

    # news_items 可由第三個參數傳入 JSON 字串
    news = []
    if len(sys.argv) >= 4:
        try:
            news = json.loads(sys.argv[3])
        except Exception:
            pass

    result = run(code, date, news)
    print(f"\n📁 輸出路徑：{result}")
