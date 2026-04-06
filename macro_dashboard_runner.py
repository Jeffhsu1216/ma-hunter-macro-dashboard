"""
總經儀表板 Runner v4（已優化）
執行：python3 macro_dashboard_runner.py
- threading 並行抓取所有數據（~8s）
- CNN Fear & Greed → alternative.me（無反爬）
- 三大法人 → TWSE BFI82U（金額彙總，正確欄位）
- 國際局勢 → Claude WebSearch 層補充
- 輸出 Telegram 推送
"""

import urllib.request, urllib.parse, json, datetime, threading, time

# ══════════════════════════════════════════
TOKEN   = '8743919766:AAG6z6YPW7Gqt7rF2KY2xC9mvbm2Ge31tjQ'
CHAT_ID = '2117347781'
DASHBOARD_URL = 'https://ma-hunter-macro-dashboard.onrender.com'
# ══════════════════════════════════════════

TZ = 8
def now_tw(): return datetime.datetime.utcnow() + datetime.timedelta(hours=TZ)

yf_results = {}
yf_lock    = threading.Lock()

def yf_quote(symbol):
    url = f'https://query1.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(symbol)}?interval=1d&range=5d'
    req = urllib.request.Request(url, headers={'User-Agent':'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            d = json.loads(r.read())
        closes = d['chart']['result'][0]['indicators']['quote'][0]['close']
        closes = [x for x in closes if x is not None]
        p = closes[-1]; prev = closes[-2] if len(closes) > 1 else p
        return p, p-prev, (p-prev)/prev*100 if prev else 0
    except: return None, None, None

def fetch_group(tickers, delay=0.15):
    for t in tickers:
        p, d, dp = yf_quote(t)
        with yf_lock: yf_results[t] = (p, d, dp)
        time.sleep(delay)

TICKERS = {
    'fx':    ['DX-Y.NYB','TWD=X','JPY=X','CNY=X','EURUSD=X','GBPUSD=X','AUDUSD=X','KRW=X'],
    'bonds': ['^IRX','^FVX','^TNX','^TYX'],
    'eq':    ['^GSPC','^IXIC','^TWII','^N225','^HSI','000001.SS','^GDAXI','^FTSE','^KS11','^VIX'],
    'comm':  ['GC=F','SI=F','CL=F','BZ=F','NG=F','HG=F'],
}

crypto_res = {}; fg_res = {}; inst_res = {}; cal_res = {}

def fetch_crypto():
    """Binance 24hr ticker — 免費、無需 Key、穩定"""
    try:
        import urllib.parse as _up
        syms = _up.quote('["BTCUSDT","ETHUSDT","SOLUSDT"]')
        req = urllib.request.Request(
            f'https://api.binance.com/api/v3/ticker/24hr?symbols={syms}',
            headers={'User-Agent':'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as r:
            data = {item['symbol']: item for item in json.loads(r.read())}
        def parse(sym):
            item = data.get(sym, {})
            return float(item['lastPrice']), float(item['priceChangePercent'])
        crypto_res['btc'] = parse('BTCUSDT')
        crypto_res['eth'] = parse('ETHUSDT')
        crypto_res['sol'] = parse('SOLUSDT')
        crypto_res['ok']  = True
    except Exception as e:
        crypto_res['ok'] = False

def fetch_fg():
    """alternative.me Fear & Greed（公開 API，無反爬，取代 CNN）"""
    try:
        req = urllib.request.Request('https://api.alternative.me/fng/?limit=2',
            headers={'User-Agent':'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=8) as r: d = json.loads(r.read())
        fg_res['score']  = int(d['data'][0]['value'])
        fg_res['rating'] = d['data'][0]['value_classification']
        fg_res['prev']   = int(d['data'][1]['value'])
        fg_res['ok']     = True
    except: fg_res['ok'] = False

def fetch_inst():
    """TWSE BFI82U — 三大法人日彙總（金額，非張數）"""
    check = now_tw()
    for _ in range(7):
        ds = check.strftime('%Y%m%d')
        try:
            req = urllib.request.Request(
                f'https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=json&dayDate={ds}&type=day',
                headers={'User-Agent':'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=10) as r: d = json.loads(r.read())
            if d.get('stat') == 'OK' and d.get('data'):
                rows = {row[0]: row for row in d['data']}
                def amt(row):
                    try: return int(row[3].replace(',',''))
                    except: return 0
                foreign = amt(rows.get('外資及陸資(不含外資自營商)',['','','','0'])) + \
                          amt(rows.get('外資自營商',['','','','0']))
                trust   = amt(rows.get('投信',['','','','0']))
                dealer  = amt(rows.get('自營商(自行買賣)',['','','','0'])) + \
                          amt(rows.get('自營商(避險)',['','','','0']))
                total   = amt(rows.get('合計',['','','','0']))
                inst_res.update(foreign=foreign, trust=trust, dealer=dealer,
                                total=total, date=check.strftime('%m/%d'), ok=True)
                return
        except: pass
        check -= datetime.timedelta(days=1)
    inst_res['ok'] = False

def fetch_cal():
    try:
        req = urllib.request.Request('https://nfs.faireconomy.media/ff_calendar_thisweek.json',
            headers={'User-Agent':'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=8) as r: raw = json.loads(r.read())
        cal_res['events'] = [e for e in raw if e.get('impact') == 'High'
                             and e.get('country') in {'USD','CNY','EUR','JPY','TWD'}]
        cal_res['ok'] = True
    except: cal_res['events'] = []; cal_res['ok'] = False

def run(geopolitics_bullets=None):
    NOW       = now_tw()
    DATE_STR  = NOW.strftime('%Y/%m/%d')
    WEEKDAY   = ['一','二','三','四','五','六','日'][NOW.weekday()]
    FETCH_TIME= NOW.strftime('%Y/%m/%d %H:%M')

    # ── 並行抓取 ──
    threads = []
    for g in TICKERS.values():
        th = threading.Thread(target=fetch_group, args=(g, 0.15))
        th.start(); threads.append(th); time.sleep(0.02)
    for fn in [fetch_crypto, fetch_fg, fetch_inst, fetch_cal]:
        th = threading.Thread(target=fn); th.start(); threads.append(th)
    for th in threads: th.join(timeout=30)

    def get(t):      return yf_results.get(t, (None, None, None))
    def pct(v):      return f'{v:+.2f}%' if v is not None else 'N/A'
    def arr(v):      return '▲' if v and v>0 else ('▼' if v and v<0 else '―')
    def fN(v, d=2):  return f'{v:,.{d}f}' if v is not None else 'N/A'
    def yi(v):       return f'{v/1e8:+.1f}億'

    def vix_lbl(v):
        for thr,lbl in [(15,'極度樂觀，市場平靜'),(20,'市場穩定，正常波動'),
                        (25,'輕度不安，需留意'),(30,'市場緊張，風險升高'),(40,'明顯恐慌，謹慎操作')]:
            if v < thr: return lbl
        return '極度恐慌'

    def fg_lbl(s):
        if s<=24: return 'Extreme Fear 極度恐懼'
        if s<=44: return 'Fear 恐懼'
        if s<=55: return 'Neutral 中性'
        if s<=74: return 'Greed 貪婪'
        return 'Extreme Greed 極度貪婪'

    L = []; A = L.append

    A(f'📊 <b>M&amp;A Hunter 總經儀表板</b>')
    A(f'<i>{DATE_STR}（週{WEEKDAY}）｜{FETCH_TIME} 台北時間</i>')
    A('')

    A('💱 <b>匯率</b>')
    for t,lbl in [('DX-Y.NYB','DXY 美元'),('TWD=X','USD/TWD'),('JPY=X','USD/JPY'),
                  ('EURUSD=X','EUR/USD'),('KRW=X','USD/KRW')]:
        p,d,dp = get(t)
        if p: A(f'  {lbl}  <code>{fN(p):>10}</code>  {arr(d)} {pct(dp)}')

    A('')
    A('📐 <b>美債殖利率</b>')
    for t,lbl in [('^IRX','2Y'),('^TNX','10Y'),('^TYX','30Y')]:
        p,d,dp = get(t)
        if p: A(f'  US {lbl}  <code>{fN(p)}%</code>  {arr(d)} {d:+.3f}bp')

    A('')
    A('📈 <b>全球股市</b>')
    for t,lbl in [('^GSPC','S&amp;P 500'),('^IXIC','Nasdaq  '),('^TWII','TAIEX   '),
                  ('^N225','日經 225'),('^HSI','恆生    '),('^GDAXI','DAX     '),('^KS11','KOSPI   ')]:
        p,d,dp = get(t)
        if p: A(f'  {lbl}  <code>{fN(p,0):>10}</code>  {arr(d)} {pct(dp)}')

    A('')
    A('🛢️ <b>原物料</b>')
    for t,lbl in [('GC=F','黃金 XAU'),('SI=F','白銀 XAG'),('CL=F','WTI 原油'),('BZ=F','布蘭特  ')]:
        p,d,dp = get(t)
        if p: A(f'  {lbl}  <code>${fN(p):>10}</code>  {arr(d)} {pct(dp)}')

    A('')
    A('🪙 <b>加密貨幣</b>')
    if crypto_res.get('ok'):
        for key,lbl in [('btc','BTC'),('eth','ETH'),('sol','SOL')]:
            p,c = crypto_res[key]
            s = f'${p:,.0f}' if p > 100 else f'${p:,.2f}'
            A(f'  {lbl}  <code>{s:>12}</code>  {arr(c)} {c:+.2f}%')

    A('')
    A('😱 <b>市場情緒</b>')
    vp,vd,vdp = get('^VIX')
    if vp: A(f'  VIX  <code>{fN(vp)}</code>  {arr(vd)} {pct(vdp)}  — {vix_lbl(vp)}')
    if fg_res.get('ok'):
        sc = fg_res['score']; pv = fg_res['prev']
        A(f'  Fear &amp; Greed  <code>{sc}</code>（{fg_lbl(sc)}）  前值 {pv}')

    A('')
    A('🏦 <b>央行利率</b>  Fed 4.25–4.50% ｜ ECB 2.50% ｜ BOJ 0.50% ｜ CBC 2.00%')

    A('')
    A(f'🇹🇼 <b>三大法人</b>（{inst_res.get("date","N/A")}）')
    if inst_res.get('ok'):
        tot = inst_res['total']
        A(f'  外資 {yi(inst_res["foreign"])} ｜ 投信 {yi(inst_res["trust"])} ｜ 自營商 {yi(inst_res["dealer"])}')
        A(f'  合計 <b>{yi(tot)}</b>  {"買超 ▲" if tot>0 else "賣超 ▼"}')

    # 本週經濟數據
    events = cal_res.get('events', [])
    if events:
        A('')
        A('📅 <b>本週重要數據</b>')
        for e in events[:5]:
            try:
                t_utc = datetime.datetime.fromisoformat(e['date'].replace('Z','+00:00'))
                tf = (t_utc + datetime.timedelta(hours=8)).strftime('%m/%d %H:%M')
                actual = e.get('actual',''); forecast = e.get('forecast','')
                status = '✅' if actual else '⏳'
                val = f'實際 {actual}' if actual else f'預期 {forecast}'
                A(f'  {status} {tf} | {e.get("country","")} | {e.get("title","")}  {val}')
            except: pass

    # 國際局勢（由 Claude 層 WebSearch 傳入）
    if geopolitics_bullets:
        A('')
        A('🌍 <b>國際局勢</b>')
        for b in geopolitics_bullets:
            A(f'  {b}')

    A('')
    A(f'🔗 <a href="{DASHBOARD_URL}">點此開啟完整儀表板 →</a>')

    msg = '\n'.join(L)

    # ── 送 Telegram ──
    payload = json.dumps({
        'chat_id': CHAT_ID,
        'text': msg,
        'parse_mode': 'HTML',
        'disable_web_page_preview': False,
        'reply_markup': json.dumps({'inline_keyboard':[[
            {'text':'📊 開啟完整儀表板','url': DASHBOARD_URL}
        ]]})
    }).encode()
    req = urllib.request.Request(
        f'https://api.telegram.org/bot{TOKEN}/sendMessage',
        data=payload, headers={'Content-Type':'application/json'})
    with urllib.request.urlopen(req, timeout=15) as r:
        result = json.loads(r.read())
    return result.get('ok', False), msg

if __name__ == '__main__':
    ok, msg = run()
    print('✅ 傳送成功' if ok else '❌ 傳送失敗')
    print(msg)
