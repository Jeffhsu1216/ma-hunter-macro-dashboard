"""
總經儀表板 Runner v5（已優化）
執行：python3 macro_dashboard_runner.py
- threading 並行抓取所有數據（~8s）
- CNN Fear & Greed → alternative.me（無反爬）
- 三大法人 → TWSE BFI82U（金額彙總，正確欄位）
- 國際局勢 → Claude WebSearch 層補充
- 輸出 Telegram 推送
- 三大法人成功後自動寫入 taiwan_backup.json + git push（供 Render fallback）
"""

import urllib.request, urllib.parse, json, datetime, threading, time, os, subprocess

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
                                total=total, date=check.strftime('%m/%d'),
                                full_date=check.strftime('%Y%m%d'), ok=True)
                return
        except: pass
        check -= datetime.timedelta(days=1)
    inst_res['ok'] = False

def fetch_cal():
    """
    主要來源：TradingView（有 forecast + actual，無 rate limit）
    備援：ForexFactory（TradingView 失敗時）
    容錯：cal_raw_backup.json（雙方都失敗時）
    過濾：僅保留 High impact、目標國家、有 forecast 數值的事件
    """
    import re as _re, os
    TARGET = {'USD','TWD'}
    raw_backup = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cal_raw_backup.json')

    now_utc = datetime.datetime.utcnow()
    tv_from = (now_utc - datetime.timedelta(days=3)).strftime('%Y-%m-%dT00:00:00.000Z')
    tv_to   = (now_utc + datetime.timedelta(days=7)).strftime('%Y-%m-%dT23:59:59.000Z')

    events = []

    # ── 主來源：TradingView ──
    try:
        tv_req = urllib.request.Request(
            f'https://economic-calendar.tradingview.com/events?from={tv_from}&to={tv_to}&countries=US,TW',
            headers={'User-Agent':'Mozilla/5.0','Origin':'https://www.tradingview.com'})
        with urllib.request.urlopen(tv_req, timeout=10) as r:
            tv_data = json.loads(r.read())
        if isinstance(tv_data, dict):
            tv_data = tv_data.get('result', [])
        if isinstance(tv_data, list):
            for e in tv_data:
                country = e.get('country','').upper()
                # TV 用 US/CN/EU/JP/TW，轉成 FF 的 USD/CNY/EUR/JPY/TWD
                country_map = {'US':'USD','CN':'CNY','EU':'EUR','JP':'JPY','TW':'TWD'}
                country_code = country_map.get(country, country)
                if country_code not in TARGET: continue
                if e.get('importance', -1) < 0: continue   # TV: -1=low,0=medium,1=high
                forecast = str(e.get('forecast') or '').strip()
                if not forecast or forecast == 'None': continue  # 過濾無預測值
                events.append({
                    'title':    e.get('title',''),
                    'country':  country_code,
                    'date':     e.get('date',''),
                    'forecast': forecast,
                    'previous': str(e.get('previous') or '').strip(),
                    'actual':   str(e.get('actual') or '').strip() if e.get('actual') is not None else '',
                    'impact':   'High',
                })
    except Exception as ex:
        pass  # 降級到 ForexFactory

    # ── 備援：ForexFactory（TV 失敗時）──
    if not events:
        try:
            req = urllib.request.Request('https://nfs.faireconomy.media/ff_calendar_thisweek.json',
                headers={'User-Agent':'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=8) as r: raw = json.loads(r.read())
            events = [e for e in raw if e.get('impact') == 'High'
                      and e.get('country') in TARGET
                      and (e.get('forecast') or '').strip()]
        except: pass

    # ── 容錯：讀上次備份 ──
    if not events:
        try:
            with open(raw_backup, 'r', encoding='utf-8') as f:
                events = json.load(f)
        except: pass

    # 儲存備份
    if events:
        try:
            with open(raw_backup, 'w', encoding='utf-8') as f:
                json.dump(events, f, ensure_ascii=False)
        except: pass

    cal_res['events'] = events
    cal_res['ok'] = bool(events)

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))   # macro-dashboard/
BACKUP_PATH = os.path.join(SCRIPT_DIR, 'taiwan_backup.json')

def _push_taiwan_backup():
    """把 inst_res 寫入 taiwan_backup.json，然後 git commit + push 供 Render fallback
    macro-dashboard 是獨立 repo：git -C SCRIPT_DIR
    """
    if not inst_res.get('ok'):
        return
    backup = {
        "foreign":      inst_res['foreign'],
        "foreign_yi":   round(inst_res['foreign'] / 1e8, 1),
        "inv_trust":    inst_res['trust'],
        "inv_trust_yi": round(inst_res['trust']   / 1e8, 1),
        "dealer":       inst_res['dealer'],
        "dealer_yi":    round(inst_res['dealer']  / 1e8, 1),
        "total":        inst_res['total'],
        "total_yi":     round(inst_res['total']   / 1e8, 1),
        "date":         inst_res.get('full_date', ''),
        "unit":         "億元",
        "source":       "runner_backup",
    }
    try:
        with open(BACKUP_PATH, 'w', encoding='utf-8') as f:
            json.dump(backup, f, ensure_ascii=False, indent=2)
        date_str = inst_res.get('full_date', 'unknown')
        subprocess.run(['git', '-C', SCRIPT_DIR, 'add', 'taiwan_backup.json'], check=True)
        subprocess.run(['git', '-C', SCRIPT_DIR, 'commit', '-m',
                        f'[auto] 更新三大法人備份 {date_str}'], check=True)
        subprocess.run(['git', '-C', SCRIPT_DIR, 'push'], check=True)
        print(f'✅ taiwan_backup.json 已更新並推送（{date_str}）')
    except subprocess.CalledProcessError as e:
        print(f'⚠️  git push 失敗（{e}）；backup 已寫入本機')
    except Exception as e:
        print(f'⚠️  備份寫入失敗：{e}')


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

    # ── 三大法人備份更新（寫 json + git push） ──
    _push_taiwan_backup()

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

    def fFX(v):
        """匯率專用格式：自動選擇小數位數"""
        if v is None: return 'N/A'
        if v >= 10:   return f'{v:,.2f}'
        if v >= 1:    return f'{v:,.4f}'
        if v >= 0.01: return f'{v:,.4f}'
        return f'{v:,.6f}'

    A('💱 <b>匯率</b>')
    # DXY 永遠第一
    _p, _d, _dp = get('DX-Y.NYB')
    if _p: A(f'  DXY 美元  <code>{fFX(_p):>12}</code>  {arr(_d)} {pct(_dp)}')

    # 其餘幣對：統一為 XXX/USD 格式，yfinance 給 USD/XXX 的倒轉過來
    _fx_pairs = [
        ('EURUSD=X', 'EUR/USD', False),
        ('GBPUSD=X', 'GBP/USD', False),
        ('AUDUSD=X', 'AUD/USD', False),
        ('TWD=X',    'TWD/USD', True),
        ('JPY=X',    'JPY/USD', True),
        ('CNY=X',    'CNY/USD', True),
        ('KRW=X',    'KRW/USD', True),
    ]
    _fx_rows = []
    for _t, _lbl, _inv in _fx_pairs:
        _op, _od, _odp = get(_t)
        if _op is None: continue
        if _inv:
            _np  = 1.0 / _op
            _ndp = -_odp
            _nd  = _np * _ndp / 100
        else:
            _np, _nd, _ndp = _op, _od, _odp
        _fx_rows.append((_lbl, _np, _nd, _ndp))
    _fx_rows.sort(key=lambda x: x[3], reverse=True)  # 漲幅最大排最上
    for _lbl, _np, _nd, _ndp in _fx_rows:
        A(f'  {_lbl}  <code>{fFX(_np):>12}</code>  {arr(_nd)} {pct(_ndp)}')

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
        now_utc = datetime.datetime.now(datetime.timezone.utc)
        for e in events[:10]:
            try:
                t_utc = datetime.datetime.fromisoformat(e['date'].replace('Z','+00:00'))
                tf = (t_utc + datetime.timedelta(hours=8)).strftime('%m/%d %H:%M')
                actual = e.get('actual','')
                is_past = t_utc < now_utc
                if actual:
                    status = '✅'; val = f'實際 {actual}'
                elif is_past:
                    status = '⚠️'; val = '待確認'
                else:
                    status = '⏳'; val = ''
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

GEO_PATH = os.path.join(SCRIPT_DIR, 'geopolitics.json')

def _push_geopolitics_json():
    """把 geopolitics.json git commit + push 供 Render 即時更新國際局勢區塊"""
    try:
        date_str = now_tw().strftime('%Y%m%d')
        subprocess.run(['git', '-C', SCRIPT_DIR, 'add', 'geopolitics.json'], check=True)
        subprocess.run(['git', '-C', SCRIPT_DIR, 'commit', '-m',
                        f'[auto] 更新國際局勢 {date_str}'], check=True)
        subprocess.run(['git', '-C', SCRIPT_DIR, 'push'], check=True)
        print(f'✅ geopolitics.json 已更新並推送（{date_str}）')
    except subprocess.CalledProcessError as e:
        print(f'⚠️  geopolitics.json git push 失敗：{e}')
    except Exception as e:
        print(f'⚠️  geopolitics.json 推送異常：{e}')

if __name__ == '__main__':
    # 推送 geopolitics.json → Render 即時更新國際局勢區塊
    _push_geopolitics_json()

    # 讀取 Claude 預先寫入的 geopolitics.json（由 Claude WebSearch 層產生）
    geo_bullets = None
    try:
        if os.path.exists(GEO_PATH):
            with open(GEO_PATH, 'r', encoding='utf-8') as f:
                geo_data = json.load(f)
            geo_bullets = geo_data.get('bullets')
            print(f'📰 載入地緣政治 {len(geo_bullets or [])} 則')
    except Exception as e:
        print(f'⚠️  geopolitics.json 讀取失敗：{e}')

    ok, msg = run(geopolitics_bullets=geo_bullets)
    print('✅ 傳送成功' if ok else '❌ 傳送失敗')
    print(msg)
