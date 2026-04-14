# 總經儀表板｜Macro Dashboard

> M&A Hunter 頻道專用。每日早晨自動更新，提供七大區塊的總體經濟數據速覽。

---

## 一、觸發與排程

**觸發詞**：`總經日報` / `macro dashboard`
**排程**：每日 **08:00** 與 **20:00**（台北時間）各執行一次
**輸出**：直接在對話輸出完整儀表板文字（含各區塊解讀）

---

## 二、輸出模板

```
📊 總經儀表板｜YYYY/MM/DD（週X）
（Telegram 每日 08:00 傳送儀表板截圖 + 完整網頁連結）
```

---

## 三、數據來源與取得策略

### 💱 匯率

**格式規則：統一以 XXX/USD（USD 在後）顯示，DXY 永遠第一，其餘按漲跌幅排序（降冪）**

| 顯示標籤 | YF 代號 | yfinance 原始 | 是否需倒轉 |
|---------|---------|--------------|-----------|
| DXY 美元指數 | DX-Y.NYB | — | 否（永遠第一） |
| EUR/USD | EURUSD=X | EUR/USD ✓ | 否 |
| GBP/USD | GBPUSD=X | GBP/USD ✓ | 否 |
| AUD/USD | AUDUSD=X | AUD/USD ✓ | 否 |
| TWD/USD | TWD=X | USD/TWD → 倒轉 | 是（price = 1/p，change% 取反） |
| JPY/USD | JPY=X | USD/JPY → 倒轉 | 是 |
| CNY/USD | CNY=X | USD/CNY → 倒轉 | 是 |
| KRW/USD | KRW=X | USD/KRW → 倒轉 | 是 |

### 📐 殖利率曲線（新）

| 數據 | 代號 | 來源 |
|------|------|------|
| US 2Y | ^IRX | yfinance |
| US 5Y | ^FVX | yfinance |
| US 10Y | ^TNX | yfinance |
| US 30Y | ^TYX | yfinance |

### 🏦 央行利率（新，hardcoded）

| 央行 | 利率 | 備註 |
|------|------|------|
| Fed | 4.25–4.50% | 需手動更新 |
| ECB | 2.50% | 需手動更新 |
| BOJ | 0.50% | 需手動更新 |
| CBC | 2.00% | 需手動更新 |

### 📈 全球股市

| 指數 | 代號 | 來源 |
|------|------|------|
| S&P 500 | ^GSPC | yfinance |
| Nasdaq | ^IXIC | yfinance |
| FTSE 100 | ^FTSE | yfinance |
| DAX | ^GDAXI | yfinance |
| STOXX 600 | ^STOXX | yfinance |
| 日經 225 | ^N225 | yfinance |
| KOSPI | ^KS11 | yfinance |
| 恆生指數 | ^HSI | yfinance |
| 上證綜合 | 000001.SS | yfinance |
| Nifty 50 | ^NSEI | yfinance |
| 加權指數（TAIEX） | ^TWII | yfinance |
| VIX | ^VIX | yfinance（移至市場情緒區塊） |

### 🛢️ 原物料

| 商品 | 代號 | 來源 |
|------|------|------|
| WTI 原油 | CL=F | yfinance |
| 布蘭特原油 | BZ=F | yfinance |
| 天然氣 | NG=F | yfinance |
| 黃金 | GC=F | yfinance |
| 白銀 | SI=F | yfinance |
| 銅 | HG=F | yfinance |

### 🪙 加密貨幣（CoinGecko，yfinance 不穩定）

**API**：`https://api.coingecko.com/api/v3/simple/price?ids=bitcoin,ethereum,solana&vs_currencies=usd&include_24hr_change=true`

| 幣種 | CoinGecko id |
|------|-------------|
| BTC | bitcoin |
| ETH | ethereum |
| SOL | solana |

**錯誤處理**：若 CoinGecko 失敗（timeout/429），fallback 至 `https://api.binance.com/api/v3/ticker/24hr?symbol=BTCUSDT`（ETHUSDT/SOLUSDT）

### 😱 市場情緒（VIX + Fear & Greed 並排）

**取得方式**：
- VIX：yfinance `^VIX`（price + 1d change%）
- Fear & Greed：`https://api.alternative.me/fng/?limit=2`（取代 CNN，無反爬限制）
  - ⚠️ CNN API 回傳 HTTP 418，已永久棄用

**輸出格式（左右並排）**：
```
😱 市場情緒
┌──────────────────────────┬──────────────────────────┐
│  📊 VIX 恐慌指數          │  🧠 CNN Fear & Greed      │
│  當前值：XX.XX            │  當前值：XX（Extreme Fear）│
│  漲跌：+X.XX（+X.X%）    │  漲跌：前值 XX → 今 XX   │
│  解讀：市場情緒XX         │  解讀：投資人情緒XX       │
└──────────────────────────┴──────────────────────────┘
```

**VIX 解讀對照表**：
| 值域 | 市場情緒 |
|------|---------|
| < 15 | 極度樂觀，市場平靜 |
| 15–20 | 正常波動，市場穩定 |
| 20–25 | 輕度不安，需留意 |
| 25–30 | 市場緊張，風險升高 |
| 30–40 | 明顯恐慌，謹慎操作 |
| > 40 | 極度恐慌，歷史性高點 |

**CNN Fear & Greed 解讀對照表**：
| 值域 | 情緒標籤 | 含義 |
|------|---------|------|
| 0–24 | Extreme Fear（極度恐懼） | 市場超賣，潛在買點 |
| 25–44 | Fear（恐懼） | 投資人保守，可逢低佈局 |
| 45–55 | Neutral（中性） | 市場平衡，觀望為主 |
| 56–74 | Greed（貪婪） | 情緒偏多，注意過熱 |
| 75–100 | Extreme Greed（極度貪婪） | 市場超買，注意修正風險 |

### 🏦 台灣三大法人買賣超（最新交易日）

**API**：`https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=json&dayDate={YYYYMMDD}&type=day`
- ⚠️ T86 改用 BFI82U（金額彙總），T86 僅回傳個股明細，合計列格式已變更

**取得策略**：
1. 先嘗試今日日期（`dayDate=今日YYYYMMDD`）
2. 若回傳 `stat != "OK"` 或資料為空 → 退一個交易日重試（最多退 7 個交易日）
3. 取回後標注實際資料日期（非執行日期）

**欄位對應**（`買賣差額` = index 3）：
- 外資 = `外資及陸資(不含外資自營商)` + `外資自營商`
- 投信 = `投信`
- 自營商 = `自營商(自行買賣)` + `自營商(避險)`
- 合計 = `合計`

**輸出欄位**：外資、投信、自營商 買賣差額（億元，+為買超、-為賣超），加總合計

### 📅 本週經濟數據（已公布 + 即將公布）

**API**：`https://nfs.faireconomy.media/ff_calendar_thisweek.json`

**篩選條件**：
- `impact = "High"` 僅顯示高重要性
- `currency` 僅取：**USD、TWD**（其他貨幣一律不顯示）
- TradingView API 請求參數：`countries=US,TW`

**顯示規則**：
- **已公布**（`actual` 不為空）：顯示實際值，評估 ▲（超預期）/ ▼（低於預期）/ ＝（符合預期），並附一句評語（如：「優於預期，對美元偏正面」）
- **未公布**（`actual` 為空）：標注「待公布」，顯示預期值與前值

**排版格式**：
```
📅 本週重要經濟數據
─────────────────────────────────────────────
✅ MM/DD HH:MM（台北） ｜ USD ｜ 事件名稱
   預期：X.X%　前值：X.X%　實際：X.X% ▲ 超預期
   評估：優於預期，有利美元/市場偏正向

⏳ MM/DD HH:MM（台北） ｜ USD ｜ 事件名稱（待公布）
   預期：X.X%　前值：X.X%
─────────────────────────────────────────────
```
- 若本週尚無任何高重要性事件：顯示「本週目前尚無重大數據」

### 🌍 國際局勢（每次執行強制更新，搜尋 12 小時視窗）

**時間視窗規則**：
- 執行時間為早上 08:00 → 搜尋「昨日 20:00 至今日 08:00」的新聞
- 執行時間為晚上 20:00 → 搜尋「今日 08:00 至今日 20:00」的新聞
- 計算方式：`搜尋起點 = 執行時間 − 12小時`，`搜尋終點 = 執行時間`

**搜尋策略**（每次執行必做，禁止使用快取）：
1. WebSearch `geopolitical military sanction trade war {YYYY/MM/DD}`（英文）
2. WebSearch `國際局勢 戰爭 制裁 貿易摩擦 {YYYY年M月D日} 最新`（中文）
3. 視情況補搜當日最重要事件的詳情（第三次 WebSearch）
4. 來源偏好：Reuters、BBC、鉅亨網國際、工商時報、聯合新聞網、ETtoday

**篩選標準**：
- 僅列影響金融市場的重大地緣事件（戰爭、制裁、貿易摩擦、軍事行動）
- 最多 3–5 則，每則：標題 + 一句話影響評估
- 若視窗內無重大事件：標記「本時段無重大地緣政治變動」
- **必須包含台灣板塊**（兩岸動態、台海局勢、美台關係）

**⚠️ 必須在執行 Runner 前完成（關鍵流程）：**
1. Claude 完成 WebSearch 後，將 bullets 寫入 `geopolitics.json`（格式見下）
2. **再執行** `python3 macro_dashboard_runner.py`
3. Runner 自動執行 `_push_geopolitics_json()`：git commit + push `geopolitics.json` → Render 即時更新網頁儀表板
4. Runner 自動讀取 `geopolitics.json` → 帶入 Telegram 訊息的 🌍 國際局勢區塊
5. ⛔ 若跳過步驟 1 直接執行 runner，國際局勢區塊將為空白

**geopolitics.json 雙格式規範**：
```json
{
  "updated": "YYYY/MM/DD",
  "bullets": [
    "🇺🇸🇮🇷 事件摘要一句話（Telegram 用）",
    "💰 事件摘要一句話",
    "💥 事件摘要一句話",
    "🌐 事件摘要一句話"
  ],
  "categories": [
    {
      "icon": "⚔️",
      "title": "分類名稱",
      "items": [
        {
          "icon": "🇮🇷",
          "title": "事件標題",
          "desc": "詳細描述（網頁版用）"
        }
      ]
    }
  ]
}
```
- `bullets`：Telegram runner 使用，每則一行，最多 5 則
- `categories`：網頁版儀表板使用，支援多層分類與詳細描述
- 兩者每次執行都必須同步更新

---

## 四、格式規則

1. 漲跌幅一律附 `+` 或 `-` 符號與百分比
2. 利率變動用 `bps`（基點）表示
3. 價格美元用 `$`，人民幣用 `¥`，台幣用 `NT$`
4. 週末 / 休市日標注「（休市，顯示上一交易日數據）」
5. 國際局勢用 bullet point，簡潔一句話
6. 央行利率為 hardcoded，需手動更新 `data_fetcher.py` 的 `CENTRAL_BANK_RATES`
7. Telegram 以截圖（PNG）傳送，附完整網頁連結作為 caption
8. 截圖工具：Playwright headless Chromium，1440×900 viewport，full_page=True

### 🕐 更新時間戳記規則

**每個數據區塊末尾必須標注更新時間，格式統一為：**

```
更新時間：YYYY/MM/DD HH:MM（台北時間）
```

**各區塊更新時間來源：**

| 區塊 | 更新時間來源 |
|------|------------|
| 💱 匯率 | yfinance 回傳的 `regularMarketTime`，轉換為台北時間 |
| 📐 殖利率曲線 | yfinance 回傳的 `regularMarketTime`，轉換為台北時間 |
| 🏦 央行利率 | 固定顯示「手動更新於 YYYY/MM/DD」（hardcoded，需人工維護） |
| 📈 全球股市 | yfinance 回傳的 `regularMarketTime`，轉換為台北時間（各指數可能不同，取最新一筆） |
| 🛢️ 原物料與加密貨幣 | yfinance 回傳的 `regularMarketTime`，轉換為台北時間 |
| 😱 CNN Fear & Greed | API 回傳時間戳，轉換為台北時間 |
| 📅 本週經濟日曆 | faireconomy API 查詢時間，顯示「資料截至 YYYY/MM/DD HH:MM（台北）」 |
| 🏦 台灣三大法人 | TWSE 回傳日期（`date` 欄位），格式轉換為台北時間 |
| 🌍 國際局勢 | 顯示 WebSearch 執行時間，格式「搜尋於 YYYY/MM/DD HH:MM（台北）」 |

**顯示規則：**
- 若取得時間失敗 → 顯示「更新時間：N/A」
- 休市商品 → 顯示「（休市）上次更新：YYYY/MM/DD HH:MM」
- 整體儀表板頁首加一行：`📊 資料擷取時間：YYYY/MM/DD HH:MM（台北時間）`

---

## 五、排程設定

- **執行頻率**：每日 08:00 台北時間
- **假日處理**：週六日仍執行，但標注休市數據
- **輸出位置**：Telegram Bot 推送（HTML parse mode + inline button）
- **腳本路徑**：`~/Desktop/Claude/Youtuber/macro-dashboard/macro_dashboard_runner.py`
- **推送方式**：`run(geopolitics_bullets=[...])` — 國際局勢由 Claude WebSearch 層傳入

## 六、推送設定

| 參數 | 值 |
|------|-----|
| 平台 | Telegram |
| Bot Token | `8743919766:AAG6z6YPW7Gqt7rF2KY2xC9mvbm2Ge31tjQ` |
| Chat ID | `2117347781` |
| Dashboard URL | `https://ma-hunter-macro-dashboard.onrender.com` |
| LINE | ⚠️ LINE Notify 已停服；Messaging API 需 Official Account，暫不使用 |

---

*最後更新：2026-04-14（v6 已優化 — 匯率統一 XXX/USD 格式並按漲跌幅排序、本週數據僅抓 USD+TWD、國際局勢必含台灣板塊、geopolitics.json 執行時自動 git push 至 Render）*
