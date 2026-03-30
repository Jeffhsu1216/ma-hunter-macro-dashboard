# 總經儀表板｜Macro Dashboard

> M&A Hunter 頻道專用。每日早晨自動更新，提供七大區塊的總體經濟數據速覽。

---

## 一、觸發與排程

**觸發詞**：`總經日報` / `macro dashboard`
**排程**：每日早上 08:00（台北時間）自動執行
**輸出**：Telegram 截圖（PNG）+ 完整網頁連結

---

## 二、輸出模板

```
📊 總經儀表板｜YYYY/MM/DD（週X）
（Telegram 每日 08:00 傳送儀表板截圖 + 完整網頁連結）
```

---

## 三、數據來源與取得策略

### 💱 匯率

| 數據 | 代號 | 來源 |
|------|------|------|
| DXY 美元指數 | DX-Y.NYB | yfinance |
| USD/TWD | TWD=X | yfinance |
| USD/JPY | JPY=X | yfinance |
| USD/CNY | CNY=X | yfinance |
| EUR/USD | EURUSD=X | yfinance |
| GBP/USD | GBPUSD=X | yfinance |
| AUD/USD | AUDUSD=X | yfinance |
| USD/KRW | KRW=X | yfinance |

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
| VIX | ^VIX | yfinance |

### 🛢️ 原物料與加密貨幣

| 商品 | 代號 | 來源 |
|------|------|------|
| WTI 原油 | CL=F | yfinance |
| 布蘭特原油 | BZ=F | yfinance |
| 天然氣 | NG=F | yfinance |
| 黃金 | GC=F | yfinance |
| 白銀 | SI=F | yfinance |
| 銅 | HG=F | yfinance |
| BTC | BTC-USD | yfinance |
| ETH | ETH-USD | yfinance |

### 🧠 新增模組

| 模組 | 來源 URL |
|------|---------|
| CNN Fear & Greed | https://production.dataviz.cnn.io/index/fearandgreed/graphdata |
| 本週經濟日曆 | https://nfs.faireconomy.media/ff_calendar_thisweek.json（High impact，USD/CNY/EUR/JPY） |
| 台灣三大法人買賣超 | https://www.twse.com.tw/rwd/zh/fund/T86?response=json |

#### 📅 本週經濟日曆 — 處理規則

**篩選條件：**
- `impact = "High"` 僅顯示高重要性事件
- `currency` 僅取：USD、CNY、EUR、JPY、TWD
- **只顯示 `actual` 欄位不為空的事件**（即已公布數據）；未公布者不顯示

**時間轉換：**
- 原始時間為 UTC，一律轉換為 **台灣時間（UTC+8）** 顯示
- 格式：`MM/DD（週X）HH:MM（台北）`

**輸出欄位（每筆）：**

| 欄位 | 說明 |
|------|------|
| 時間 | 台灣時間，格式 MM/DD HH:MM |
| 貨幣 | 對應央行貨幣（USD / EUR / JPY / CNY / TWD） |
| 事件名稱 | 原文事件名，必要時附繁中說明 |
| 預期值 | `forecast` 欄位；若為空則顯示「—」 |
| 前值 | `previous` 欄位 |
| 實際值 | `actual` 欄位（**已公布才顯示**） |

**排版格式（文字輸出）：**
```
📅 本週重要經濟數據（已公布）
─────────────────────────────────────────────
MM/DD HH:MM（台北） ｜ 貨幣 ｜ 事件名稱
  預期：X.X%　前值：X.X%　實際：X.X% ▲/▼
─────────────────────────────────────────────
```
- 實際值 > 預期值 → 加 ▲（超預期）
- 實際值 < 預期值 → 加 ▼（低於預期）
- 無預期值則不顯示符號
- 若本週尚無已公布高重要性數據：顯示「本週目前尚無重大數據公布」

### 🌍 國際局勢

**搜尋策略**：
1. WebSearch `"geopolitical risk today {date}"` — 英文全球視角
2. WebSearch `"國際局勢 戰爭 衝突 {年月日}"` — 繁中視角
3. 來源偏好：Reuters、BBC、鉅亨網國際、工商時報國際

**篩選標準**：
- 僅列影響金融市場的重大地緣事件（戰爭、制裁、貿易摩擦、軍事行動）
- 最多 3-5 則，每則一句話摘要
- 若當日無重大事件：標記「今日無重大地緣政治變動」

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
- **輸出位置**：Telegram Bot 推送截圖 + 網頁連結

---

*最後更新：2026-03-30（v2 — 本週經濟日曆改為「已公布數據」台北時間顯示，新增全模組更新時間戳記）*
