# Financial_Analysis.md｜財務分析工作規範

> 製作財務文件時載入。涵蓋檔案命名、目錄結構、配色規範、API端點與交付物清單。

---

## 一、專案背景

本工作區專門處理**企業財務研究與投資分析**，涵蓋：

- **上市公司**：台灣證券交易所、櫃買中心掛牌公司（含台股，未來可擴展至港股、美股）
- **未上市公司**：私募股權目標、Pre-IPO 企業、家族企業
- **分析目的**：投資決策支援、客戶簡報 Demo、盡職調查、股票研究

> ⚠️ 每次分析的目標公司可能是全新未知的，不預設任何公司背景。
> 以使用者提供的數據為主要依據。

---

## 二、字型規範

- 所有中文輸出文件（DOCX / PPTX / XLSX）使用 `PingFang TC` 字型
- 英文備用字型：`Arial` 或 `Calibri`

---

## 三、檔案命名與目錄結構

### 命名格式
```
股票代號TT 公司名稱_檔案類型.副檔名
```

| 檔案類型 | 範例 |
|---------|------|
| 研究報告 | `8103TT 瀚荃股份_研究報告.docx` |
| 財務建模 | `8103TT 瀚荃股份_財務建模.xlsx` |
| 同業比較 | `8103TT 瀚荃股份_同業比較.xlsx` |
| 公司簡介 | `8103TT 瀚荃股份_公司簡介.pptx` |
| 盡職調查 | `8103TT 瀚荃股份_盡職調查數據包.xlsx` |
| 盡調清單 | `8103TT 瀚荃股份_盡調清單.xlsx` |
| 盈餘分析 | `8103TT 瀚荃股份_盈餘分析報告.docx` |

> 未上市公司無股票代號時，使用公司縮寫：`私募_XX公司_盡職調查.xlsx`

### 目錄結構
```
/Users/jeffhsu/Desktop/Claude/
│
├── CLAUDE.md                         ← 角色設定與觸發規則（自動載入）
├── Financial_Analysis.md             ← 本文件（製作財務文件時載入）
├── Company_Tracking.md               ← 追蹤清單規則（研究公司時載入）
├── Company Overview_YYYYMMDD.xlsx    ← 上市櫃公司追蹤清單
│
├── 8103TT 瀚荃股份/
│   ├── 8103TT 瀚荃股份_研究報告.docx
│   ├── 8103TT 瀚荃股份_財務建模.xlsx
│   ├── 8103TT 瀚荃股份_同業比較.xlsx
│   ├── 8103TT 瀚荃股份_公司簡介.pptx
│   ├── 8103TT 瀚荃股份_盡職調查數據包.xlsx
│   ├── 8103TT 瀚荃股份_盡調清單.xlsx
│   ├── 8103TT 瀚荃股份_盈餘分析報告.docx
│   ├── charts/
│   ├── scripts/
│   └── source_data/
│
└── _共用工具/
```

> 每次為新公司建立文件前，先確認或建立對應的公司資料夾。

---

## 四、色彩與設計規範

| 用途 | 色碼 |
|------|------|
| 主色（深藍標題） | `#1F4E79` |
| 副色（淺藍表頭） | `#D9E1F2` |
| 強調色（金色亮點） | `#C9A84C` / `#FFF2CC` |
| 正面指標（綠色） | `#E2EFDA` / `#375623` |
| 負面指標（紅色） | `#FCE4D6` / `#C00000` |
| 淺灰統計列 | `#F2F2F2` |

---

## 五、財務數據處理原則

1. **數據優先順序**：使用者提供的資料 > API即時數據 > 公開財報 > 估算值
2. **單位標示**：台股預設單位為新台幣億元（TWD 100M），需在文件頂部明確標示
3. **估算值標記**：同業比較或預測數據需標注「估計值」或「*」
4. **公式原則**：Excel 中所有計算欄位必須使用公式，禁止直接輸入計算結果
5. **藍黑慣例**：Excel 中手動輸入數字用藍色字體，公式計算結果用黑色字體

---

## 六、數據來源 API

---

### 6.1 五九一財報 API（five91.onrender.com）— 主要財務數據

> 涵蓋台灣所有上市櫃公司，每日更新。股價、財務指標、估值倍數全部串接此處。

**核心原則：先鎖定股票代號，再查詢。絕對不要先載入全部公司再篩選。**

#### ✅ 確認可用端點

| 用途 | 端點 | 說明 |
|------|------|------|
| **單股所有財務指標**（主要使用） | `GET https://five91.onrender.com/api/metrics?stock_id={代號}` | 回傳 51 欄位完整 JSON |
| 全市場所有股票指標 | `GET https://five91.onrender.com/api/metrics` | 勿輕易呼叫（高 token） |
| 產業分類清單 | `GET https://five91.onrender.com/api/valuation/categories` | 27 個產業 |
| 估值計算 | `POST https://five91.onrender.com/api/valuation/calculate` | 傳入 JSON body |
| 個股概覽頁 | `GET https://five91.onrender.com/stock/{代號}` | 含 EV、EV/EBITDA |
| 損益表（6季） | `GET https://five91.onrender.com/stock/{代號}?tab=income_statement` | 季度損益明細 |
| 資產負債表（6季） | `GET https://five91.onrender.com/stock/{代號}?tab=balance_sheet` | 季度資產負債 |
| 現金流量表（6季） | `GET https://five91.onrender.com/stock/{代號}?tab=cash_flow` | 含 D&A |

#### `/api/metrics?stock_id={代號}` 回傳欄位（51欄）

```
id, name, category, quarter,
revenue, gross_margin, operating_margin, net_margin,
net_income, net_income_ttm, revenue_ttm, ebitda, eps,
debt_ratio, current_ratio, quick_ratio, ocf, free_cf,
share_capital, preferred_stock, cash, financial_assets_current,
minority_interest, notes_payable, short_term_debt, long_term_debt, equity,
cash_dividend_ps_ttm, roe_12q, roa_12q,
gross_margin_12q, operating_margin_12q, net_margin_12q,
price, change, market_cap, ev, ev_ebitda,
ps, pb, p_ebitda, cash_ratio, ebitda_yield, dividend_yield, pe
```

#### 欄位與 Company_Tracking.md 欄位對應

| Excel 欄 | 欄位名稱 | five91 欄位 |
|----------|---------|------------|
| D | 2025 全年營收 | `revenue_ttm`（TTM）或取全年加總 |
| E | 2025 全年毛利率 | `gross_margin_12q` |
| F | 2025 全年營益率 | `operating_margin_12q` |
| G | 2025 全年淨利率 | `net_margin_12q` |
| H | 2025 全年稅後淨利 | `net_income_ttm` |
| I | 最新資本額 | `share_capital` |
| J | 2025 全年 EPS | `eps`（最新季度；TTM需累加） |
| L | 產業分類 | `category` |
| — | 最新股價 | `price` |
| — | 市值 | `market_cap` |
| — | EV/EBITDA | `ev_ebitda` |

#### 查詢策略（節省 Token）

```
使用者提供股票代號（如 4906）
    ↓
需要股價/估值/利潤率 → GET /api/metrics?stock_id=4906   (~500 tokens)
需要季度損益明細     → GET /stock/4906?tab=income_statement
需要資產負債詳情     → GET /stock/4906?tab=balance_sheet
需要 D&A / FCF      → GET /stock/4906?tab=cash_flow
```

---

### 6.2 公開資訊觀測站 MOPS — 法說會簡報 PDF（產品別佔比來源）

> 用途：取得 K 欄「產品別佔比」，來源為公司法說會簡報 PDF。

#### 查詢步驟

**Step 1：取得法說會列表**

```
POST https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1
Content-Type: application/x-www-form-urlencoded

co_id={股票代號}&step=1&firstin=1&off=1&keyword4=&code1=&TYPEK=sii
```

回傳 HTML 表格，解析找到最新一筆的中文 PDF 檔名（格式：`{代號}{YYYYMMDD}M{序號}.pdf`）。

**Step 2：下載 PDF**

```
https://mopsov.twse.com.tw/server-java/FileDownLoad?step=1&fileName={檔名}&filePath=/t100/
```

範例（4906 正文 2025/12/11 法說會）：
```
https://mopsov.twse.com.tw/server-java/FileDownLoad?step=1&fileName=490620251211M001.pdf&filePath=/t100/
```

**Step 3：讀取 PDF，擷取產品別佔比**

用 `pypdf` 或 `pdfplumber` 讀取，搜尋關鍵字「產品」「佔比」「Revenue Mix」等段落，整理格式：
```
產品A XX%；產品B XX%；產品C XX%
```

若 PDF 無法取得或無產品佔比資訊，填入「待補充」。

#### MOPS 查詢網址

- 法說會查詢頁：`https://mopsov.twse.com.tw/mops/web/t100sb07_1`
- AJAX 端點：`https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1`

---

## 七、標準交付物清單

| # | 任務 | Skill | 輸出格式 |
|---|------|-------|---------|
| 1 | 公司簡介（Blind Teaser） | `investment-banking:teaser` | `.pptx` |
| 2 | 財務建模（DCF） | `financial-analysis:dcf` | `.xlsx` |
| 3 | 股票研究報告 | `equity-research:initiate` | `.docx` |
| 4 | 同業比較分析 | `financial-analysis:comps-analysis` | `.xlsx` |
| 5 | 盡職調查數據包 | `investment-banking:datapack-builder` | `.xlsx` |
| 6 | 盡職調查清單 | `private-equity:dd-checklist` | `.xlsx` |
| 7 | 盈餘分析報告 | `equity-research:earnings` | `.docx` |
| 8 | LBO 模型 | `financial-analysis:lbo` | `.xlsx` |
| 9 | 投資委員會備忘錄 | `private-equity:ic-memo` | `.docx` |

---

## 八、已完成參考案例

### 8103.TW 瀚荃股份有限公司（2026年3月）

**公司簡介：** 台灣上市精密連接器製造商，市值 NT$66.9 億

**關鍵財務數據（FY2025）：**
- 營收：NT$33.55 億（YoY +4.5%）
- 毛利率：37.4%（YoY +1.4pp）
- EBITDA：NT$6.70 億（利潤率 20.0%）
- EPS：NT$3.92 | 淨現金：NT$19.60 億
- 估值：EV/EBITDA 6.14x（同業折價 37%）

**已交付文件（`8103TT 瀚荃股份/`）：**
`_公司簡介.pptx` / `_財務建模.xlsx` / `_研究報告.docx` / `_同業比較.xlsx` / `_盡職調查數據包.xlsx` / `_盡調清單.xlsx` / `_盈餘分析報告.docx`

---

## 九、數據來源優先順序（速查表）

| 需求 | 優先來源 | 備用 |
|------|---------|------|
| 股價、市值、EV/EBITDA | `five91 /api/metrics?stock_id=` | TWSE API |
| 毛利率、營益率、淨利率 | `five91 /api/metrics?stock_id=` (`_12q` 欄位) | WebSearch |
| EPS、資本額、稅後淨利 | `five91 /api/metrics?stock_id=` | 公開財報 |
| 季度損益明細 | `five91 /stock/{代號}?tab=income_statement` | — |
| 產品別佔比 | MOPS 法說會 PDF（見 6.2 節） | 公司官網 IR |
| 歷史收盤價 | TWSE `STOCK_DAY` API | — |
| 三大法人買賣超 | TWSE T86 API（`?date=YYYYMMDD`） | — |
| 近期新聞題材 | WebSearch | — |

---

*最後更新：2026-03-25*
