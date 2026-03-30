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

## 六、數據來源 API（five91.onrender.com）

> 使用者自建財報網站，涵蓋台灣所有上市櫃公司，每日更新。

### 核心原則：先鎖定股票代號，再查詢

**絕對不要**先載入全部公司再篩選（耗費大量 Token）。
正確做法：取得股票代號 → 直接用代號查詢對應端點。

### 端點列表

#### 個股完整財務資料（主要使用）

| 用途 | URL |
|------|-----|
| 個股概覽（EV、EV/EBITDA、股價、市值） | `GET https://five91.onrender.com/stock/{代號}` |
| 損益表（6季） | `GET https://five91.onrender.com/stock/{代號}?tab=income_statement` |
| 資產負債表（6季） | `GET https://five91.onrender.com/stock/{代號}?tab=balance_sheet` |
| 現金流量表（含D&A，6季） | `GET https://five91.onrender.com/stock/{代號}?tab=cash_flow` |

#### 全市場指標（次要使用）

| 方法 | 端點 | 說明 |
|------|------|------|
| GET | `/api/metrics` | 全部股票51個欄位，以股票代號為 key |
| GET | `/api/valuation/categories` | 27個產業分類清單 |
| POST | `/api/valuation/calculate` | 估值計算（傳入 JSON body） |

**`/api/metrics` 51個欄位：**
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

### 查詢策略（節省 Token）

```
使用者提供股票代號（如 6257）
    ↓
需要 EV/估值/股價  → GET /api/metrics → filter by id（~2,000 tokens）
需要季度損益      → GET /stock/6257?tab=income_statement
需要季度 D&A      → GET /stock/6257?tab=cash_flow
需要資產負債      → GET /stock/6257?tab=balance_sheet
```

> 每次查詢約 2,000–3,000 tokens（vs 全量掃描 100,000+ tokens，效率提升 30-50x）

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

## 九、未來 API 串接計劃

| 數據類型 | 計劃來源 | 用途 |
|---------|---------|------|
| 法說會資料 | MOPS 電子書 / 公司 IR 網站 | 管理層展望自動摘要 |
| 新聞監控 | Google News API / 財經新聞 RSS | 重大事件觸發報告更新 |
| 海外股票 | Alpha Vantage / Polygon.io | 港股、美股 ADR 分析擴展 |

---

*最後更新：2026-03-25*
