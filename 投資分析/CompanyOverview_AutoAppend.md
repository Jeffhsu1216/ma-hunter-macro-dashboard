---
name: 公司追蹤總表自動 Append 規範
description: 每次查詢台灣上市櫃公司（短線分析、月初選股、日常詢問）皆自動 append 至 Company Overview xlsx
type: project
---

# 公司追蹤總表自動 Append 規範

> **觸發時機（四種）：**
> 1. `短線分析 {代號}` — 執行短線技術/基本面評估（單筆，腳本自動 append）
> 2. `月初選股 YYYYMM` — 批次分析選股清單（40 支，比對去重）
> 3. `{代號} {公司} {稱謂} 感謝信` — 法說會感謝信（Phase 5 自動 append）
> 4. 日常詢問 — `研究 XXXX` / `查 XXXX` / `幫我看 XXXX`，取得財務資料時
>
> **上述任一觸發後，無需詢問使用者，直接執行 append。**

---

## 一、檔案資訊

- **存放路徑**：`~/Desktop/Stock Analysis/`
- **命名格式**：`Company Overview_{YYYYMMDD}.xlsx`（YYYYMMDD = 執行當天日期）
- **工作表**：`公司追蹤清單`
- **更新方式**：
  1. 讀取現有最新日期檔案（`glob` 取日期最大者）
  2. 若該代號已存在 → 更新該列
  3. 若不存在 → 新增一列，序號自動遞增
  4. 以今日日期另存為新檔名（若日期不同）

---

## 二、欄位規格

> **與 Company_Tracking.md 完全一致**，欄位不可自行增減。

| 欄 | 欄位名稱 | 說明 | 資料來源 |
|----|---------|------|---------|
| A | 序號 | 由 1 開始接續排列 | 自動 |
| B | 股票代號 | 台灣股票代號（如 4906） | 使用者輸入 |
| C | 股票名稱 | 公司中文名稱 | API / 搜尋 |
| D | 2025 全年營收（億元） | 單位：新台幣億元 | five91 API |
| E | 2025 全年毛利率（%） | 百分比 | five91 API |
| F | 2025 全年營益率（%） | 百分比 | five91 API |
| G | 2025 全年淨利率（%） | 百分比 | five91 API |
| H | 2025 全年稅後淨利（億元） | 單位：新台幣億元 | five91 API |
| I | 最新資本額（億元） | 單位：新台幣億元 | five91 API |
| J | 2025 全年 EPS（元） | 新台幣元 | five91 API |
| K | 產品別佔比 | 各產品線營收佔比；來源：MOPS 法說會簡報；無法取得填「待補充」 | MOPS |
| L | 產業分類 | 依 TWSE / five91 分類 | five91 API |
| M | 主要題材 | 投資主題摘要（50字內） | WebSearch 彙整 |

---

## 三、執行規則

### 3.1 讀取現有檔案

```python
from pathlib import Path
import datetime, openpyxl

folder = Path.home() / 'Desktop/Stock Analysis'
files = sorted(folder.glob('Company Overview_*.xlsx'))
# 取日期最大的檔案（唯一一份，無論原始日期為何）
latest = files[-1] if files else None
```

### 3.2 Upsert 邏輯

```python
# 以 B 欄（股票代號）為 key
# 掃描現有資料列，找到代號 → 更新；找不到 → 新增末列
```

### 3.3 存檔（覆蓋更新，永遠只保留一份）

```python
import os
today = datetime.date.today().strftime('%Y%m%d')
new_path = folder / f'Company Overview_{today}.xlsx'
wb.save(new_path)
# 若原檔日期與今日不同 → 刪除舊檔，只留今天的
if latest and latest != new_path and latest.exists():
    os.remove(latest)
# ⚠️ 資料夾內永遠只有一份 Company Overview xlsx
# ⚠️ 檔名日期 = 最後更新日（今天）
```

### 3.4 新增列格式規範

```python
from openpyxl.styles import PatternFill, Alignment

no_fill = PatternFill(fill_type=None)          # 無底色
center = Alignment(horizontal='center', vertical='center', wrap_text=True)

for cell in new_row:
    cell.fill = no_fill
    cell.alignment = center
# ⚠️ 新增的資料列一律：無底色 + 全欄置中（含文字欄）
```

---

## 四、格式規範

| 元素 | 規格 |
|------|------|
| 標題列背景 | `#1F4E79`（深藍），白字，粗體 |
| 偶數列背景 | `#F2F2F2`（淺灰） |
| 奇數列背景 | 白色 |
| 字型 | PingFang TC（中文）/ Arial（英文） |
| 對齊 | 全欄 Middle Align + 文字 Left Align（D、G 欄數字 Right Align） |
| 欄寬 | 自動調整至內容寬度（`column_dimensions` auto） |

---

## 五、資料取得來源（完整串接）

| 欄位 | 來源 | 端點 / 方法 |
|------|------|------------|
| C 名稱、L 產業 | five91 API | `curl` → `five91.onrender.com/api/metrics?stock_id={代號}` → `name`, `category` |
| D 營收 | five91 income_statement | `/stock/{代號}?tab=income_statement` 四季加總 |
| E 毛利率 | five91 income_statement | 四季毛利加總 ÷ 四季營收（非 12q rolling） |
| F 營益率 | five91 income_statement | 四季營業利益加總 ÷ 四季營收 |
| G 淨利率 | five91 income_statement | 四季稅後淨利加總 ÷ 四季營收 |
| H 稅後淨利 | five91 API | `net_income_ttm`（驗證用） |
| I 資本額 | 計算 | `market_cap ÷ price × 10`（億元） |
| J EPS | 計算 | `net_income_ttm ÷ shares`（shares = market_cap ÷ price） |
| K 產品別佔比 | MOPS 法說會 PDF | POST `ajax_t100sb07_1`（sii 優先，otc 備用），智慧掃描 |
| M 主要題材 | WebSearch | 近期新聞彙整 50 字內 |

> - 取得失敗填「N/A」，不中斷整體流程
> - 產品別佔比無法取得填「待補充」
> - 詳細 API 規格與 MOPS PDF 下載流程見 `Financial_Analysis.md` 第六節

---

## 六、效能規範（優化後標準）

> 實測基準：6435 大中 79s / ~3,200 tokens；6651 全宇昕 83s / ~3,500 tokens

### 6.1 四支 curl 並行（Phase 1，目標 ≤ 2s）

```bash
curl -s "https://five91.onrender.com/api/metrics?stock_id={代號}" > /tmp/{代號}_metrics.json &
curl -s "https://five91.onrender.com/stock/{代號}?tab=income_statement" > /tmp/{代號}_income.html &
curl -s -X POST "https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1" \
  -d "co_id={代號}&step=1&firstin=1&off=1&keyword4=&code1=&TYPEK=sii" > /tmp/{代號}_mops_sii.html &
curl -s -X POST "https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1" \
  -d "co_id={代號}&step=1&firstin=1&off=1&keyword4=&code1=&TYPEK=otc" > /tmp/{代號}_mops_otc.html &
wait
```

### 6.2 單一 Python 腳本處理（Phase 2，目標 ≤ 60s）

所有解析合併成一個腳本，避免多次工具呼叫 overhead：
1. 解析 metrics JSON → C、L 欄及備用數值
2. 解析 income_statement HTML → D、E、F、G、H、J 四季計算
3. 解析 MOPS HTML → PDF 檔名（sii 優先，otc fallback）
4. `urllib.request.urlretrieve` 下載 PDF
5. `pdfplumber` 智慧掃描 → K 欄
6. 儲存結果至 `/tmp/{代號}_result.json`

### 6.3 PDF 智慧掃描規則

```python
with pdfplumber.open(pdf_path) as pdf:
    for i in range(3, min(15, len(pdf.pages))):   # 從第4頁開始，跳過封面/目錄
        text = pdf.pages[i].extract_text() or ''
        if re.search(r'\d+\s*%', text) and any(
            k in text for k in ['應用','產品','市場','Revenue','佔比','比重','業務']
        ):
            lines = [l.strip() for l in text.split('\n') if '%' in l and re.search(r'\d', l)]
            if lines:
                product_k = '；'.join(lines[:6])
                break   # ← 找到即停，不再掃後面頁
```

- 第 4 頁起掃（index=3），跳過封面、目錄、公司沿革
- 找到含 `數字%` 且有應用/產品關鍵字的頁面即停
- 最多掃到第 15 頁；超過仍未命中填「待補充」

### 6.4 Phase 3：WebSearch + Excel 寫入

- WebSearch 1 次取主要題材（M 欄）
- openpyxl upsert + `PatternFill(fill_type=None)` + `Alignment(center)` + 存檔

---

## 七、注意事項

- **禁止留 N/A**：five91 未收錄的股票（如金控、小型股），必須 WebSearch 查詢補齊全部欄位
  - 金融業（金控、證券）毛利率/營益率填「金融業不適用」
  - five91 `/api/metrics` 有 `gross_margin`、`operating_margin`、`net_margin` 欄位，務必讀取
  - 名稱(C)絕對不可為空，five91 無資料時從 MANUAL_EPS 或 WebSearch 取得
- 月初選股批次 40 支可一次批量寫入，不須逐筆觸發
- 此檔案為歷史查詢紀錄，**禁止刪除任何已存在資料列**
- **禁止使用 WebFetch 取代 curl**：WebFetch 會過一層 AI 模型，浪費 token 且易遺漏指定股票

---

*建立日期：2026-04-01（v1）*
*最後更新：2026-04-07（v3 — 新增感謝信觸發、禁止留N/A規則、five91 margin 欄位確認、金融業特殊處理）*
