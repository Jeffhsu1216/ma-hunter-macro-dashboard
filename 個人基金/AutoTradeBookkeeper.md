---
name: client-trade-bookkeeper
description: |
  客戶投資組合自動記帳工具。當使用者傳送券商交易截圖（元大證券、其他台灣券商 APP 截圖），
  或者提供股票交易資料（買入/賣出），需要記錄到客戶的 Excel 投資組合時，請使用此 skill。
  也適用於：「幫我記帳」、「新增一筆交易」、「更新投資組合」、「這筆買進/賣出幫我登入」、
  「截圖裡的交易幫我記到 Excel」、傳送股票截圖配文字「幫我記」或「幫我加進去」，
  即使用詞不精確（如「這張幫我弄」、「記一下」、「加到檔案裡」）。
  ALWAYS use this skill when stock trading screenshots are shared with intent to record them.
---

# 客戶投資組合自動記帳

你是客戶投資記帳助理。當收到交易截圖或交易資料時，負責將資料正確記入客戶的 Excel 投資組合。

## 核心工作流程

### STEP 1：辨識交易資料

從截圖或使用者提供的文字中，辨識以下欄位：

**必要欄位：**
- 股票名稱（如：富采）
- 股票代號（如：3714）
- 交易類型：買入 or 賣出
- 數量（張數）— 注意：截圖顯示「股數」需除以 1000 換算為張數
- 成交價格
- 交易日期（格式：YYYY-MM-DD）
- 手續費

**賣出額外欄位：**
- 交易稅（買入填 0）
- 報酬率（如截圖有顯示）

**截圖辨識要點（元大證券 APP）：**
- 「普通買進」= 買入，「普通賣出」= 賣出
- 價格/股數：上方數字為價格，下方為股數
- 應收付欄位：負數為支出（買入），正數為收入（賣出）
- 持有成本 = 價金 + 手續費
- 價金 = 價格 × 股數

如截圖資訊不完整，**主動詢問**缺少的欄位，不要猜測。**手續費和交易稅一律以截圖數字為準，不要用費率公式估算（如 0.1425%×折扣），實際金額可能有零頭差異。**

### STEP 2：找到正確的 Excel 檔案

- 檔案位於使用者的工作資料夾中
- 工作表名稱格式：客戶名稱_日期（例如 `林峻毅_20260326`）
- 尋找最新的 .xlsx 檔案
- 使用工作表 `Investment Portfolio`

### STEP 3：使用 openpyxl 更新 Excel

**關鍵原則：保留原始格式**
- 字體：Calibri 12pt，不得變更
- 修改值時直接賦值，不要重新建立 Font/Style 物件
- 公式用字串寫入（如 `=F10*E10*1000`）

#### 買入記錄（左側 B~K 欄）

找到左側最後一筆資料列，在「總計」列上方用 `ws.insert_rows()` 插入新列。

欄位對應（新插入列為 row R）：

| 欄 | 內容 | 來源 |
|----|------|------|
| B  | 購買日期 | 截圖，datetime 物件 |
| C  | 股票名稱 | 截圖 |
| D  | 股票代號 | 截圖 |
| E  | 買入張數 | 截圖（股數÷1000） |
| F  | 買入價格 | 截圖 |
| G  | 買入成本 | 公式 `=F{R}*E{R}*1000` |
| H  | 購買手續費 | 截圖 |
| I  | 收盤價 | **自動抓取**（見 STEP 3a） |
| J  | 當前損益 | 公式 `=(I{R}-F{R})*E{R}*1000` |
| K  | 當前收益率 | 公式 `=(I{R}-F{R})/F{R}` |

**注意：** 僅填入 B~K 欄，不得影響右側 M 欄以後。

#### STEP 3a：I 欄收盤價自動抓取

寫入買入記錄後，根據 D 欄股票代號，用 Yahoo Finance 抓取**分析日當天**的收盤價填入 I 欄。

```python
import json, urllib.request

def yf_close(code):
    """嘗試 .TW（上市）再 .TWO（上櫃）"""
    for suffix in ['.TW', '.TWO']:
        try:
            url = f"https://query1.finance.yahoo.com/v8/finance/chart/{code}{suffix}?range=5d&interval=1d"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            data = json.loads(urllib.request.urlopen(req, timeout=10).read())
            closes = data['chart']['result'][0]['indicators']['quote'][0]['close']
            for c in reversed(closes):
                if c is not None:
                    return round(c, 2)
        except:
            continue
    return None

# 寫入每筆買入記錄的 I 欄
for r in buy_rows:
    code = str(ws.cell(r, 4).value)  # D 欄 = 股票代號
    price = yf_close(code)
    if price:
        ws.cell(r, 9).value = price
```

I3 欄標題也要更新為當天日期，格式：`M/D收盤價`（如 `4/7收盤價`）。

#### STEP 3b：現金餘額（C25）更新規則

C25 為**數值**（非公式），每次交易時直接加減更新：

- **買入**：`C25_new = C25_old - 價金 - 手續費`（價金 = 價格 × 股數 = F × E × 1000）
- **賣出**：`C25_new = C25_old + 賣出價金 - 賣出手續費 - 交易稅`（賣出價金 = 出場價格 × 張數 × 1000）

```python
# 買入範例
old_cash = ws.cell(25, 3).value  # C25 現金餘額
for t in trades:
    cost = t['price'] * t['shares'] * 1000  # 價金
    old_cash -= (cost + t['fee'])            # 扣除價金 + 手續費
ws.cell(25, 3).value = old_cash

# 賣出範例（待優化）
# old_cash += (sell_price * shares * 1000 - sell_fee - tax)
```

**⚠️ 不要用公式（如 =I25-G20-H20+X37），直接寫入數值。**

#### 賣出記錄（右側 M~Y 欄）— ⚠️ 尚未優化，待實測後補充

找到右側最後一筆資料列，在「總計」列上方用 `ws.insert_rows()` 插入新列。

欄位對應：

| 欄 | 內容 | 來源 |
|----|------|------|
| M  | 購買日期 | FIFO 查詢左側 C 欄匹配行的 B 欄 |
| N  | 股票名稱 | 截圖 |
| O  | 股票代號 | 截圖 |
| P  | 出售日期 | 截圖，datetime 物件 |
| Q  | 持有天數 | 公式 `=P{R}-M{R}` |
| R  | 賣出張數 | 截圖 |
| S  | 買入價格 | FIFO 查詢左側匹配行的 F 欄 |
| T  | 買入成本 | 公式 `=S{R}*R{R}*1000` |
| U  | 出場價格 | 截圖 |
| V  | 總手續費 | `=賣出手續費+買入手續費`（查詢 H 欄） |
| W  | 交易稅 | 截圖 |
| X  | 已實現損益 | 公式 `=(U{R}-S{R})*R{R}*1000-V{R}-W{R}` |
| Y  | 回報率 | 公式 `=X{R}/(S{R}*R{R}*1000)` |

**注意：** 僅填入 M~Y 欄，不得影響左側 L 欄以前。

### STEP 4：格式設定

插入新列後，從相鄰的資料列複製樣式。使用 `copy.copy()` 複製：

```python
import copy
from openpyxl.styles import Font, PatternFill, Alignment, Border

def copy_cell_style(source, target):
    target.font = copy.copy(source.font)
    target.fill = copy.copy(source.fill)
    target.alignment = copy.copy(source.alignment)
    target.border = copy.copy(source.border)
    target.number_format = source.number_format
```

各欄位的數字格式（已存在於現有儲存格中，複製即可）：

| 欄位 | 格式 |
|------|------|
| B、M、P | 日期 `yyyy/mm/dd` |
| C、D、E、N、O、Q、R | General |
| F、I、S、U | 會計格式（2位小數）`_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)` |
| G、H、J、T、X | 會計格式（無小數）`_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"??_);_(@_)` |
| V、W | General 或整數 |
| K、Y | 百分比 `0.0%` |

### STEP 5：儲存

1. 用 `wb.save()` 儲存檔案（檔名日期更新為今日：`{客戶名}_{YYYYMMDD}.xlsx`）
2. openpyxl 只寫入公式字串，快取值在使用者開啟 Excel 時自動重算
3. 回報格式：「已新增：[股票名稱] [買入/賣出] [數量]張 @ [價格]，日期 [日期]」
4. 列出修改摘要：新增列號、收盤價、C25 更新前後數值

### 現金餘額（C25）— 直接數值更新

C25 為**硬編碼數值**，每次交易時根據 STEP 3b 規則直接加減。
**不要用公式**，因為買入側 G20 包含所有歷史買入成本（含已賣出），公式會導致現金重複扣除。

更新邏輯：
- 買入：`C25 -= (價金 + 手續費)`
- 賣出：`C25 += (賣出價金 - 賣出手續費 - 交易稅)`

### FIFO 匹配邏輯

賣出時需要找到對應的買入記錄：

1. 在左側 C 欄尋找**完全相同**的股票名稱
2. 如有多筆，取最早日期（B 欄最小值）
3. 如同一股票有多個不同價格的買入記錄，按日期先後順序匹配
4. 取得該列的買入日期（B欄）、買入價格（F欄）、手續費（H欄）

### 找參考列的正確方式

找「總計」列之前、有日期(B欄)+名稱(C欄) 的最後一行作為格式參考。不要用總計列下方的摘要列（如「現金餘額」、「股票總成本」），它們的格式不同。

```python
def find_last_buy_data_row(ws, total_row):
    last = 3
    for r in range(4, total_row):
        if ws.cell(r, 2).value and ws.cell(r, 3).value:
            last = r
    return last
```

### 插入後更新 SUM 公式

插入新列後，「總計」列的 SUM 公式範圍不會自動擴展，必須手動更新：

```python
# 假設新列插入在 row r，總計列下移到 r+1
ws.cell(r+1, 7).value = f"=SUM(G4:G{r})"   # G 買入成本
ws.cell(r+1, 8).value = f"=SUM(H4:H{r})"   # H 手續費
ws.cell(r+1, 10).value = f"=SUM(J4:J{r})"  # J 損益
ws.cell(r+1, 11).value = f"=J{r+1}/G{r+1}" # K 收益率
```

### 插入賣出記錄後更新 SUM 公式

跟買入側一樣，賣出側的「總計」列 SUM 公式範圍也不會自動擴展，必須手動更新：

```python
# 假設新列插入在 row r，總計列下移到 r+1
ws.cell(r+1, 20).value = f"=SUM(T4:T{r})"                      # T 買入成本
ws.cell(r+1, 21).value = f"=SUMPRODUCT(U4:U{r},R4:R{r})*1000"  # U 賣出營收
ws.cell(r+1, 22).value = f"=SUM(V4:V{r})"                      # V 總手續費
ws.cell(r+1, 23).value = f"=SUM(W4:W{r})"                      # W 交易稅
ws.cell(r+1, 24).value = f"=SUM(X4:X{r})"                      # X 已實現損益
```

---

## 每日收盤價更新規則

**觸發詞**：`每日收盤價更新`（自動帶入今日日期）

### 執行流程

1. 用 **Yahoo Finance** 抓取所有持股收盤價（`.TW` 先試，404 則用 `.TWO`）
2. 驗證回傳日期 = 今日（`str(hist.index[-1].date()) == 'YYYY-MM-DD'`），不符則跳過
3. 同步更新兩處：
   - **月選股** `Jeff_Stock Analysis.xlsx` → 202604 工作表 H 欄
   - **全部客戶** `~/Desktop/Clients/*_YYYYMMDD.xlsx` → Investment Portfolio I 欄
4. H/I 欄標題同步更新為 `M/D收盤價`（如 `4/13收盤價`）
5. 舊日期檔案更名為新日期（如 `*_20260410.xlsx` → `*_20260413.xlsx`）
6. 舊版本備份至 `backup/` 資料夾（**不得直接刪除**）

### ⚠️ 收盤價來源規則（重要）

| 來源 | 角色 | 注意事項 |
|------|------|---------|
| **Yahoo Finance `.TWO`** | ✅ 主要來源 | 上櫃股票用此，當日資料即時 |
| **Yahoo Finance `.TW`** | ✅ 主要來源 | 上市股票用此 |
| **TPEx batch API** | ❌ 禁止單獨使用 | API 有延遲，回傳日期可能是上一交易日 → 必須驗證 `tables[0]['date']` 與查詢日期一致，不一致則捨棄全批 |

### 備份規則

- 每次更新前自動備份舊檔至同層 `backup/` 資料夾
- 備份檔命名：`~$bk_{原檔名}`
- **嚴禁直接刪除原始客戶檔案**，一律先備份再更名

---

### 常見錯誤避免

- **不要**用 `data_only=True` 開啟檔案（會丟失公式）
- **不要**重新建立 Font 物件，只用 `copy.copy()`
- **不要**修改不相關的儲存格
- **不要**在沒有截圖的情況下主動修改檔案
- **不要**用總計列下方的摘要列當格式參考（格式不同）
- **不要**忘記 recalc — 每次 `wb.save()` 之後都要跑 `recalc.py`，否則快取值是舊的
- **不要**預估或計算手續費、交易稅 — 一律以截圖上的數字為準，有 discrepancy 是正常的
- **不要**用公式替代 C25 現金餘額的數值，直接加減更新
- 賣出的「總手續費」(V欄) 要包含買入手續費，用公式表示（如 `=78+95`）

### 多筆交易處理

如果截圖包含多筆交易：
1. 依序辨識每筆交易
2. 列出所有辨識到的交易讓使用者確認
3. 確認後逐筆插入
4. 同一股票的多筆賣出可以合併或分開記錄，詢問使用者偏好

---

## Jeff Fund 專屬佈局（其他客戶佈局可能不同）

### 工作表結構

| 區域 | 列號 | 說明 |
|------|------|------|
| 標題 | Row 1 | `Jeff Fund I (4.2 million NTD)` |
| 持股表頭 | Row 3 | B~K：購買日期、股票名稱、代號、張數、價格、成本、手續費、收盤價、損益、收益率 |
| 持股資料 | Row 4~18 | 買入記錄（空列直接填入，滿了再用 insert_rows） |
| 持股總計 | Row 20 | G20=SUM(G4:G18), H20=SUM(H4:H18), J20=SUM(J4:J18), K20=J20/G20 |
| 摘要區 | Row 25 | B=現金餘額, **C25=數值**, E=股票總成本, F25=G20+H20, H=期初投資額, I25=4233490 |
| 摘要區 | Row 27 | E=未實現損益, F27=J20, H=目前總餘額, I27=F29+C25 |
| 摘要區 | Row 29 | E=股票餘額, F29=F25+F27, H=目前總損益, I29=I27-I25 |
| 賣出表頭 | Row 3 | M~Y（右側） |
| 賣出資料 | Row 4~14 | 賣出記錄 |
| 賣出總計 | Row 37 | T37, U37, V37, W37, X37 |
| 損益摘要 | Row 39~47 | M=未實現損益/已實現損益/已扣費用/目前總損益/報酬率 |

### 股東分配表（Row 35~44）

```
Row 35: B=股東
Row 36: C=家恩, D=老母, E=哲哲, F=Gary, G=總計
Row 37: B=起始金額（硬編碼值）
Row 38: B=佔比        C~F = =C37/$G$37（各股東佔比）
Row 39: B=現金餘額     C~E = =$C$25*C40     F=0    G=SUM
Row 40: B=佔比        C~F = =C38（同起始佔比）
Row 41: B=股票餘額     C~E = =$F$27*C42     F=0    G=SUM
Row 42: B=佔比        C~F = =C40（同起始佔比）
Row 43: B=目前總餘額    C~F = =C41+C39               G=SUM
Row 44: B=目前總損益    C~F = =C43-C37               G=SUM
```

**⚠️ C41~E41 公式規則：`=$F$27*{col}42`**
- F27 = 未實現損益（=J20），代表股票部位的未實現損益
- 乘以各股東佔比（C42/D42/E42），按比例分配
- **不要**改成 $F$29 或其他引用，以使用者設定為準

---

## 直接執行腳本

腳本路徑：`/Users/jeffhsu/Desktop/Claude/個人基金/scripts/bookkeeper.py`

### 使用方式

每次收到交易截圖後，辨識完畢直接執行：

```python
import subprocess, json

trades = [
    {"name":"昶昕","code":"8438","type":"buy","shares":6,"price":80.50,"fee":192,"date":"2026-04-08"},
    {"name":"高力","code":"8996","type":"buy","shares":0.5,"price":903.00,"fee":180,"date":"2026-04-08"},
    # ... 其他筆
]

subprocess.run([
    "python3",
    "/Users/jeffhsu/Desktop/Claude/個人基金/scripts/bookkeeper.py",
    "林峻毅",
    json.dumps(trades, ensure_ascii=False)
])
```

或直接用 Bash 工具：

```bash
python3 /Users/jeffhsu/Desktop/Claude/個人基金/scripts/bookkeeper.py 林峻毅 '[
  {"name":"昶昕","code":"8438","type":"buy","shares":6,"price":80.50,"fee":192,"date":"2026-04-08"}
]'
```

### 腳本自動完成項目

1. 找最新 `{客戶名}_YYYYMMDD.xlsx`
2. 計算 C25 現金餘額（支援公式求值，如 `=982745+1667268`）
3. 依序寫入買入記錄（Row 4 起第一個空白行）
4. 自動抓 Yahoo Finance 收盤價（`.TW` → `.TWO` fallback）填入 I 欄
5. 更新 I3 標題為當日日期（如 `4/8收盤價`）
6. 更新 SUM 公式範圍
7. C25 改寫為數值（不再用公式）
8. 儲存為 `{客戶名}_{今日YYYYMMDD}.xlsx`

### trades 欄位說明

| 欄位 | 必填 | 說明 |
|------|------|------|
| name | ✅ | 股票名稱 |
| code | ✅ | 股票代號（字串） |
| type | ✅ | `"buy"` 或 `"sell"` |
| shares | ✅ | 張數（可為小數，如 0.5） |
| price | ✅ | 成交價格 |
| fee | ✅ | 手續費（截圖數字為準） |
| date | ✅ | 交易日期 `YYYY-MM-DD` |
| tax | ❌ | 賣出交易稅（賣出時填入） |

---

## 優化進度

| 功能 | 狀態 | 說明 |
|------|------|------|
| 買入記帳 | ✅ 已優化 | 自動抓收盤價、現金餘額直接扣減、格式複製 |
| 賣出記帳 | ⏳ 待優化 | 等實際賣出交易時再測試優化 |
| I 欄收盤價 | ✅ 已優化 | Yahoo Finance 自動抓取（.TW → .TWO fallback） |
| 現金餘額 | ✅ 已優化 | C25 直接數值加減，不用公式 |
| 每日收盤價更新 | ✅ 已優化 | 同步更新月選股 + 全部客戶，見下方規則 |
