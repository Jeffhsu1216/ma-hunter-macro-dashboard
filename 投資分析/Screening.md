# Screening.md｜月初選股分析規範

> 每月初輸入 40 支股票代號，自動建立新工作表並填入基本資料，加速選股流程。
> **已優化**：腳本自動化，~3–5 秒完成建表（含 API 並行抓取 + 產業/題材/分群上色）。

---

## 一、主要檔案路徑

| 項目 | 路徑 |
|------|------|
| 正式檔案 | `~/Desktop/Stock Analysis/Jeff_Stock Analysis.xlsx` |
| 作業草稿 | `~/Desktop/Stock Analysis/Jeff_Stock Analysis_Draft.xlsx` |
| 自動化腳本 | `~/Desktop/Claude/投資分析/scripts/screening_builder.py` |

> ⚠️ **預設操作 Draft 檔案**，除非使用者明確說「寫入正式檔案」。

---

## 二、觸發規則

### 觸發詞 A：建立新月工作表
`月初選股 YYYYMM`（附上 40 支股票代號清單，通常為 TQT List PDF 截圖）

**Claude 執行流程：**
1. 從 PDF/截圖讀取 40 支股票代號與名稱
2. 確認 D/E/H 欄日期對應（D=上月E欄、E=上月H欄、H=分析當天）
3. 更新 `screening_builder.py` 的 CONFIG 區：
   - `SHEET` / `PREV_SHEET`
   - `D_COL_SRC=5`（上月 E 欄）、`E_COL_SRC=8`（上月 H 欄）
   - `D_LABEL` / `E_LABEL` / `H_LABEL` / `H_DATE`
   - `STOCKS` 清單（40 支）
   - `MANUAL_EPS`：five91 缺失股的 TTM EPS（需查詢）
   - `PRODUCT_MIX`：40 支的產品別佔比（需 WebSearch 法說會簡報）
   - `NEWS_THEMES`：40 支的投資題材（需 WebSearch 近期新聞，≤10 字）
   - `THEME_GROUPS`：3 支以上重疊題材的分群與底色
4. 執行腳本：`python3 ~/Desktop/Claude/投資分析/scripts/screening_builder.py`
5. 驗證結果：40/40、格式、價格、公式

### 觸發詞 B：更新月中 / 月末追蹤價格
`更新 YYYYMM 收盤價 M/D`

### 觸發詞 C：月底績效回顧
`選股回顧 YYYYMM`

---

## 三、工作表欄位規格（A-R，18 欄）

### 3.1 欄位對照表

| 欄 | 欄位名稱 | 來源 | 說明 |
|----|---------|------|------|
| A | 排序 | 使用者輸入順序 | 1–40 |
| B | 公司代號 | 使用者輸入 | 4-digit |
| C | 公司名稱 | five91 / MANUAL_EPS | |
| D | {前前月末}收盤價 | 上月工作表 E 欄 | 舊股直接複製，新股 YF 補抓 |
| E | {前月末}收盤價 | 上月工作表 H 欄 | 舊股直接複製，新股 YF 補抓 |
| F | 前個月漲幅 | 公式 `=E/D-1` | CF: >20% 紅字紅底 |
| G | 排名 | 公式 `=RANK(F,$F$2:$F$41)` | |
| H | {分析日}收盤價 | TPEx batch（OTC）+ YF（TWSE） | 分析當日收盤價 |
| I | 第1個月漲幅 | 公式 `=H/E-1` | CF: >0 紅字（複製上月） |
| J | 排名 | 公式 `=RANK(I,$I$2:$I$41)` | |
| K | 產業 | **腳本自動填**（WebSearch 法說會） | 產品別佔比，黑字無底色 |
| L | 題材 | **腳本自動填**（WebSearch 新聞） | ≤10 字，3+支同題材上色分群 |
| M | 第一階段選擇 | 自動預篩 + 人工確認 | EPS<0 自動「否」灰底黑字 |
| N | 第二階段選擇 | **人工判斷** | |
| O | 2025EPS | five91 TTM 計算 | `net_income_ttm / share_capital` |
| P | Current PE | 公式 / 文字 | EPS<0 → 「EPS為負數」紅字 |
| Q | 已發行普通股數 | TWSE/TPEx 官方 OpenAPI | 非面額推算 |
| R | 市值（億） | 公式 `=Q*E/1e8` | |

### 3.2 格式規則

- **格式複製上月工作表**：font/fill/alignment/border/number_format
- **Grid lines OFF**
- **K 欄**：黑字（`Font(color="000000")`），無底色
- **L 欄**：題材文字 + 產業動能分群底色（3+ 支同題材）
- **M 欄**：「否」= 灰底 `D9D9D9` 黑字，其餘無底色
- **P 欄**：EPS 為負 → 「EPS為負數」紅字 `C00000`
- **F 欄 CF**：>20% → 紅字 `9C0006` + 紅底 `FFC7CE`（DifferentialStyle bgColor 格式）
- **Row 44**：大盤加權指數（D/E/H 三欄），無標籤

### 3.3 收盤價資料來源

| 欄位 | 資料來源 |
|------|---------|
| D/E（舊股）| 上月工作表 E 欄(col 5) / H 欄(col 8) |
| D/E（新股）| Yahoo Finance `.TW` / `.TWO` |
| H（OTC）| TPEx batch 一次抓完 |
| H（TWSE）| Yahoo Finance `.TW` → `.TWO` fallback |
| TAIEX D/E | 上月工作表 Row 44 |
| TAIEX H | Yahoo Finance `^TWII` |
| 已發行普通股數 | TWSE `t187ap03_L` + TPEx `mopsfin_t187ap03_O` |

### 3.4 L 欄產業動能分群規則

- 將 40 支股票依投資題材歸類
- **3 支以上**同一大類 → 同底色標記
- 預設 4 色：

| 群組 | 底色 |
|------|------|
| AI伺服器 | 粉紅 `FFCCCC` |
| CoWoS/先進封裝 | 淡紫 `E2CCFF` |
| 光通訊/CPO | 淡綠 `CCFFCC` |
| 半導體設備/耗材 | 淡藍 `CCE5FF` |

> 每月依實際題材重新分群，顏色可沿用或調整。

### 3.5 J46:L56 績效統計區

| Row | J | K | L |
|-----|---|---|---|
| 46 | 個人績效 | 留空（選完後手動填） | |
| 47 | 基金績效 | `=AVERAGE(I2:I41)` | |
| 48 | 加權指數 | `=H44/E44-1` | |
| 49 | 打敗基金 | `=IF(K46="","",IF(K46>K47,"Yes","No"))` | |
| 50 | 打敗大盤 | `=IF(K46="","",IF(K46>K48,"Yes","No"))` | |
| 52 | 上漲家數 | `=COUNTIF($I$2:$I$41,">0")` | `=K52/(K52+K53)` |
| 53 | 下跌家數 | `=COUNTIF($I$2:$I$41,"<0")` | `=K53/(K53+K52)` |
| 55 | 未跳出 | 與上月重疊股數（自動計算） | |
| 56 | 跳出 | 40 - K55 | |

### 3.6 重疊股標記

- 與上月工作表重疊的股票 → **A~G 欄淺綠底色** `C6EFCE`
- **F 欄除外**（保留給 >20% CF 紅色優先）

---

## 四、腳本 CONFIG 區（每月更新）

```python
SHEET       = "YYYYMM"           # 本月工作表名
PREV_SHEET  = "YYYYMM"           # 上個月工作表名
D_COL_SRC   = 5                  # 上月 E 欄 (col index, 1-based)
E_COL_SRC   = 8                  # 上月 H 欄
D_LABEL     = "M/D收盤價"         # D 欄標題
E_LABEL     = "M/D收盤價"         # E 欄標題
H_LABEL     = "M/D收盤價"         # H 欄標題
H_DATE      = "YYYYMMDD"         # H 欄日期（API 用）
STOCKS      = [('代號','名稱'), ...]  # 40 支股票清單
MANUAL_EPS  = {'代號': {'ttm_eps': X.XX, 'name': '...'}, ...}  # five91 缺失股
PRODUCT_MIX = {'代號': '產品A ~X%；...', ...}  # K 欄（WebSearch 法說會）
NEWS_THEMES = {'代號': '題材', ...}             # L 欄（WebSearch 新聞，≤10字）
THEME_GROUPS = {'群組名': {'color':'XXXXXX','codes':[...]}, ...}  # L 欄分群上色
```

---

## 五、API 說明

### five91（財務指標，非價格）
`GET https://five91.onrender.com/api/metrics`
- `name`、`net_income_ttm`、`share_capital` → 計算 EPS
- ⚠️ `price` 欄位為舊價格，**不用於收盤價填入**
- 部分股票未收錄（如 2887、8438、6894、6015）→ `MANUAL_EPS` 手動補

### TPEx 上櫃收盤價（OTC batch）⚠️ 有延遲風險
`GET https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&d={ROC日期}&o=json`
- `tables[0]['data']`，`row[0]`=代號、`row[2]`=收盤價
- ⚠️ **必須驗證** `tables[0]['date']` 與查詢日期一致，否則是舊資料 → 捨棄，改用 Yahoo Finance

### Yahoo Finance（TWSE 個股 + TAIEX）
- `https://query1.finance.yahoo.com/v8/finance/chart/{CODE}.TW` / `.TWO`
- `https://query1.finance.yahoo.com/v8/finance/chart/%5ETWII`

### TWSE/TPEx 已發行普通股數（官方 OpenAPI）
- TWSE: `GET https://openapi.twse.com.tw/v1/opendata/t187ap03_L` → `已發行普通股數或TDR原股發行股數`
- TPEx: `GET https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O` → `IssueShares`
- ⚠️ 不用面額推算，直接取官方已發行股數

---

## 六、每月操作 SOP

1. 使用者提供 TQT List PDF（40 支代號）
2. Claude 讀取 PDF → 解析代號與名稱
3. Claude 確認日期對應：
   - D 欄 = 上月 E 欄日期（前前月末收盤）
   - E 欄 = 上月 H 欄日期（前月末收盤）
   - H 欄 = 今天（分析日收盤）
4. Claude 並行搜尋：
   - **K 欄**：40 支法說會簡報 → 產品別佔比（Agent）
   - **L 欄**：40 支近期新聞 → 投資題材 ≤10 字（Agent）
   - **MANUAL_EPS**：five91 缺失股查 TTM EPS（Agent）
5. Claude 分析 L 欄題材 → 歸類 THEME_GROUPS（3+ 支同色）
6. Claude 更新腳本 CONFIG 區所有欄位
7. 執行腳本 → 驗證 40/40 + 格式
8. 批次 append 至 Company Overview（比對已存在 → 跳過，新股 → 新增）
9. 輸出建表確認

---

## 七、選股輔助提示（每次建表後自動輸出）

```
第一階段「考慮」篩選條件：
  ✅ 前月漲幅排名前 30%（排名 ≤ 12）
  ✅ 有明確題材（L 欄有內容）
  ✅ PE < 30x（排除高估值）
  ❌ 虧損股（TTM EPS < 0）→ 已自動填「否」

第二階段「選擇」篩選條件：
  ✅ 第一階段考慮 + PE 10-25x
  ✅ 有催化劑（題材 / 儲董看好 / 法說會）
  ✅ 市值 < 500 億（流動性適中的中小型股）
  目標選出 3-6 支
```

---

## 八、Company Overview 批次 Append

建表完成後，自動將 40 支股票寫入 `Company Overview_{YYYYMMDD}.xlsx`：

1. 讀取現有最新 Company Overview 檔案
2. 以 B 欄（股票代號）比對，**已存在 → 跳過**，不存在 → 新增
3. 新增列填入：序號(A)、代號(B)、名稱(C)、營收(D)、毛利率(E)、營益率(F)、淨利率(G)、稅後淨利(H)、資本額(I)、EPS(J)、產品佔比(K)、產業(L)、題材(M)
4. 資料來源：five91 `/api/metrics`（財務）+ PRODUCT_MIX（K欄）+ NEWS_THEMES（M欄）
5. 格式：無底色 + 全欄置中
6. 存檔後刪除舊日期檔案（資料夾永遠只保留一份）

> 詳細欄位規格見 `CompanyOverview_AutoAppend.md`

---

## 九、注意事項

- 預設操作 `Jeff_Stock Analysis_Draft.xlsx`，確認結果正確後再說「寫入正式檔案」
- 所有計算欄位使用 Excel 公式，不寫死數字
- 工作表命名格式嚴格為 `YYYYMM`（6 碼）
- 格式完全複製上月工作表（font/fill/alignment/border）
- K 欄強制黑字（上月模板可能是白字配彩色底）
- 已發行普通股數從官方 API 抓取，不用面額推算
- 若代號查無資料，該列保留代號，其餘欄位留空並於回報中列出

---

*最後更新：2026-04-07（v5 — K產品佔比+L題材分群+已發行股數+績效統計區+重疊股標記+Company Overview batch append）*
