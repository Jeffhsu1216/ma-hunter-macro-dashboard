# refs/PersonalInvestment_Script.md｜個人投資回報分析完整腳本

> 本檔僅在 debug / 腳本維護時讀取。日常執行讀 `PersonalInvestment.md` 即可。

---

## 一、版面規格（依 20260327 模板）

### Page 1 — 績效總覽

```
[標題列]  h=0.070  dark wine bg  左："個人投資回報分析" gold 24pt bold  右：日期 9pt
[Hero]    b=0.830  3格均分 各 w=0.295 h=0.088  CREAM bg + DIVIDER 邊框（無內部條）
          標籤 gray 9pt 上方，數值 colored 20pt bold 下方
[長條圖]  b=0.390 l=0.06 w=0.455 h=0.410  y軸:"月度損益額 (NTD)"  bar標籤實際NTD值
[分隔線]  l=0.526 w=0.003 h=0.410  DIVIDER色
[散佈圖]  b=0.412 l=0.548 w=0.415 h=0.388  wine=已賣出 gold=持倉中 圖例右上
[KPI卡片] b=0.205 5張均分  GRAY_LT bg + DIVIDER 邊框  ← 無 wine 頂條！
          灰色標籤 9pt 上方（y=0.75）大值 11pt bold colored 中間（y=0.38）
[底欄]    b=0.000 h=0.195  WINE_DK bg  4行 GOLD_LT 8.5pt 文字
```

**KPI 卡片顏色規則：**
- 已出售筆數：GRAY_DK
- 勝率：GREEN_C（正）/ RED_C（負）
- 平均持有天數：GRAY_DK
- 最佳單月：RED_C（台灣：獲利=紅）
- 最差單月：GREEN_C（台灣：虧損=綠）

### Page 2 — 月度交易分析

```
[標題列]  h=0.050  dark wine  左:"月度交易分析" gold 16pt bold
[雙軸圖]  b=0.665 l=0.09 w=0.82 h=0.260
          左Y:損益NTD  右Y:勝率0-120%  金色菱形◆標記 數值標在點上方
[月卡片]  b=0.055 h=0.590  7張均分 各 w=0.125
          wine 頂條(h=9%)：月份 white bold
          損益數字(y=0.840)：colored 9.5pt bold
          報酬率%(y=0.775)：colored 8pt
          勝率文字(y=0.715)：gray 7pt  "勝率 XX%"
          分隔線(y=0.700)
          交易列表：名稱左對齊 損益右對齊 5.5pt，每行 y間距0.062
          開倉持股以 * 前綴標記
```

### Page 3 — 交易明細

```
[標題列]  h=0.050  dark wine  左:"交易明細與持倉" gold 16pt bold
[表格]    b=0.020 l=0.02 w=0.96 h=0.915
          表頭行：WINE_MD bg white text 7.5pt bold
          欄位(hx)：#=0.005 股票名稱=0.032 張數=0.190 買入日=0.238
                    賣出日=0.308 天數=0.382 成本=0.440 損益=0.603 報酬率=0.772
          交替行：GRAY_LT / CREAM
          持倉中 → 賣出日顯示「持倉中」
          頁腳："共 N 筆交易（已賣出 N / 持倉中 N）" centered 8pt
```

### Page 4 — 績效比較

```
[標題列]  h=0.050  dark wine  左:"績效比較" gold 16pt bold
[Grouped Bar]  b=0.660 l=0.09 w=0.84 h=0.265  個人=WINE 基金=GOLD 大盤=BLUE_C
[累積折線]     b=0.060 l=0.07 w=0.52 h=0.570  三線+點標籤（每點標%數值）
[績效表]       b=0.060 l=0.62 w=0.36 h=0.570
               表頭 WINE_MD：月份|個人|基金|贏基金|贏大盤
               是/否 → 顯示「是」「否」（中文）  ← 非 Yes/No！
               是=RED_C  否=GREEN_C  （台灣慣例）
               頁腳："累計：個人 XX%｜基金 XX%｜大盤 XX%"
```

---

## 二、配色與字型

```python
WINE    = '#7B1C2E'   # 酒紅（正值長條）
WINE_DK = '#4A0E1A'   # 深酒紅（標題背景）
WINE_MD = '#9B2335'   # 中酒紅（表頭/卡片頭）
GOLD    = '#B8860B'   # 金色（折線/未實現）
GOLD_LT = '#D4A843'   # 淺金（標題文字）
GOLD_PAL= '#F5E6B8'   # 極淺金（交替行）
CREAM   = '#FDFBF7'   # 頁面底色
GRAY    = '#9E9E9E'   # 灰（負值長條）
GRAY_LT = '#F0EDE8'   # 淺灰（卡片/交替行）
GRAY_DK = '#555555'   # 深灰（次要文字）
DIVIDER = '#C8B89A'   # 分隔條/邊框
GREEN_C = '#2E6B30'   # 綠（正值/打贏）
RED_C   = '#9B2335'   # 紅（負值/輸）
BLUE_C  = '#1565C0'   # 藍（大盤）
FONT    = 'Heiti TC'
DPI     = 150
```

---

## 三、腳本關鍵邏輯

### 資料讀取
```python
# 原始 Excel（formula-cached values，data_only=True）
XL_PATH = '~/Desktop/Personal Financial Management/Personal Financial Management.xlsx'
wb = openpyxl.load_workbook(XL_PATH, data_only=True)
ws = wb['Investment']
CUT = datetime.datetime(2025, 9, 1)  # 分析起始日

# 讀取行資料
for r in range(2, ws.max_row + 1):
    name  = ws.cell(r, 2).value   # B：股票名稱
    buy_d = ws.cell(r, 5).value   # E：購買日期（datetime）
    ...
    paper  = ws.cell(r, 13).value  # M：帳面盈利
    profit = ws.cell(r, 16).value  # P：實際盈利
```

### Jeff 持倉補注邏輯（⚠️ 關鍵：APPEND 模式）
```python
# 從 Jeff_YYYYMMDD.xlsx 讀取未出場部位 → 乘以 Jeff 比例
# 用 ws.cell(r, col).value 找最後一個有名稱的行
last_data_row = max(r for r in range(2, ws.max_row+1)
                    if ws.cell(r, 2).value and
                    isinstance(ws.cell(r, 5).value, datetime.datetime))
# 從 last_data_row + 1 開始 APPEND，絕對不插入空白行
```

### Benchmark 讀取
```python
bm_cols   = list(range(14, 21))           # 欄 N~T（col 14~20）
mkt_ret   = [ws.cell(104, c).value ...]   # 列 104：大盤月度績效
fund_ret  = [ws.cell(106, c).value ...]   # 列 106：基金月度績效
ind_ret   = [ws.cell(108, c).value ...]   # 列 108：個人月度績效
beat_fund = [ws.cell(110, c).value ...]   # 列 110：是否打贏基金
beat_mkt  = [ws.cell(111, c).value ...]   # 列 111：是否打贏大盤
# ⚠️ 以上列號為原始檔案位置。
# 若新增資料超出列 81，benchmark 列號不受影響（benchmark 區固定於列 102-111）
```

### QR_ROWS（月底批次損益摘要列）
```python
QR_ROWS = {
    '2025/09': 48,
    '2025/10': 50,
    '2025/11': 52,
    '2025/12': 56,
    '2026/01': 60,
    '2026/02': 67,
    '2026/03': 71,
    # 每月新增：'YYYY/MM': 對應 Excel 列號
}
```

---

## 四、已知陷阱

| 陷阱 | 說明 | 解法 |
|------|------|------|
| formula cells 在 draft 回傳 None | openpyxl data_only 無法計算公式 | 讀原始 Excel（已被 Excel 存檔過），Jeff 持倉手動補注 |
| benchmark 列號偏移 | APPEND 的新行不影響固定 benchmark 區（列 102-111） | 永遠 APPEND，不 INSERT |
| 持倉報酬率顯示 | open_pos_r 用 paper/cost 計算（不用 ret 欄） | 腳本已正確處理 |

---

*最後更新：2026-04-12（v1 — 從 PersonalInvestment.md 拆出）*
