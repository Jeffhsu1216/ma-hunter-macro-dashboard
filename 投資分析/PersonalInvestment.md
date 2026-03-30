# 個人投資回報分析規範

> 觸發詞：「個人投資回報分析」
> 分析期間：**2025/09/01 起**（之前資料不納入）
> ⚠️ **嚴禁對 Excel 進行任何寫入或修改，以 `data_only=True` 唯讀開啟**

---

## 一、資料來源

| 項目 | 內容 |
|------|------|
| Excel 路徑 | `~/Desktop/Investment/Personal Financial Management.xlsx` |
| 工作表 | `Investment` |
| 讀取方式 | `openpyxl.load_workbook(..., data_only=True)` |
| 輸出目錄 | `~/Desktop/Investment/Reports/` |
| 輸出檔名 | `個人投資回報分析_YYYYMMDD.pdf` |

---

## 二、欄位對應（Investment 工作表）

| 欄 | 欄位名稱 | Python key | 說明 |
|----|---------|------------|------|
| A | 序號 | `seq` | |
| B | 股票名稱 | `name` | |
| C | 數量（張） | `qty` | |
| D | 購買價格 | `buy_p` | |
| E | 購買日期 | `buy_d` | 用於判斷買入月份與過濾 2025/09 起 |
| F | 購買成本 | `cost` | |
| G | 出售價格 | `sell_p` | |
| H | 出售日期 | `sell_d` | 有值 = 已實現；None = 持倉中 |
| J | 持有天數 | `days` | |
| K | 持有期間報酬率 | `ret` | |
| M | 帳面盈利 | `paper` | 持倉中使用此欄作為未實現損益 |
| N | 手續費 | `fee` | |
| O | 交易稅 | `tax` | |
| P | 實際盈利 | `profit` | 扣除手續費 + 稅後淨損益 |
| **Q** | **月度批次總損益** | — | 每買入月份最後一列有值，含已實現＋未實現 |
| **R** | **月度批次報酬率** | — | 對應 Q 欄同列 |

---

## 三、月度 Q/R 欄對應列號

> Q 欄在每個「買入月份群組」的最後一列標記月度總損益，直接讀取即可，無需自行計算。

| 買入月份 | Q/R 所在列 | 代表股票 |
|---------|-----------|---------|
| 2025/09 | Row 48 | 南電 (8460) |
| 2025/10 | Row 50 | 瀚宇博 (5469) |
| 2025/11 | Row 52 | 緯軟 (4953) |
| 2025/12 | Row 56 | 景碩 (3189) |
| 2026/01 | Row 60 | 群創 (3481) |
| 2026/02 | Row 67 | 朋億* (6613) |
| 2026/03 | Row 71 | 凡甲 (3526) |

```python
QR_ROWS = {
    '2025/09': 48, '2025/10': 50, '2025/11': 52, '2025/12': 56,
    '2026/01': 60, '2026/02': 67, '2026/03': 71,
}
for ym in MONTHS:
    ri = QR_ROWS[ym]
    q  = ws.cell(ri, 17).value or 0   # Q欄
    rv = ws.cell(ri, 18).value or 0   # R欄
```

---

## 四、Benchmark 比較資料（L102:T111）

| 列 | 內容 |
|----|------|
| 102 | 期間標籤（月份日期） |
| 103 | 大盤指數絕對值 |
| 104 | **大盤月度績效** |
| 105 | 大盤累計績效（僅最末欄有值） |
| 106 | **基金月度績效** |
| 107 | 基金累計績效（僅最末欄有值） |
| 108 | **個人月度績效** |
| 109 | 個人累計績效（僅最末欄有值） |
| 110 | 是否打贏基金（Yes / No） |
| 111 | 是否打贏大盤（Yes / No） |

**欄位對應：**

| Excel欄 | 期間 | 備註 |
|---------|------|------|
| M（col 13） | 2025/09 | 基準期，所有績效欄為 None |
| N（col 14） | 2025/10 | |
| O（col 15） | 2025/11 | |
| P（col 16） | 2025/12 | |
| Q（col 17） | 2026/01 | |
| R（col 18） | 2026/02 | |
| S（col 19） | 2026/03 | |
| T（col 20） | 2026/03/26 | 最新截止日 |

```python
bm_cols  = list(range(14, 21))   # N~T，共 7 期有效報酬
mkt_ret  = [ws.cell(104, c).value for c in bm_cols]
fund_ret = [ws.cell(106, c).value for c in bm_cols]
ind_ret  = [ws.cell(108, c).value for c in bm_cols]
beat_fund = [ws.cell(110, c).value for c in bm_cols]
beat_mkt  = [ws.cell(111, c).value for c in bm_cols]
BM_LABELS = ['25/10','25/11','25/12','26/01','26/02','26/03','03/26']
```

**累計報酬計算（複利，含基準期起始點 0%）：**
```python
def cum_ret(rets):
    vals = [0.0]; cur = 1.0
    for r in rets:
        if r is not None: cur *= (1 + r)
        vals.append((cur - 1) * 100)
    return vals  # 長度 = len(rets) + 1（含 09月基準）
```

---

## 五、Hero 指標計算

```python
CUT = datetime.datetime(2025, 9, 1)
recent      = [r for r in all_rows if r['buy_d'] and r['buy_d'] >= CUT]
completed_r = [r for r in recent if r['sell_d']]
open_pos_r  = [r for r in recent if not r['sell_d']]

tot_real    = sum(t['profit'] for t in completed_r)   # 已實現損益
tot_unreal  = sum(p['paper']  for p in open_pos_r)    # 未實現損益
tot_pnl     = tot_real + tot_unreal                    # 總損益

# 勝率：持倉中以 paper > 0 計勝
wins = sum(1 for t in trades if t['profit'] > 0 or (not t['sell_d'] and t['paper'] > 0))
```

---

## 六、報告結構（4頁 PDF）

### Page 1 — 績效總覽

| 區塊 | 座標 `[l, b, w, h]` | 內容 |
|------|-------------------|------|
| 標題 | `[0, 0.924, 1, 0.076]` | 酒紅底 / 金色字 |
| Hero | `[0, 0.820, 1, 0.104]` | 已實現損益 / 未實現損益 / 總損益 |
| 左圖（長條） | `[0.048, 0.350, 0.490, 0.440]` | 月度批次總損益（Q欄），**僅長條圖，不加累積折線與右側Y軸** |
| 分隔條 | `[DIV_L, 0.350, 0.004, 0.440]` | 顏色 `#C8B89A` |
| 右圖（散佈） | `[0.645, 0.375, 0.330, 0.415]` | 持有天數 vs 報酬率（點大小＝投入成本） |
| KPI 卡片 | `[0, 0.155, 1, 0.195]` | 5張：已出售筆數 / 勝率 / 平均持有天數 / 最佳單筆 / 最差單筆 |
| 分析文字 | `[0, 0, 1, 0.155]` | 酒紅底，4行摘要 |

> **散佈圖獨立往上偏移**（`bottom=0.375` vs 長條圖 `0.350`），確保 X 軸標籤「持有天數」不被截斷。

### Page 2 — 月度交易分析

- 上半：月度損益長條 + 勝率折線（雙 Y 軸）
  - **勝率折線須在每個數據點標注百分比數值**（如 `80%`），使用 `ax.text()` 標注於點上方
- 下半：7 個買入月份卡片
  - 卡片內容：月份標題、Q 欄損益、R 欄報酬率、勝率、個別交易列表
  - `◆` 標記持倉中交易
  - **交易列表須顯示完整股票名稱含代號**（如 `南電 (8460)`），字體縮小至 5pt 以容納

### Page 3 — 交易明細與持倉

- 當前持倉明細表（合併同名持倉）
- 持倉成本圓餅圖
- **全部交易明細表必須完整顯示所有交易**（不可截斷），使用自適應行高：
  - `row_height = min(0.026, 0.92 / max(total_trades, 1))`
  - 字體大小隨交易筆數調整（超過 25 筆時縮小至 5.5-6pt）
- **「股票名稱」欄須顯示完整名稱含代號**（如 `南電 (8460)`），不可截斷
  - 欄位佈局：`t_hx = [0.02, 0.06, 0.22, 0.28, 0.38, 0.48, 0.55, 0.68, 0.84]`
  - 股票名稱欄寬度從 0.06~0.22 擴展至 0.06~0.22（增加空間容納代號）
- 依買入日排序，含持倉中

### Page 4 — 績效比較

- 月度報酬率 Grouped Bar（個人 / 基金 / 大盤，7期）
- 累積報酬率折線圖（複利，含基準點 09月(基)）
  - **三條折線（個人/基金/大盤）須在每個數據點標注累計報酬率百分比**，使用 `ax.text()` 並設定偏移避免重疊
- 月度績效紀錄表（含是否打贏基金 / 大盤，顯示「是 / 否」）

---

## 七、配色主題

```python
WINE    = '#7B1C2E'   # 主色（酒紅）正值長條 / 已實現
WINE_DK = '#4A0E1A'   # 深酒紅（標題背景 / 總損益文字）
WINE_MD = '#9B2335'   # 中酒紅（表格標頭）
GOLD    = '#B8860B'   # 金色（累積折線 / 未實現損益）
GOLD_LT = '#D4A843'   # 淺金（標題文字）
GOLD_PAL= '#F5E6B8'   # 極淺金（基金列背景）
CREAM   = '#FDFBF7'   # 米白（頁面底色）
GRAY    = '#9E9E9E'   # 灰（負值長條）
GRAY_LT = '#F0EDE8'   # 淺灰（卡片/列背景）
GRAY_DK = '#555555'   # 深灰（次要文字）
DIVIDER = '#C8B89A'   # 分隔條
GREEN_C = '#2E7D32'   # 綠（打贏標記）
BLUE_C  = '#1565C0'   # 藍（大盤）
```

---

## 八、字型規範

```python
FONT = 'Heiti TC'
plt.rcParams.update({
    'font.family': FONT,
    'font.size': 9,
    'axes.unicode_minus': False,
    'figure.dpi': 150,
})
# 所有 text / title / legend 均需明確指定 fontfamily=FONT
```

---

## 九、PDF 合併（單頁重新產生時使用）

當只需更新特定頁，不重跑全部時：

```python
from pypdf import PdfReader, PdfWriter
w = PdfWriter()
w.add_page(PdfReader(new_page_pdf).pages[0])   # 新頁
for p in PdfReader(old_pdf).pages[1:]:         # 保留其他頁
    w.add_page(p)
with open(old_pdf, 'wb') as f:
    w.write(f)
os.remove(new_page_pdf)   # 清理暫存檔
```

---

## 十、注意事項

1. 分析期間固定從 `2025/09/01` 起，過濾條件：`buy_d >= datetime(2025, 9, 1)`
2. Q 欄值已含已實現＋未實現，**直接讀取，不要自行重算**
3. Benchmark Sep 欄（col M）為基準期，值為 None，累積從 1.0 開始複利
4. 年化報酬率計算易受短線交易影響，**不列為主要指標**（散佈圖呈現即可）
5. 散佈圖 `bottom=0.375`（比長條圖高 0.025），避免 X 軸標籤被截斷
6. 打贏基金 / 大盤顯示「是 / 否」，不使用特殊符號（Heiti TC 不支援 ✓ ✗）
7. 輸出前確認 `~/Desktop/Investment/Reports/` 目錄存在，否則 `os.makedirs(out_dir, exist_ok=True)`
8. **嚴禁寫入 Excel**

---

*最後更新：2026-03-27*
*修正紀錄：股票名稱全面顯示完整代號（Page 2 月度卡片 + Page 3 持倉/交易明細）*
