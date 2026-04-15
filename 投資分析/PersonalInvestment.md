# PersonalInvestment.md｜個人投資回報分析規範（Slim v2）

> 執行時讀此檔即可。版面規格 + 腳本邏輯 → `refs/PersonalInvestment_Script.md`

---

## 一、主要檔案路徑

| 項目 | 路徑 |
|------|------|
| 個人 Excel | `~/Desktop/Personal Financial Management/Personal Financial Management.xlsx` |
| 唯一腳本 | `~/Desktop/Claude/投資分析/scripts/personal_inv.py` |
| PDF 輸出 | `~/Desktop/Personal Financial Management/` |

---

## 二、觸發規則：`個人投資回報分析`

### Step 1 — 同步 Jeff 基金持倉

讀取 `JeffFundSync.md` → 執行完整轉換流程（比例計算 + APPEND + 月度統計更新）

### Step 2 — 執行分析腳本

```bash
python3 ~/Desktop/Claude/投資分析/scripts/personal_inv.py
```

腳本自動：讀 Excel → 統計損益 → 繪製 4 頁圖 → 合併 PDF → 儲存

### Step 3 — 更新技能熟練度

`skill_proficiency.md` PersonalInvestment +1

---

## 三、四頁 PDF 內容

| 頁 | 標題 | 主要內容 |
|----|------|---------|
| 1 | 績效總覽 | Hero（已實現/未實現/總損益）+ 月度長條圖 + 散佈圖 + KPI 卡片 |
| 2 | 月度交易分析 | 雙軸圖（損益/勝率）+ 7張月卡片（交易列表，開倉以 `*` 標記） |
| 3 | 交易明細與持倉 | 全部交易表格，持倉中顯示「持倉中」 |
| 4 | 績效比較 | 個人 vs 基金 vs 大盤（月度分組 + 累積折線 + 績效表） |

---

## 四、QR_ROWS 更新（每月月底）

1. 開啟 Excel → 找該月最後一筆交易列 → 確認 Q 欄有月度損益值
2. 更新腳本 `QR_ROWS`：`'YYYY/MM': 列號`
3. 更新 `BM_LABELS` 最後一個標籤（格式 `MM/DD`）

**目前設定：**
```python
QR_ROWS = {
    '2025/09': 48, '2025/10': 50, '2025/11': 52, '2025/12': 56,
    '2026/01': 60, '2026/02': 67, '2026/03': 71,
}
BM_LABELS = ['25/10','25/11','25/12','26/01','26/02','26/03','03/31']
```

---

## 五、欄位對應（Investment 工作表）

| 欄 | 欄位 | col# |
|----|------|------|
| B | 股票名稱 | 2 |
| C | 數量（張） | 3 |
| D | 購買價格 | 4 |
| E | 購買日期 | 5 |
| F | 購買成本 | 6 |
| G | 出售價格 | 7 |
| H | 出售日期 | 8 |
| J | 持有天數 | 10 |
| K | 持有期間報酬率 | 11 |
| M | 帳面盈利 | 13 |
| P | 實際盈利 | 16 |
| Q | 月度批次總損益 | 17 |
| R | 月度批次報酬率 | 18 |

**Benchmark（固定區段）**：列 104=大盤 / 106=基金 / 108=個人 / 110=贏基金 / 111=贏大盤；欄 N~T（col 14~20）

---

## 六、重要注意事項

| 規則 | 說明 |
|------|------|
| ⚠️ 同步規則 | APPEND / 比例計算 / 欄位格式 → 全部見 `JeffFundSync.md` |
| ⚠️ 腳本勿重寫 | 直接執行 personal_inv.py，不重新生成腳本 |
| formula 陷阱 | draft 檔 formula cells 回傳 None → 永遠讀原始 Excel |

---

*更新：2026-04-12（Slim v2 — 腳本規格移至 refs/PersonalInvestment_Script.md，token ~1.5k）*
