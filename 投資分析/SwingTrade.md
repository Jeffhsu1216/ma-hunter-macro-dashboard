# SwingTrade.md｜短線交易分析規範（Slim v10）

> 執行時讀此檔即可。評分方法論 → `refs/SwingTrade_Scoring.md`

---

## 一、主要檔案路徑

| 項目 | 路徑 |
|------|------|
| 唯一腳本 | `~/Desktop/Claude/投資分析/scripts/swing_analysis.py` |
| 交易紀錄 | `~/Desktop/Investment/SwingTrade_Log.xlsx` |
| PDF 輸出 | `~/Desktop/Stock Analysis/Reports/SwingTrade/` |

---

## 二、觸發規則

### A — 短線分析（標的評估）`短線分析 {代號}`

1. 兩個 WebSearch **並行**（禁止手動 WebFetch TWSE）：
   - A：近 7 日催化劑新聞 → `catalyst` / `catalyst_quality` / `sector_name`
   - B：產品別佔比（法說會/研究報告）→ `product_mix`
2. 執行腳本：
   ```python
   from swing_analysis import run
   run(CODE, '', TODAY,
       sector_name="...",            # WebSearch A
       catalyst="...",               # WebSearch A（30字內）
       catalyst_quality="強/中/弱",  # WebSearch A
       product_mix="產品A ~X%；產品B ~X%")  # WebSearch B
   ```
3. 腳本自動：OHLCV + 法人 + TAIEX 並行抓取（~2s）→ 評分 → PDF → Company Overview Append
4. `skill_proficiency.md` SwingTrade +1

**評分維度：** 基本面(20) + 產業動能(40⭐) + 技術面(20) + 籌碼面(20) = 100分
- 75+ 🟢 建議進場 ／ 55–74 🟡 觀望 ／ <55 🔴 不建議
- 詳細評分邏輯見 `refs/SwingTrade_Scoring.md`

**罰分快速參考：**
| 條件 | 處理 |
|------|------|
| TTM EPS ≤ 0 | 封頂 54 |
| 產業動能 < 20/40 | 封頂 54 |
| RSI > 75 | −10 |
| 外資 5 日 < −2,000 張 | −8 |
| 融資 5 日增 > 3,000 張 | −5 |

---

### B — 進場記錄 `短線進場 {代號} {張數} {價格}`

1. 開啟 `SwingTrade_Log.xlsx` → `交易紀錄` 末列新增
2. 填入：編號(ST-00N) / 進場日期 / 代號 / 名稱(API) / 產業 / 進場價 / 張數
3. 計算：進場金額 / ATR(10)停損(×1.5, max -8%) / ATR(10)停利(×3) / R/R
4. 詢問消息來源、進場理由（若有短線分析評分自動帶入）

---

### C — 出場記錄 `短線出場 {代號} {價格}`

1. `SwingTrade_Log.xlsx` 找對應未出場紀錄
2. 填入出場日期 / 出場價 → 公式自動算報酬率 / 損益 / 持有天數 / 結果
3. 提示填寫「事後檢討」，對照進場評分做回顧

---

### D — 週報回顧 `短線週報`

輸出：本週出場筆數/損益/勝率 + 持有中部位表（含追蹤止盈線）+ 累積統計
⚠️ 接近停損（報酬率 < -5%）/ 🎯 接近停利（報酬率 > +12%）

---

### E — 績效總覽 `短線績效`

PDF（輸出至 `~/Desktop/Investment/Reports/SwingTrade/`）：
- 累積損益曲線 / 月度勝率長條圖 / 消息來源勝率比較
- 紀律內 vs 紀律外勝率 / 進場評分 vs 報酬率散佈圖 / 最佳最差5筆

---

## 三、交易紀錄欄位（SwingTrade_Log.xlsx `交易紀錄`）

| 欄 | 欄位 | 說明 |
|----|------|------|
| A | 編號 | ST-001, ST-002, ... |
| B | 進場日期 | YYYY/MM/DD |
| C | 代號 | |
| D | 名稱 | API 帶入 |
| E | 產業 | 子產業分類 |
| F | 進場價 | |
| G | 張數 | |
| H | 進場金額 | `=F×G×1000` |
| I | 消息來源 | 內線/技術面/題材/法說會/其他 |
| J | 進場理由 | 50字內 |
| K | 停損價 | ATR(10)×1.5，上限 -8% |
| L | 停利價 | ATR(10)×3，無上限 |
| M | R/R | `=(L-F)/(F-K)` |
| N | 進場評分 | 100分制 |
| O | 出場日期 | |
| P | 出場價 | |
| Q | 報酬率 | `=(P-F)/F` |
| R | 損益金額 | `=(P-F)×G×1000` |
| S | 持有天數 | `=O-B` |
| T | 結果 | ✓獲利 / ✗虧損 |
| U | 符合紀律 | 評分≥70為是 |
| V | 事後檢討 | 出場後填寫 |

---

## 四、風險控管

| 規則 | 設定 |
|------|------|
| 停損上限 | 1.5×ATR(10)，max -8% |
| 單筆部位 | ≤ 總資金 20% |
| 同時持有 | ≤ 5 檔 |
| 單週最大虧損 | 達 -5% → 暫停至下週 |
| 連續虧損 | 連3筆 → 部位降至 50% |
| 強制觀望 | EPS≤0 / 日均量<500張 / 外資連賣>10日 |

---

## 五、API 端點快速參考

| 數據 | 端點 |
|------|------|
| 基本面 | `five91.onrender.com/stock/{code}` |
| 日K（上市） | TWSE `STOCK_DAY?date={YM}01&stockNo={code}` |
| 日K（上櫃） | TPEX `afterTrading/tradingStock?date={YY}/{MM}/01&code={code}` |
| 三大法人（上市） | TWSE `T86?date={YYYYMMDD}&selectType=ALLBUT0999` |
| 三大法人（上櫃） | TPEX `insti/dailyTrade?date={YY}/{MM}/{DD}&code={code}` |
| 加權指數 | TWSE `FMTQIK?date={YM}01` |
| 融資餘額 | TWSE `MI_MARGN?date={YYYYMMDD}&selectType=ALL` |

自動偵測市場：先試 TWSE 當月→上月，有資料=上市，否則=上櫃

---

*更新：2026-04-11（v10 Slim — 評分方法論已移至 refs/SwingTrade_Scoring.md，token ~1.5k）*
