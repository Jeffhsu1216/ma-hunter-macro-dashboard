---
name: 每日股價更新
description: 每日台灣收盤後（15:00）自動更新所有客戶 Excel 與月選股工作表的收盤價
type: automation
---

# 每日股價更新｜DailyPriceUpdate

> 收盤後自動抓取台股收盤價，更新客戶持倉損益與月選股漲幅計算。

---

## 一、觸發方式

| 方式 | 說明 |
|------|------|
| **自動排程** | 每日 15:05 台北時間（收盤後 5 分鐘） |
| **手動執行** | `python3 ~/Desktop/Claude/個人基金/daily_price_update.py` |
| **Claude 觸發詞** | `更新收盤價` / `每日收盤更新` |

---

## 二、腳本路徑

```
~/Desktop/Claude/個人基金/daily_price_update.py
```

---

## 三、更新範圍

### 客戶 Excel（`~/Desktop/Clients/`）
- 工作表：`Investment Portfolio`
- **I3**：更新標題（如 `4/8收盤價` → `4/9收盤價`）
- **I4 起**：依 D 欄股票代號更新各列收盤價
- **存檔方式**：直接更新為今日日期檔名（如 `靜怡_20260408.xlsx` → `靜怡_20260409.xlsx`），**舊檔自動刪除**
- 每位客戶只保留一個最新日期的 xlsx

| 客戶 | 持股數 | 備註 |
|------|--------|------|
| Jeff | 6 支 | 個人帳戶 |
| 林峻毅 | 5 支 | — |
| 靜怡 | 3 列（2 支不重複） | I6 為公式 `=I5`，自動覆寫 |
| 小阿姨 | 0（全現金） | 僅更新 I3 標題 |
| Gary | 動態 | 自動掃描 Clients/ 目錄，有 xlsx 即納入 |

> ⚠️ **改名規則（所有客戶統一）**：每次執行後，客戶 xlsx 自動改名為當天日期（如 `Gary_20260409.xlsx` → `Gary_20260410.xlsx`），舊檔刪除，每人只保留一個最新日期檔案。

### Jeff_Stock Analysis（`~/Desktop/Stock Analysis/Jeff_Stock Analysis.xlsx`）
- 工作表：自動偵測**最新月份**（如 `202604`）
- **H1**：更新標題
- **H2~H43**：依 B 欄代號更新各股收盤價（最多 40 支）
- **H44**：台灣加權指數（TAIEX）
- **存檔方式**：覆寫原檔（不建新檔）

---

## 四、股價來源策略

```
主來源：yfinance 批次下載（一次 HTTP，所有代號同時取得）
  Step 1：全部加 .TW → yf.download() 一次
  Step 2：失敗的補一次 .TWO → yf.download() 一次
  TAIEX：^TWII 與股票同批下載
```

**效率說明：**
- 兩次 batch download（.TW + .TWO）取代逐支查詢
- 41 支股票 + TAIEX，約 **15 秒**完成
- 台股上市/上櫃自動識別，無需手動設定

> five91 已移除（更新較慢，改用 yfinance 即時收盤價）

---

## 五、執行輸出範例

```
📅 每日股價更新 v4 | 4/9收盤價
=============================================
📋 掃描到 41 個代號
📡 yfinance 批次下載 42 支（含TAIEX）...
  🔄 25 支試 .TWO：[6568, 6291, ...]
  ✅ 取得 41 支股票｜TAIEX = 34861.16

【客戶 Excel】
  ✅ Jeff：7 支持股 → Jeff_20260409.xlsx（舊檔已移除）
  ✅ 小阿姨：0 支持股 → 小阿姨_20260409.xlsx（舊檔已移除）
  ✅ 林峻毅：5 支持股 → 林峻毅_20260409.xlsx（舊檔已移除）
  ✅ 靜怡：3 支持股 → 靜怡_20260409.xlsx（舊檔已移除）

【Jeff_Stock Analysis】
  ✅ Jeff_Stock Analysis [202604]：40 支｜TAIEX=34861.16

✅ 完成
```

---

## 六、排程設定（本機 cron）

> ⚠️ 腳本需存取本機 Excel 檔案，**必須用本機 cron**，不可用 Claude Remote Trigger（雲端無法讀本機檔案）。

**已設定的 crontab 條目：**
```
5 7 * * 1-5 /Users/jeffhsu/anaconda3/bin/python3 /Users/jeffhsu/Desktop/Claude/個人基金/daily_price_update.py >> /Users/jeffhsu/Desktop/Claude/個人基金/price_update.log 2>&1
```

| 欄位 | 說明 |
|------|------|
| `5 7` | UTC 07:05 = **台北時間 15:05** |
| `1-5` | 週一至週五 |
| `>> ... 2>&1` | stdout + stderr 寫入 log |

```bash
crontab -l          # 查看排程
crontab -e          # 編輯排程
tail -50 ~/Desktop/Claude/個人基金/price_update.log   # 查看執行紀錄
```

---

## 七、注意事項

1. **收盤前執行**：yfinance 回傳最後交易日收盤，未收盤前請勿手動觸發
2. **假日**：台股休市日 cron 仍執行，寫入最後交易日資料（無害）
3. **新增持股**：客戶買進新股後，D 欄填入代號即可，下次執行自動更新
4. **上市 vs 上櫃**：腳本先試 `.TW`，失敗自動補 `.TWO`，無需手動設定
5. **檔案管理**：每次執行後各客戶只保留一個最新日期 xlsx，舊檔自動刪除
6. **嚴禁修改**：腳本不修改 C25（現金餘額）等非收盤價欄位

---

*最後更新：2026-04-09（v4 — 舊檔自動刪除，每客戶只保留今日日期 xlsx）*
