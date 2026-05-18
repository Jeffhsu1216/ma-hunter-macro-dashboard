# MAContentCreator.md — 每週併購新聞內容製作

> 隸屬：📺 Youtuber（M&A Hunter）
> 版本：v1.0｜2026-04-12

---

## 一、功能說明

將每週併購新聞（`MA_Weekly.md` 輸出 / LINE 紀錄）自動轉換為：
- **YouTube Shorts / Instagram Reels**：HTML 動態預覽（可螢幕錄影發布）
- **Podcast 腳本**（未來擴充）

---

## 二、觸發格式

```
製作shorts [MM/DD-MM/DD]
製作shorts （省略日期 → 自動讀最新一週）
```

---

## 三、數據來源

| 來源 | 說明 |
|------|------|
| `~/Desktop/Claude/併購新聞整理.txt` | LINE 聊天紀錄匯出，含 emoji 標記（📕📙📒📗📘）|
| `MA_Weekly.md` 輸出快取 | 本週最新彙整（若已執行「併購週報」）|

**解析規則：**
- `📕` = 超大型（>100億美元）
- `📙` = 大型（10–100億）
- `📒` = 中型（1–10億）
- `📗` = 小型（<1億）
- `📘` = 台灣在地
- 篩選：僅取台灣（📘）+ 全球（📕📙📒），略過大陸

---

## 四、Shorts 製作流程

### Step 1：解析新聞
```
1. 讀取 ~/Desktop/Claude/併購新聞整理.txt 末尾 200 行
2. 找出最新一週的分隔標記（日期區間）
3. 提取各交易：買方 × 賣方、金額、產業、地區
4. 排序：全球依金額由大到小，台灣依規模
```

### Step 2：選出 Scenes
```
Scene 0:  Hook — 本週最大交易金額 + 反差 hook 句
Scene 1:  全球最大（📕 單筆聚焦）
Scene 2:  第二大或主題交易（PE/科技/生技選一）
Scene 3:  生技/pharma 多筆（若有 2+ 筆則合為 card 式）
Scene 4:  🇹🇼 台灣（1–3 筆 card 式，綠色配色）
Scene 5:  本週觀點（1 句話結語）
Scene 6:  CTA（追蹤 M&A Hunter 固定模板）
```

> Scene 數量依當週交易量調整（最少 5 場、最多 8 場）

### Step 3：生成 HTML
- 輸出至 `~/Desktop/Claude/Youtuber/shorts_preview/index.html`
- 格式：390×693 手機框 + 全螢幕錄影模式
- 背景配色依主題：
  - 全球大型 PE/金融 → 深藍紫
  - 生技/醫療 → 深黑黃
  - 台灣 → 深黑綠
  - 觀點/CTA → 深橙/黑

### Step 4：自動產出 PNG
```bash
python3 ~/Desktop/Claude/Youtuber/shorts_preview/export_scenes.py
```
- 使用 Playwright headless Chrome 截圖
- 輸出：`~/Desktop/Youtuber/Shorts/01_MA_Hunter.png` ～ `07_MA_Hunter.png`
- 解析度：1080×1920（Instagram / YouTube Shorts 標準）
- 完成後預覽第 1 張確認品質

### Step 5：回報產出摘要
```
✅ 7 張 PNG 已存至 ~/Desktop/Youtuber/Shorts/
📱 直接上傳 IG 限時動態 / LINE 貼文即可
```

---

## 五、HTML 結構規範

### 固定元素（每週不變）
- 進度條（頂部 7 格）
- 頭像 `M` + 頻道名 `M&A Hunter` + 追蹤按鈕
- 右側 sidebar（❤️ 留言 分享 更多）
- 底部標題列（`M&A HUNTER · 週報`）
- 全螢幕錄影模式 + 倒數計時

### 每週更換
- 日期區間（標題、Scene 0 標籤）
- 各 Scene 的數字、公司名稱、故事角度
- Hook 句（第 0 場反差文案）
- 本週觀點（第 5 場結語）

### 配色系統
```
主色：#C9A84C（金色）
背景 0（Hook）：深藍黑    radial-gradient(#1a1a3a, #050510)
背景 1（大型）：深紫黑    radial-gradient(#1a0e2e, #060008)
背景 2（PE）：  深青黑    radial-gradient(#0e1a1a, #020c10)
背景 3（生技）：深黃黑    radial-gradient(#1a1400, #0a0800)
背景 4（台灣）：深綠黑    radial-gradient(#001a0a, #000a04)
背景 5（觀點）：深橙黑    radial-gradient(#1a0a00, #0a0400)
背景 6（CTA）： 純黑      radial-gradient(#111111, #000)
台灣accent：   #69DB7C（綠色）
警示色：       #FF6B6B（紅色）
```

---

## 六、Podcast 腳本（未來擴充）

> 預計架構：
> - 開場（30 秒）→ 本週數字鉤子
> - 全球主題交易解析（各 90 秒）
> - 台灣交易（60 秒）
> - 本週觀點（60 秒）
> - 結語 + CTA（30 秒）

---

## 七、注意事項

1. **不製作大陸相關交易的 Shorts**（避免敏感議題）
2. Hook 句要有反差感（關稅危機 vs 資本不停、熊市 vs 大型 LBO）
3. 數字要具體（640億 > 幾百億）
4. 每週最後一 Scene 固定為 CTA，不改變
5. 全螢幕模式使用 `vw/vh` 相對單位自動縮放

---

*最後更新：2026-04-12（v1.0 — 首版，Shorts 功能完成）*
