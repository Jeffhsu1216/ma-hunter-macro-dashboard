# Thank_You_Letter.md｜法說會感謝信工作規範（已優化）

> 當使用者輸入「{股票代號} {公司名稱} {稱謂} 感謝信」時載入。
> 自動查詢財務數據與新聞，填入第一段，第二、三段照抄固定模板。

---

## 一、觸發格式

```
{股票代號} {公司名稱} {稱謂} 感謝信
範例：8103 瀚荃股份 楊董事長 感謝信
```

解析規則：
- **股票代號**：用於 API 查詢
- **公司名稱**：用於信件稱呼與內文
- **稱謂**：格式為「{姓氏}{職銜}」，作為信件開頭敬稱（如：楊董事長、王總經理）

---

## 二、執行流程（已優化版）

### Phase 1：1 支 curl（目標 ≤ 2s）

```bash
curl -s "https://five91.onrender.com/api/metrics?stock_id={代號}" > /tmp/{代號}_metrics.json
```

> ✅ metrics 已包含 `revenue_ttm`、`gross_margin`、`eps_ttm`，**不需再抓 income_statement HTML**

---

### Phase 2：Python JSON 解析（目標 ≤ 3s）

```python
import json
from pathlib import Path

metrics = json.loads(Path(f'/tmp/{代號}_metrics.json').read_text())

name          = metrics.get('name', '')           # 公司名稱（驗證用）
category      = metrics.get('category', '')       # 產業類別
revenue_ttm   = metrics.get('revenue_ttm', 0)     # TTM 營收（元）
gross_margin  = metrics.get('gross_margin', 0)    # 毛利率（小數，如 0.32）
eps_ttm       = metrics.get('eps_ttm', 0)         # TTM EPS

# 格式化
revenue_100M  = round(revenue_ttm / 1e8, 2)       # 轉億元
gm_pct        = round(gross_margin * 100, 1)      # 毛利率 %
```

---

### Phase 3：單次 WebSearch（目標 ≤ 15s）

```
查詢詞："{公司名稱} 法說會 法人說明會 {最近年份} 論壇 營收成長"
目標（一次搜尋取得全部）：
  ① 法說會論壇名稱（如凱基論壇、元大論壇）
  ② 法說會舉辦日期
  ③ 最近完整年度 YoY 營收成長率（%）
  ④ 任何一句營運亮點或挑戰（供第一段補充句使用）

查無論壇名稱 → 填「法人說明會」
查無日期     → 填「近期」
查無 YoY     → 改寫為「維持穩健成長動能」，不填具體數字
```

---

### Phase 4：撰寫並輸出信件

- 填入第一段變數（見第三節）
- 貼上固定第二、三段
- 直接輸出完整信件，**不加任何前言說明**

---

### Phase 5：自動 Append 至公司研究追蹤（目標 ≤ 90s）

```python
from pathlib import Path
import openpyxl

folder = Path.home() / 'Desktop/Stock Analysis'
files = sorted(folder.glob('Company Overview_*.xlsx'))
latest = files[-1] if files else None

# 檢查股票代號是否已存在
already_exists = False
if latest:
    wb = openpyxl.load_workbook(latest)
    ws = wb['公司追蹤清單']
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[1]) == '{代號}':   # B 欄 = 股票代號
            already_exists = True
            break
```

- **已存在** → 跳過，不重複寫入，靜默略過
- **不存在** → 依 `CompanyOverview_AutoAppend.md` 完整流程 append（4並行curl → Python解析 → WebSearch → 寫入）
- Append 完成後回報：「已將 {代號} {公司名稱} 新增至公司研究追蹤」或「{代號} 已存在，略過 append」

---

## 三、第一段模板與變數說明

```
我們是《圖靈金融集團》的投資團隊，很高興參加貴公司{法說會時間}於{論壇名稱}所舉辦的線下法說會，
讓我們對{公司名稱}的營運狀況與未來發展有了更深入的了解。
首先恭喜貴公司於{財務年度}全年營收達新台幣{營收}億元、年成長{YoY成長率}%，{是否創新高描述}，
展現公司在{主要產業或產品描述}中的穩固市場地位。
{一句話補充：可提毛利率、EPS、或市場布局亮點，也可提波動或挑戰（若有）}，
也讓我們感受到{公司名稱}在{核心競爭力描述}下所累積的產業實力，
進一步加深我們對公司中長期發展方向的信心。
```

**變數填寫規則：**

| 變數 | 來源 | 備註 |
|------|------|------|
| `{法說會時間}` | WebSearch（Phase 3） | 如「上週四」「本月」，查無則省略，直接寫「近期」 |
| `{論壇名稱}` | WebSearch（Phase 3） | 如「凱基論壇」「元大論壇」，查無則寫「法人說明會」 |
| `{公司名稱}` | 使用者輸入 | 直接使用 |
| `{財務年度}` | WebSearch（Phase 3） | 最新完整年度，如「2025 年」；查無則依 metrics 推算 |
| `{營收}` | `metrics.revenue_ttm` | 格式：xxx.xx（億元） |
| `{YoY成長率}` | WebSearch（Phase 3） | 格式：+x.x 或 -x.x（百分比）；查無則省略改用描述句 |
| `{是否創新高描述}` | 判斷 | 若歷史新高 → 「再創歷史新高」；否則 → 「展現穩健成長動能」 |
| `{主要產業或產品描述}` | `metrics.category` | 如「聲學與電聲產品」「精密連接器製造」 |
| `{一句話補充}` | `metrics.gross_margin` / `metrics.eps_ttm` + WebSearch | 選用毛利率、EPS 亮點，或如實描述挑戰 |
| `{核心競爭力描述}` | WebSearch（Phase 3） | 如「技術深化與全球客戶基礎」「垂直整合與供應鏈布局」 |

---

## 四、固定段落（第二段、第三段）

> ⚠️ 以下兩段照抄，**不做任何修改**。

**第二段：**
《圖靈金融集團》目前管理兩檔AI量化基金與兩檔併購基金。在勤業眾信近期發布的《2025台灣私募股權基金白皮書》中，特別提及 "圖靈金融為台灣少數結合量化工具與私募投資的機構"，並指出我們 "結合金融運用AI演算法、大數據分析與雲端運算，提供智能、創新及整合的併購投資與量化交易金融服務給予台灣上市上櫃公司"。

**第三段：**
《圖靈金融集團》長期服務多家台灣上市上櫃公司，涵蓋半導體、電子、化學、營建等不同產業。其中包含：聖暉工程、強茂集團、南寶樹脂、中鼎集團、中國砂輪、世豐世鎧等具規模企業。我們希望在不久的將來，由張堯勇董事長率領投資團隊親自前往貴公司拜訪，進行更深入的交流，分享我們在併購與量化投資上的專業與案例。

---

## 五、完整輸出格式

```
{稱謂} 您好：

{第一段（客製化）}

{第二段（固定）}

{第三段（固定）}

《圖靈金融集團》敬上
```

---

## 六、注意事項

- 全程繁體中文，語氣專業恭賀
- 第一段不超過 250 字
- 財務數字四捨五入至小數點後兩位
- 若 API 數據取得失敗，改以 WebSearch 搜尋公開財報資訊補充
- 信件直接輸出，**不加任何前言或說明**（如「以下是信件內容」等）
- **禁止使用 WebFetch 取代 curl**：WebFetch 會過一層 AI 模型，浪費 token 且易取錯數據

---

## 七、效能規範

> 實測基準目標：≤ 20s / ≤ 1,000 tokens

| Phase | 內容 | 目標耗時 |
|-------|------|---------|
| Phase 1 | **1 支 curl**（metrics，含 revenue_ttm / gross_margin / eps_ttm） | ≤ 2s |
| Phase 2 | Python JSON 直接解析，無需 HTML parsing | ≤ 3s |
| Phase 3 | **單次 WebSearch**（論壇名稱 + 日期 + YoY + 亮點，一次取得） | ≤ 15s |
| Phase 4 | 撰寫第一段 + 貼入固定段落 + 輸出信件 | ≤ 3s |
| Phase 5 | 檢查公司是否已在追蹤清單 → 不存在則 append | ≤ 90s |

**優化說明（v2 → v3）：**
- ❌ 移除 `income_statement` HTML curl（原 Phase 1 第二支）
- ❌ 移除 Python HTML table parsing（原 Phase 2 重頭戲）
- ✅ metrics API 已直接提供 TTM 營收、毛利率、EPS，無需重算
- ✅ YoY 改由 WebSearch 一併取得，合併為單次查詢
- 💡 總耗時從 ~30s 縮短至 ~20s，token 消耗減少約 35%

---

*最後更新：2026-04-06（v3 已優化 — 1 curl、無 HTML parse、單次 WebSearch + Phase 5 自動 append 公司研究追蹤）*
