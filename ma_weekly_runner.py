"""
每週併購週報自動化 Runner
- 每週日 14:00 台北時間由 GitHub Actions 觸發
- 透過 Claude API (claude-opus-4-6) + web_search 工具搜尋 M&A 新聞
- 結果推送至 Telegram，同時存入 ma_weekly_log/YYYYMMDD.txt
"""

import os
import json
import subprocess
import urllib.request
import urllib.parse
from datetime import datetime, timedelta

import anthropic
import pytz

# ══════════════════════════════════════════
TOKEN   = '8743919766:AAG6z6YPW7Gqt7rF2KY2xC9mvbm2Ge31tjQ'
CHAT_ID = '2117347781'
TAIPEI_TZ = pytz.timezone('Asia/Taipei')
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(SCRIPT_DIR, 'ma_weekly_log')
# ══════════════════════════════════════════

MA_WEEKLY_INSTRUCTIONS = """
## 輸出格式

三個區域（台灣、大陸、全球），每區最多 5 則，依交易金額由大到小排序：

```
🌎 台灣併購投資新聞
📆 MM/DD-MM/DD

📕 {買方}斥資{金額}{幣別}收購{標的} {補充說明}
{短網址}

📙 ...
📒 ...
📗 ...
📘 ...


🌎 大陸併購投資新聞
📆 MM/DD-MM/DD

📕 ...（最多5則）


🌎 全球併購投資新聞
📆 MM/DD-MM/DD

📕 ...（最多5則）
```

Emoji 色碼：📕第1大 / 📙第2大 / 📒第3大 / 📗第4大 / 📘第5大

---

## 搜尋策略

### 台灣（通常 2–4 則）
1. ctee.com.tw（工商時報）— 逐日搜尋週期內每一天
2. money.udn.com（經濟日報）
3. Google：`台灣 併購 收購 {日期範圍}`，翻到第 15–20 頁
- 關鍵詞：併購、收購、合併、股權交易、換股、合併核准、生技收購
- 只收本週宣布的交易；金額後面必須加幣種（新台幣/美元）

### 大陸（通常 3–5 則）
1. WebSearch `site:aastocks.com 收購 OR 併購 {YYYY年MM月}` → 找文章 URL → WebFetch 讀取
2. WebSearch `site:news.futunn.com 收購 {YYYY年MM月}`
3. hk.investing.com、cnyes.com 補充
4. Google：`中國 收購 併購 {年月}`
- 必須搜尋港股（HK$計價）交易，純 A 股新聞會漏掉大量港股
- 排除「金額未揭露」交易

### 全球（取最大 5 則，先廣泛搜集再排序）
1. `largest M&A deals "{日期}" billion {年}` — 掌握本週最大交易
2. `PE buyout LBO acquisition billion announced {年月}` — ⚠️ 最容易漏，金額最大
3. `pharma biotech oncology acquisition billion "{日期}" {年}` — 筆數最多
4. `technology software acquisition billion announced {年月}`
5. ctee.com.tw 國際版、money.udn.com、cnyes.com

必搜產業：PE槓桿收購、生技/製藥、科技/軟體、基礎設施/能源、消費品

---

## 品質規則
- ⚠️ 所有交易必須在本週（{date_range}）宣布，不收前幾週的舊交易
- ⚠️ 全球最終 5 則必須是本週宣布中金額最大的 5 筆（找到 5 則 ≠ 停止，要確認沒有更大的漏掉）
- ⚠️ 大陸必須跑 AASTOCKS，港股交易只在此可找到
- URL 超過 50 字元 → 用 https://tinyurl.com/api-create.php?url={url} 縮短
- 標題格式：{買方}斥資{金額}{幣別}收購{標的} {補充說明}
- 金額統一用「億」為單位
- 排除傳言、金額未揭露的交易
"""


def get_date_range():
    """計算本週一到今日（週日）的日期範圍"""
    now = datetime.now(TAIPEI_TZ)
    # 週日 weekday() = 6，往回 6 天 = 週一
    monday = now - timedelta(days=now.weekday() if now.weekday() != 6 else 6)
    start = monday.strftime('%m/%d')
    end   = now.strftime('%m/%d')
    year  = now.strftime('%Y')
    return start, end, year


def shorten_url(url):
    """TinyURL 縮短，失敗回傳原始 URL"""
    if len(url) <= 50:
        return url
    try:
        api = f'https://tinyurl.com/api-create.php?url={urllib.parse.quote(url, safe="")}'
        req = urllib.request.Request(api, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as r:
            short = r.read().decode('utf-8').strip()
        return short if short.startswith('http') else url
    except Exception:
        return url


def run_claude(prompt: str) -> str:
    """呼叫 Claude API + web_search 工具，回傳最終文字"""
    client = anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])

    messages = [{"role": "user", "content": prompt}]

    for _ in range(40):  # 最多 40 輪工具呼叫
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=16000,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=messages,
        )

        # 把 assistant 回覆加入對話
        messages.append({"role": "assistant", "content": response.content})

        if response.stop_reason == "end_turn":
            return "".join(
                block.text for block in response.content
                if hasattr(block, 'text') and block.type == "text"
            )

        if response.stop_reason == "tool_use":
            # 收集工具結果，回傳給 Claude 繼續處理
            user_content = []
            for block in response.content:
                if hasattr(block, 'type') and block.type == "tool_result":
                    user_content.append({
                        "type": "tool_result",
                        "tool_use_id": block.tool_use_id,
                        "content": block.content,
                    })
            if user_content:
                messages.append({"role": "user", "content": user_content})

    return "[錯誤] 超過最大工具呼叫次數"


def send_telegram(text: str) -> bool:
    """推送純文字至 Telegram"""
    payload = json.dumps({
        'chat_id': CHAT_ID,
        'text': text,
        'disable_web_page_preview': False,
    }).encode('utf-8')
    req = urllib.request.Request(
        f'https://api.telegram.org/bot{TOKEN}/sendMessage',
        data=payload,
        headers={'Content-Type': 'application/json'},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            result = json.loads(r.read())
        return result.get('ok', False)
    except Exception as e:
        print(f'Telegram 推送失敗：{e}')
        return False


def save_and_push(text: str, date_str: str):
    """存檔 + git push"""
    os.makedirs(LOG_DIR, exist_ok=True)
    filepath = os.path.join(LOG_DIR, f'{date_str}.txt')
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(text)
    try:
        subprocess.run(['git', '-C', SCRIPT_DIR, 'add', 'ma_weekly_log/'], check=True)
        subprocess.run(['git', '-C', SCRIPT_DIR, 'commit', '-m',
                        f'[auto] 併購週報 {date_str}'], check=True)
        subprocess.run(['git', '-C', SCRIPT_DIR, 'push'], check=True)
        print(f'✅ 已存檔並推送：ma_weekly_log/{date_str}.txt')
    except subprocess.CalledProcessError:
        print('⚠️  git push 失敗（可能無異動）')


def main():
    start, end, year = get_date_range()
    date_range = f'{start}–{end}'
    date_str = datetime.now(TAIPEI_TZ).strftime('%Y%m%d')

    print(f'📅 執行週報：{year} {date_range}')

    prompt = f"""你是 M&A Hunter 的資深投資經理助手，專精於全球併購新聞彙整。

請執行本週（{year}/{date_range}）的每週併購週報，搜尋台灣、大陸、全球三大區域的 M&A 新聞。

{MA_WEEKLY_INSTRUCTIONS.replace('{date_range}', date_range)}

**重要**：直接輸出完整週報，不加任何前言說明。
"""

    print('🤖 呼叫 Claude API（可能需要 2–5 分鐘）...')
    result = run_claude(prompt)

    if not result or result.startswith('[錯誤]'):
        print(f'❌ Claude 執行失敗：{result}')
        return

    print('\n=== 週報內容 ===')
    print(result)
    print('================\n')

    ok = send_telegram(result)
    print('✅ Telegram 推送成功' if ok else '❌ Telegram 推送失敗')

    save_and_push(result, date_str)


if __name__ == '__main__':
    main()
