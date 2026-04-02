---
name: network-report-generator
description: 從人物姓名自動搜集資料並生成完整的人脈網絡報告 PPT（18頁）。包含：封面、目錄、個人介紹、朋友圈總覽 Dashboard、核心圈、企業圈、協會與政商圈、政策與產業推手圈、學術校友圈、思想與論壇引路人圈、媒體與公開平台圈、迴避對象與引薦建議、家族樹、核心成員介紹（多頁）、家族企業版圖、總結。每頁含來源出處與頁碼。
---

# 人脈網絡報告 PPT 一鍵生成器

## 總覽

輸入一個人物姓名+公司名稱，自動完成：
1. **用瀏覽器操作外部 AI 平台做 Deep Research**（Gemini + Perplexity）
2. **整合多平台研究結果，交叉驗證**
3. **整理成結構化 JSON**
4. **18 頁 PPT 自動生成**（python-pptx）

## 安裝依賴

```bash
pip install python-pptx --break-system-packages
```

## 工作流程（必須按順序執行，不得跳過 Step 1）

### Step 1：Deep Research（用瀏覽器操作外部 AI）

⚠️ **這一步是必須的，不要跳過，不要只用自己的 web_search 代替。**
⚠️ **你有 Claude in Chrome 的瀏覽器控制權限，直接開網頁操作。**

讀取 `templates/research_prompt.md` 取得 prompt template，將 {TARGET_NAME}、{COMPANY_NAME}、{STOCK_CODE} 替換成目標人物的資料後，依序執行：

**平台 1：Gemini**
1. 用瀏覽器打開 https://gemini.google.com
2. 開新對話
3. 貼上替換好的 research prompt
4. 等 Gemini 跑完（可能需要 1-3 分鐘）
5. 把 Gemini 的完整回覆記錄下來

**平台 2：Perplexity**
1. 用瀏覽器打開 https://www.perplexity.ai
2. 開新對話
3. 貼上同一段 research prompt
4. 等 Perplexity 跑完
5. 把 Perplexity 的完整回覆記錄下來

**平台 3（補充）：自己的 web_search**
用自己的 web_search 補充搜集 Gemini 和 Perplexity 可能遺漏的資料，特別是：
- 公開資訊觀測站的董監持股
- 關聯投資公司的代表人
- 最近 3 個月的新聞

### Step 2：整合研究結果

把 Gemini + Perplexity + 自己 web_search 的結果合併，交叉驗證。原則：
- 三個來源都提到的 → 直接採用
- 兩個來源提到的 → 採用但標註來源
- 只有一個來源提到的 → 採用但標註「資訊待確認」
- 沒有任何來源 → 標註「N/A」

搜集的 8 大類資料：
1. 個人介紹（學歷、職涯 timeline 15-25 個里程碑、經營哲學語錄 3-5 句、持股）
2. 朋友圈 8 類人脈（每類 3-8 人，每人含姓名、現職、關係🟢🔴🟡⚪、互動脈絡 20-40 字、來源 URL）
3. 家族關係人（G0-G3，四親等內，每人含姓名、關係、出生年份、學歷、現職）
4. 家族企業版圖（關聯企業 + 投資控股公司 + 持股比例）
5. 媒體事件（正面/負面各 5 筆以上，含時間、媒體名、摘要）
6. 迴避名單 + 引薦建議
7. 家族樹結構（用於 tree_data 自動生成家族樹圖）
8. 總結分析（正面觀察 5 點、風險關注 5 點、建議追蹤 5 點、關鍵結論 100 字）

### Step 3：整理成 JSON

整理成 `references/data_schema.json` 定義的格式。三個硬性要求：

**A. 來源欄位必須用結構化格式：**
```json
"source": [
  {"title": "文章完整標題", "url": "https://完整網址"},
  {"title": "另一篇文章標題", "url": "https://完整網址"}
]
```
不要用純文字（如 "來源：經濟日報、天下雜誌"）。每筆都要有完整文章標題和可點擊的 URL。

**B. family.tree_data 必須填入：**
```json
"tree_data": [
  {"generation": "G0", "members": [
    {"name": "父親姓名", "relation": "父親", "title": "職稱或N/A", "gender": "M"},
    {"name": "母親姓名", "relation": "母親", "title": "職稱或N/A", "gender": "F"}
  ]},
  {"generation": "G1", "members": [
    {"name": "主角姓名", "relation": "本人", "title": "主要職稱", "gender": "M"},
    {"name": "配偶姓名", "relation": "配偶", "title": "職稱或N/A", "gender": "F"},
    {"name": "兄弟姓名", "relation": "弟弟", "title": "職稱", "gender": "M"}
  ]},
  {"generation": "G2", "members": [...]}
]
```
找不到的人也要列，name 標 N/A。

**C. overview_cards 8 張卡片全部要填。**

### Step 4：生成 PPT

```bash
python scripts/generate_network_report.py --input data.json --output /mnt/user-data/outputs/{人名}_人脈網絡報告.pptx
```

## PPT 架構（18 頁）

| Slide | 內容 | 說明 |
|-------|------|------|
| 1 | 封面 | 深藍全屏背景+人名+職稱+標語 |
| 2 | 目錄 | 11 章節，左右分欄排列 |
| 3 | 個人介紹 | 左：基本資料+照片位 / 右：職涯里程碑+經營哲學 |
| 4 | 朋友圈總覽 | 8 個 card 用 grid 排列，每個有 emoji+代表人物 |
| 5 | 核心圈 | 🟢正向/🔴負向分區表格 |
| 6 | 企業圈 | 董事會+經營團隊表格 |
| 7 | 協會與政商圈 | 組織+基金會+政界表格 |
| 8 | 政策與產業推手圈 | 5欄大表格 |
| 9 | 學術校友圈 | 學歷+校友表+引言 |
| 10 | 思想與論壇引路人圈 | 表格+經典語錄框 |
| 11 | 媒體與公開平台圈 | 正面/負面事件分區表格 |
| 12 | 迴避對象+引薦建議 | 左：迴避名單 / 右：引薦管道 |
| 13 | 家族樹 | 自動生成（shapes+連接線，從 tree_data 讀取） |
| 14-16 | 核心成員介紹 | 每位成員 detailed profile（自動分頁） |
| 17 | 家族企業版圖 | 關聯企業+控股公司表格 |
| 18 | 總結 | 正面觀察/風險關注/建議追蹤+關鍵結論 |

## 視覺設計規範

### 色彩系統
```
#1A365D - 深藍（標題欄、主標題）
#C6A052 - 金色（卡片標題 accent）
#2C5282 - 中藍（章節標題）
#38A169 - 綠色（正向指標）
#C53030 - 紅色（負向指標）
#2D3748 - 深灰（正文）
#718096 - 灰色（來源、頁碼）
#78350F - 棕色（語錄文字）
#92400E - 深棕紅（哲學標題）
#A0AEC0 - 淺灰（副標題）
```

### 字體大小
- 封面人名：60pt bold 白色
- 封面職稱：22pt 白色
- 頁面標題：32pt bold 深藍
- 章節小標：15pt bold 中藍
- 正文：12pt 深灰
- 表格內容：12pt
- 來源/頁碼：11pt 灰色
- 最小字體：10pt

### 關係 Emoji
- 🟢 正向/強連結
- 🟡 中性/一般
- 🔴 負向/迴避
- ⚪ 中性/資訊不足

## 輸出

`/mnt/user-data/outputs/{人名}_人脈網絡報告.pptx`
