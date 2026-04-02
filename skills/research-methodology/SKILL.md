# 企業研究自動化 Pipeline — Cowork Master Instruction

## 使用方式
在 Cowork 中輸入：「請讀取 MASTER_RESEARCH_PIPELINE.md，然後研究 [公司名稱]，角度：[研究角度]，類型：[公司生平/併購/家族/全部]」
Cowork 會按照以下步驟自動執行。全程不要停下來等老闆回覆，做完為止。

## 前置條件
- Claude in Chrome extension 已啟用
- 已登入 Gemini（gemini.google.com）
- 已登入 NotebookLM（notebooklm.google.com）
- 已登入 Perplexity（perplexity.ai）
- 已登入 claude.ai

## 資料夾結構
每次研究新公司時，在指定的工作資料夾下建立：
```
{公司名稱}_{YYYYMMDD}/
├── 01_research_questions/
│   ├── gemini_initial_questions.md
│   ├── perplexity_gap_review.md
│   └── final_questions.md
├── 02_deep_research/
│   ├── gemini_round1_batch1.md
│   ├── gemini_round1_batch2.md
│   ├── gemini_round1_batch3.md
│   ├── nlm_round1_batch1.md
│   ├── nlm_round1_batch2.md
│   ├── nlm_round1_batch3.md
│   └── supplementary/
├── 03_cross_check/
│   ├── perplexity_gap_analysis.md
│   └── supplementary_questions.md
├── 04_structured_report/
│   ├── module_01_xxx.md
│   ├── module_02_xxx.md
│   └── merged_full_report.md
├── 05_final_output/
│   └── {公司名稱}_report.docx
└── sources/
    └── all_sources.md
```

---

## Step 1：生成 Research Questions（Gemini web UI）

### 操作流程
1. 用 Chrome 打開 gemini.google.com
2. 開新對話
3. 根據研究類型，從 prompts/ 資料夾讀取對應的 prompt template
4. 把以下 prompt paste 進 Gemini：

```
我正在從零開始研究 {公司名稱}。
研究角度是：{研究角度}
請幫我生成 20-25 個深度研究問題，涵蓋以下面向：
- 企業起源與創辦人背景
- 歷史關鍵決策點
- 競爭格局與護城河
- 企業治理與決策文化
- 財務體質與資本配置
- 重大併購與失敗案例
- 當前戰略與未來風險
請按面向分類列出，每個問題都要具體到可以直接拿去做 Deep Research。
```

5. 等 Gemini 回覆完畢
6. 複製全部 output，存到 01_research_questions/gemini_initial_questions.md

### 如果是併購分析類型，追加這段：
```
另外請針對 {買方公司} 收購 {目標公司} 的交易，生成 10 個專門的併購調查問題，涵蓋：
- 賣方家族背景與出售動機
- 交易幕後推手與財務顧問
- 估值爭議與市場反應
- 退場機制與停損分析
```

### 如果是家族研究類型，追加這段：
```
另外請針對 {家族名稱} 的家族史，生成 10 個專門問題，涵蓋：
- 家族世系結構與婚姻聯盟
- 財富傳承機制（信託、基金會、遺囑）
- 家族成員的個人命運與爭議
- 家族企業的治理權演變
```

---

## Step 1b：AI 交叉驗證（Perplexity web UI）

### 操作流程
1. 用 Chrome 打開新 tab，進入 perplexity.ai
2. 把 Step 1 的 questions 全部 paste 進去，前面加上：

```
以下是我準備用來研究 {公司名稱} 的 research questions。
請以資深商業分析師的角度檢視這些問題：
1. 有沒有重要的 blind spot（盲點）被遺漏？
2. 哪些問題的切入角度可以更銳利？
3. 請補充 5-8 個你認為不可或缺但上面沒有的問題。
```

3. 等 Perplexity 回覆
4. 複製 output，存到 01_research_questions/perplexity_gap_review.md
5. 回到 Gemini tab，把 Perplexity 的補充意見 paste 給 Gemini：

```
以下是另一位分析師對我 research questions 的補充意見：
{paste Perplexity 的 output}

請整合原始問題與補充意見，產出一份最終版的 research questions list，
去除重複、合併相似問題、按優先順序排列。
```

6. 複製 Gemini 的 final output，存到 01_research_questions/final_questions.md

---

## Step 2：Deep Research（Gemini + NotebookLM 平行處理）

### 分配批次
把 final_questions.md 裡的問題分成 6 批：
- Batch G1, G2, G3 → Gemini Deep Research
- Batch N1, N2, N3 → NotebookLM Deep Research

### 操作流程（同時跑，不要一個一個等）
1. Chrome Tab 1-3 → gemini.google.com → 各自啟動 Deep Research → 各 paste 一批問題 → 按開始研究
2. Chrome Tab 4-6 → notebooklm.google.com → 各自建新 notebook → 各 paste 一批問題 → 跑 Deep Research
3. 6 個同時跑。每 3-5 分鐘 check 一次進度。
4. 哪個先完成就先提取 output
5. 存到 02_deep_research/ 對應的檔案

### 極度重要：逐筆提取原始來源（每一批都要做）

每個 Deep Research 完成後，必須提取原始來源。

**Gemini 的操作：**
Gemini Deep Research output 底部有「資料來源」列表。
逐筆複製每個來源的標題和 URL。

**NotebookLM 的操作：**
NotebookLM output 中每段事實旁有 source indicators。
點擊每個 source indicator，找到原始來源的標題和 URL。

**存到 sources/all_sources.md，格式：**
```
[1] 來源機構或作者. (年份). 文章標題. URL
[2] 來源機構或作者. (年份). 文章標題. URL

範例：
[1] Duke Homestead. (n.d.). Duke family history. https://dukehomestead.org/duke-family-history/
[2] 工商時報. (2022年3月13日). 台泥張安平企業界最佳救援投手. https://www.ctee.com.tw/news/20220313700107-430505
[3] Durden, R. F. (1975). The Dukes of Durham, 1865-1929. Duke University Press.
```

**目標：每份 Deep Research output 至少提取 15-25 個獨立原始來源。**
**全部 6 批合計至少 80 筆來源。**

### 分批處理
- 每批之間等待 1-2 分鐘，避免觸發 rate limit
- 每一批都要獨立提取來源，不能偷懶只記 batch 編號

---

## Step 3：Cross-check 找 Gap（Perplexity web UI）

### 操作流程
1. 回到 Perplexity tab（或開新 tab）
2. 把 Step 2 所有 rounds 的 output 合併，paste 給 Perplexity：

```
以下是我透過 Deep Research 得到的關於 {公司名稱} 的研究報告。
請以資深分析師角度檢視：

1. 這份研究在哪些面向的覆蓋度不足？
2. 有哪些重要的歷史事件、關鍵人物、或決策被遺漏？
3. 財務數據是否完整？缺少哪些年份或指標？
4. 請列出 5-8 個補充研究問題，針對上述 gaps。
```

3. 等 Perplexity 回覆
4. 存到 03_cross_check/perplexity_gap_analysis.md
5. **Perplexity 的回覆自帶 inline citations → 也提取到 sources/all_sources.md**
6. 如果有重大 gaps，把補充問題存到 03_cross_check/supplementary_questions.md
7. 回到 Step 2，用補充問題再跑一輪 Deep Research
8. 這個 loop 最多跑 3 輪，直到 Perplexity 確認覆蓋度足夠

---

## Step 4：Structure 跟梳理（claude.ai web UI）

### 操作流程
1. 用 Chrome 打開新 tab，進入 claude.ai
2. 開一個新的 Project（如果有 Project 功能），命名為「{公司名稱}_Report」
3. 按 module 分批整理。每個 module 一個獨立的對話。

### Module 處理順序
讀取 prompts/ 資料夾中對應的 prompt template，按以下順序處理：

**如果類型是「公司生平分析」或「全部」：**
- Module 1：企業基本定位 + 企業史與戰略演化
- Module 2：關鍵決策點分析
- Module 3：競爭格局與行業定位
- Module 4：企業治理與決策文化
- Module 5：財務健康快照
- Module 6：當前戰略與未來展望

**如果類型是「併購分析」或「全部」：**
- Module 7：併購案全貌與買方戰略動機
- Module 8：賣方背景深度調查

**如果類型是「家族研究」或「全部」：**
- 按家族世系結構分 module

### 每個 Module 的操作
1. 在 claude.ai 開新對話
2. Paste 對應的 prompt template（從 prompts/ 資料夾讀取）
3. 把該 module 相關的 raw research 段落 paste 進去
4. 同時 paste 對應的來源（從 sources/all_sources.md 中選取相關的）
5. 加上這段指令：

```
請基於以上原始研究資料，按照 prompt 的框架要求整理這個模組。

【Footnote 格式要求】

一、每一筆事實性陳述（數字、日期、引述、事件描述）都必須在句尾標上腳註編號。
    格式為：「......事實描述。[1]」
    
二、不可使用 generic 來源如「NotebookLM Round 1」或「Perplexity分析」。
    每一個腳註必須指向一個獨立的原始來源。
    
三、腳註放在該模組的末尾，格式為：
    [1] 來源機構或作者. (年份). 文章標題. URL
    [2] 來源機構或作者. (年份). 文章標題. URL
    
    範例：
    [1] Duke Homestead. (n.d.). Duke family history. https://dukehomestead.org/duke-family-history/
    [2] 工商時報. (2022年3月13日). 台泥張安平企業界最佳救援投手. https://www.ctee.com.tw/news/20220313700107-430505

四、我提供的原始資料中已包含來源資訊，請直接使用。
    如果某段事實的來源不明確，標記為 [待確認來源]。

五、全程繁體中文，純文字輸出，不使用任何 Markdown 語法。
    用適當的換行與全形標點符號排版。
    展現決策者視角、強化因果邏輯、進行替代方案比較。
```

6. 等 Claude 回覆完畢
7. 檢查 output 的 footnotes：
   - 是否每個關鍵事實都有獨立腳註？
   - 腳註是否指向原始來源（不是 generic batch 編號）？
   如果不合格，追問：「你的腳註太粗略了。請重新標註，確保每一筆事實都有獨立的原始來源腳註。」
8. 複製合格的 output，存到 04_structured_report/module_XX_{模組名}.md

### 合併與腳註統一
所有 module 完成後：
1. 把所有 module 檔案按順序合併成 merged_full_report.md
2. 重新統一腳註編號（避免不同 module 之間編號重複）
3. 在文件末尾建立完整的「參考來源」清單
4. 搜尋全文：如果「NotebookLM」「Perplexity」「研究內部文件」出現在腳註中，必須修正

---

## Step 5：合併輸出 merged markdown（不要產 .docx）

### 操作流程
1. 把 04_structured_report/ 裡所有 module 檔案按順序合併
2. 合併成一個完整的 merged_full_report.md
3. 腳註全文重新統一編號
4. 搜尋全文：如果「NotebookLM」「Perplexity」「研究內部文件」出現在腳註中，必須修正
5. 存到 04_structured_report/merged_full_report.md

### 不要產 .docx
Cowork 產的 .docx 排版永遠有問題。不要浪費時間。
只產 markdown 就好。.docx 排版由人工在 Word 裡面完成。

---

## Step 6：Human Review（手動）
此步驟無法自動化，由 Demos 或團隊成員手動完成：
- 通讀全文，標記有故事性的亮點段落
- 調整敘事節奏與邏輯流暢度
- 確認 footnotes 準確性
- 最終排版微調

---

## 異常處理（自己解決，不要停）

### 網站打不開
等 30 秒重試。連續 3 次失敗就跳過先做下一步，回頭補。

### Deep Research 跑太久（>20 分鐘）
重新提交。重試 2 次還是超時就跳過這一批。

### Claude.ai 回覆被截斷
輸入「請繼續」。多次截斷就把 module 拆更小。

### 任何問題
不要停下來等老闆。自己想辦法解決，做完為止。
老闆可能在睡覺。他希望醒來時看到完成的報告。
