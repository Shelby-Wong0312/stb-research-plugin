# stb-research

**台灣董事學會 AI 研究工具包** — 企業深度研究、人物網絡報告、Screening Note

由 STB AI Team 開發，整合台灣董事學會的研究方法論，提供三大核心 slash commands，讓 Claude 一鍵完成企業調研、人物 profiling 和公司治理篩選。

## 安裝方式

### 從 Plugin Marketplace 安裝

```
/plugin marketplace add stb-research
```

### 從本地路徑安裝

```
/plugin install /path/to/stb-research-plugin
```

## Slash Commands

### `/research [公司名稱]`

企業深度研究 — 自動跑完三大分析模組：

- **公司生平分析**：歷史脈絡、關鍵決策、競爭格局、治理財務
- **財報分析**：核心財務表現、業務板塊拆解、現金流與資本配置
- **併購分析**：歷年重大併購案、買方動機、賣方背景、市場反應

用法範例：
```
/research 台積電
/research 鴻海精密
```

### `/profile [人物姓名]`

人物網絡報告 — 自動生成 18 頁 LP Profile PPT：

- 個人介紹（學歷、職涯、經營哲學）
- 朋友圈 8 類人脈分析
- 家族樹（自動生成圖表）
- 家族企業版圖
- 迴避對象與引薦建議

用法範例：
```
/profile 郭台銘
/profile 林百里
```

### `/screening [公司名稱]`

公司 Screening Note — 快速產出治理、股權、營運、財務面篩選報告：

- 公司治理面：董事會組成、獨董比例、治理評鑑
- 股權結構面：大股東、董監持股、質押比例
- 營運面：業務結構、產業地位、重大事件
- 財務面：營收獲利趨勢、負債比率、現金流

用法範例：
```
/screening 聯發科
/screening 中華電信
```

## Skills

| Skill | 說明 |
|-------|------|
| `network-report-skill` | 人脈網絡報告 PPT 一鍵生成器 — 從 Deep Research 到 18 頁 PPT 全自動。使用 Gemini + Perplexity + web search 三源交叉驗證，透過 python-pptx 生成專業簡報。 |
| `research-methodology` | 企業研究方法論 Master Pipeline — 定義從資料搜集到報告產出的完整 SOP，確保研究品質一致。 |
| `screening-note` | Screening Note 模板與生成器 — 基於台灣董事學會通用模板，針對上市公司快速產出結構化篩選報告。 |

## 版本

- **v1.0.0** — 初始版本，包含三大 commands 和三個 skills
