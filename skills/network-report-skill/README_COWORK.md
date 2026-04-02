# Cowork 使用說明

## 快速開始

在 Cowork 中輸入以下指令即可：

```
請幫我製作【劉揚偉】（鴻海科技集團，股票代號：2317）的人脈網絡報告。

使用 network-report-skill：
1. 先用 web_search 搜集資料（參考 templates/research_prompt.md 的格式）
2. 整理成 references/data_schema.json 定義的 JSON 格式
3. 執行 scripts/generate_network_report.py 生成 PPT
```

## 步驟詳解

### Step 1: 資料搜集
Cowork/Claude 會根據 `templates/research_prompt.md` 的模板自動搜集資料。

### Step 2: JSON 整理
搜集完成後，將資料整理成符合 `references/data_schema.json` 格式的 JSON 文件。

### Step 3: 生成 PPT
```bash
pip install python-pptx --break-system-packages
python scripts/generate_network_report.py --input data.json --output /mnt/user-data/outputs/{人名}_人脈網絡報告.pptx
```

## 拿到 PPT 後你可能要做的事

1. **放人物照片**：Slide 3 個人介紹頁目前沒有照片，你可以手動插入
2. **放家族樹圖**：Slide 13 是 placeholder，你可以用 drawio 匯出 PNG 後插入
3. **微調資料**：AI 搜到的資料如果有誤，直接在 PPT 裡改
4. **調整排版**：如果某頁表格太長，可以手動調整行高

## 文件結構

```
network-report-skill/
├── SKILL.md                          # Skill 定義文件
├── README_COWORK.md                  # 本說明文件
├── scripts/
│   └── generate_network_report.py    # PPT 生成腳本（核心）
├── templates/
│   └── research_prompt.md            # Deep Research 搜集模板
└── references/
    └── data_schema.json              # JSON 資料格式定義（含範例）
```
