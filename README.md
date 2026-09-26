# OfficeCLI Workflow · 批量改 Office 文檔

<div align="center">

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://microsoft.com/powershell)
[![Part of the MEP Automation Toolkit](https://img.shields.io/badge/Toolkit-MEP%20automation-1565C0?logo=github&logoColor=white)](https://github.com/David-CB666)

**OfficeCLI batch document modification workflow**

Modify existing .docx/.xlsx/.pptx files with batch JSON mode and stable path addressing.

[快速開始](#快速開始) · [文件結構](#文件結構)

</div>

![Office 批次改文件工作流](assets/workflow-overview.jpg)

---

> AI-native Office CLI 實戰工作流。用於**修改現有** .docx / .xlsx / .pptx 文檔——改日期、替換文字、更新表格、find/replace。

## 解決什麼問題

修改現有 Office 文檔時常見痛點：
- 用 python-docx/openpyxl 改，格式/字體/圖片全部走樣
- 手動打開改幾十個文件，重複機械勞動
- 多個修改操作並行執行，file-lock 競爭隨機失敗
- Git Bash 下路徑 `/` 被自動轉換，命令完全唔work

**OfficeCLI Workflow** 提供一套經過實戰驗證嘅操作流程，解決以上所有問題。

## 核心特性

### ⚡ 單體二進制，零依賴
`officecli.exe` 單一文件，唔需要安裝 Python / Node / 任何 runtime。

### 🚨 PowerShell 硬性要求
Git Bash 會將 `/` 自動轉換為 `C:/Program Files/Git/`，導致文檔根路徑 `/` 失效。
**所有 officecli 命令必須通過 PowerShell 調用。**

```powershell
# ✅ 正確
& "<OFFICECLI_DIR>/officecli.exe" add doc.pptx / --type slide --prop "title=Hello"

# ❌ 錯誤（Git Bash）
officecli add doc.pptx / --type slide    # / 被轉換
```

### 🔄 標準 4 步工作流

| 步驟 | 動作 | 目的 |
|:---|:---|:---|
| 1 | `view` / `get` 讀取結構 | 確認 paraId、表格坐標 |
| 2 | `close` 釋放 file-lock | view 會自動啟動 resident 鎖檔 |
| 3 | `batch --input changes.json` 批次修改 | 一次執行所有修改，避免競爭 |
| 4 | 驗證結果 | 讀取修改後內容確認正確 |

### 📦 Batch JSON 批次模式
唔好平行執行多個 `set` 命令（file-lock 競爭會隨機失敗）。所有修改打包成 JSON，一次過執行：

```json
[
  {"command":"set","path":"/body/p[@paraId=XXXXXXXX]","props":{"find":"舊文字","replace":"新文字"}},
  {"command":"set","path":"/body/tbl[1]/tr[3]/tc[4]","props":{"text":"7/21(二)"}},
  {"command":"set","path":"/body/tbl[2]/tr[8]/tc[6]","props":{"text":"9"}}
]
```

```powershell
officecli batch file.docx --input changes.json --force --json
```

### 🎯 穩定路徑尋址
- 段落用 `paraId`（唔會因為插入/刪除段落而移位）
- 表格用索引路徑 `/body/tbl[1]/tr[3]/tc[4]`
- 形狀/圖片用 shape ID

## 適用 vs 不適用

| ✅ 適用場景 | ❌ 唔適用場景 |
|:---|:---|
| 改幾個字 / 改日期 / 改數字 | 從零生成文檔 |
| 表格內少量單元格更新 | 複雜格式/排版邏輯 |
| find/replace 全文替換 | PDF 操作 |
| 批量修改多份同模板文件 | 新增大量內容 |

唔適用嘅場景，請用 Python 生態對應技能（python-docx / openpyxl / 材料報批 / 甘特圖等）。

## 快速開始

1. 確保使用 **PowerShell**（唔好用 Git Bash）
2. 讀 [DOCUMENTATION.md](DOCUMENTATION.md) 熟悉完整工作流
3. 先用 `view` / `get` 睇清楚文檔結構
4. 用 `close` 釋放鎖定
5. 寫 batch JSON，一次過執行所有修改
6. 再次 `view` 驗證結果

## 文件結構

```
officecli-workflow/
├── README.md       # 本文件（GitHub 預覽頁）
└── DOCUMENTATION.md        # 完整技能文檔（工作流 + 常見坑 + 實戰案例）
```

## 常見坑位提醒

1. **Git Bash 路徑問題** — 必須用 PowerShell
2. **忘記 close** — view/get 後檔案被 resident 鎖住，修改會失敗
3. **並行 set 命令** — 要用 batch 模式，唔好逐條 call
4. **find/replace 位置** — 要放喺 `props` 內，唔係頂層欄位
5. **paraId 變化** — 如果內容被完全替換，paraId 可能改變

---

## License

MIT License — feel free to use, modify, and share.

---

## Related repositories

Part of the **[MEP & construction document automation toolkit](https://github.com/David-CB666)** — open-source tools built from real jobsite workflows.

- **Handbook** — [ai-agent-manual](https://github.com/David-CB666/ai-agent-manual) (8-level AI cultivation for engineers)
- **Document generation** — [material-approval-pipeline](https://github.com/David-CB666/material-approval-pipeline) · [material-submittal-generator](https://github.com/David-CB666/material-submittal-generator) · [excel-template-filler](https://github.com/David-CB666/excel-template-filler) · [python-docx-photo-grid](https://github.com/David-CB666/python-docx-photo-grid) · [daily-construction-log](https://github.com/David-CB666/daily-construction-log)
- **Engineering calculation** — [lighting-lux-calculator](https://github.com/David-CB666/lighting-lux-calculator) · [ups-discharge-time-calculator](https://github.com/David-CB666/ups-discharge-time-calculator) · [gantt-chart-pro](https://github.com/David-CB666/gantt-chart-pro) · [electrical-test-report-generator](https://github.com/David-CB666/electrical-test-report-generator)
- **CAD & drawings** — [electrical-panel-label-plates](https://github.com/David-CB666/electrical-panel-label-plates)
- **Data & OCR** — [ocr-skill](https://github.com/David-CB666/ocr-skill) · [VBA-Macro-Reader-v2.0.0](https://github.com/David-CB666/VBA-Macro-Reader-v2.0.0)
- **Compliance & AI ops** — [confined-space-planner](https://github.com/David-CB666/confined-space-planner) · [skill-router](https://github.com/David-CB666/skill-router) · [consulting-services](https://github.com/David-CB666/consulting-services)
