---
name: officecli-workflow
triggers: ["OfficeCLI", "批量修改文檔", "改日期", "替換文字", "更新表格", "find/replace", "docx修改", "xlsx修改", "pptx修改", "Office修改"]
description: OfficeCLI 批量修改文檔工作流。當需要修改現有 .docx/.xlsx/.pptx 文檔（改日期、替換文字、更新表格、find/replace 等）時使用此技能。不適用於從零生成文檔、PDF 操作或複雜格式邏輯（這些場景優先使用 Python 生態：GanttChart Pro、macau-material-approval 等）。
agent_created: true
---

# OfficeCLI 實戰工作流

AI-native Office CLI（`officecli`），單體二進制，零依賴。用於修改現有 Office 文檔。

## 環境

- 二進制：`<OFFICECLI_DIR>/officecli.exe`
- SKILL.md（內建）：`<SKILLS_DIR>/officecli/SKILL.md`
- MCP：`<MCP_CONFIG_PATH>`（stdio 模式）

---

## 🚨 關鍵規則：必須使用 PowerShell

**Git Bash 會將 `/` 自動轉換為文件系統路徑（轉為 `C:/Program Files/Git/`），導致 officecli 的文檔根路徑 `/` 無效。任何涉及 `/` 路徑的 officecli 命令必須通過 PowerShell 調用。**

```powershell
# ✅ PowerShell
& "<OFFICECLI_DIR>/officecli.exe" add doc.pptx / --type slide --prop "title=Hello"

# ❌ Git Bash
officecli add doc.pptx / --type slide    # / 被轉換為 C:/Program Files/Git/
```

---

## 核心工作流

### 步驟 1：讀取文檔結構

```powershell
# 查看大綱
officecli view file.docx outline

# 查看全文
officecli view file.docx text --max-lines 200

# 取得表格 JSON
officecli get file.docx "/body/tbl[1]" --depth 2 --json
```

### 步驟 2：先關閉 resident，再做修改

由於之前的 `view`/`get` 命令會自動啟動 resident 模式鎖定檔案，修改前必須先釋放鎖定：

```powershell
officecli close file.docx
```

### 步驟 3：使用 batch 模式批量修改

**不要平行執行多個 `set` 命令**（會競爭 file-lock，隨機失敗）。所有修改打包成一個 JSON，一次執行：

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

**batch JSON 格式要點：**
- 段落文字替換：`{"command":"set","path":"...","props":{"find":"X","replace":"Y"}}`
- 表格儲存格：`{"command":"set","path":"...","props":{"text":"新值"}}`
- `find`/`replace` 必須放在 `props` 內（非頂層欄位）
- 使用 `--force` 確保遇錯繼續執行

### 步驟 4：驗證

```powershell
officecli view file.docx text --max-lines 200
officecli close file.docx
```

---

## 常用命令速查

| 任務 | 命令 |
|------|------|
| 查看大綱 | `view <file> outline` |
| 查看全文 | `view <file> text` |
| 取得 JSON | `get <file> <path> --depth N --json` |
| 段落替換文字 | `set <file> <path> --find X --replace Y` |
| 修改儲存格 | `set <file> <cell-path> --prop text="值"` |
| 批量修改 | `batch <file> --input changes.json --force --json` |
| 關閉 resident | `close <file>` |

---

## 穩定路徑定址

使用穩定 ID 而非位置索引（位置索引在插入/刪除時會偏移）：

| 類型 | 格式 |
|------|------|
| Word 段落 | `/body/p[@paraId=XXXXXXXX]` |
| PPT 形狀 | `/slide[N]/shape[@id=XXXXXXXX]` |
| 表格儲存格 | `/body/tbl[N]/tr[N]/tc[N]` |

---

## 何時不要用 OfficeCLI

- 從零生成文檔（用 GanttChart Pro / python-docx）
- 報批 + BQ 合併 PDF 操作（用 macau-material-approval）
- 複雜 Python 邏輯（用 S3/S4 生態）
- 甘特圖生成（用 GanttChart Pro v15.0）
