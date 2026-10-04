---
name: officecli-workflow
triggers: ["OfficeCLI", "批量修改文档", "改日期", "替换文字", "更新表格", "find/replace", "docx修改", "xlsx修改", "pptx修改", "Office修改"]
description: OfficeCLI 批量修改文档工作流。当需要修改现有 .docx/.xlsx/.pptx 文档（改日期、替换文字、更新表格、find/replace 等）时使用此技能。不适用于从零生成文档、PDF 操作或复杂格式逻辑（这些场景优先使用 Python 生态：GanttChart Pro、material-approval-pipeline 等）。
agent_created: true
---

# OfficeCLI 实战工作流

AI-native Office CLI（`officecli`），单体二进制，零依赖。用于修改现有 Office 文档。

## 环境

- 二进制：`<OFFICECLI_DIR>/officecli.exe`
- SKILL.md（内建）：`<SKILLS_DIR>/officecli/SKILL.md`
- MCP：`<MCP_CONFIG_PATH>`（stdio 模式）

---

## 🚨 关键规则：必须使用 PowerShell

**Git Bash 会将 `/` 自动转换为文件系统路径（转为 `C:/Program Files/Git/`），导致 officecli 的文档根路径 `/` 无效。任何涉及 `/` 路径的 officecli 命令必须通过 PowerShell 调用。**

```powershell
# ✅ PowerShell
& "<OFFICECLI_DIR>/officecli.exe" add doc.pptx / --type slide --prop "title=Hello"

# ❌ Git Bash
officecli add doc.pptx / --type slide    # / 被转换为 C:/Program Files/Git/
```

---

## 核心工作流

### 步骤 1：读取文档结构

```powershell
# 查看大纲
officecli view file.docx outline

# 查看全文
officecli view file.docx text --max-lines 200

# 取得表格 JSON
officecli get file.docx "/body/tbl[1]" --depth 2 --json
```

### 步骤 2：先关闭 resident，再做修改

由于之前的 `view`/`get` 命令会自动启动 resident 模式锁定档案，修改前必须先释放锁定：

```powershell
officecli close file.docx
```

### 步骤 3：使用 batch 模式批量修改

**不要平行执行多个 `set` 命令**（会竞争 file-lock，随机失败）。所有修改打包成一个 JSON，一次执行：

```json
[
  {"command":"set","path":"/body/p[@paraId=XXXXXXXX]","props":{"find":"旧文字","replace":"新文字"}},
  {"command":"set","path":"/body/tbl[1]/tr[3]/tc[4]","props":{"text":"7/21(二)"}},
  {"command":"set","path":"/body/tbl[2]/tr[8]/tc[6]","props":{"text":"9"}}
]
```

```powershell
officecli batch file.docx --input changes.json --force --json
```

**batch JSON 格式要点：**
- 段落文字替换：`{"command":"set","path":"...","props":{"find":"X","replace":"Y"}}`
- 表格储存格：`{"command":"set","path":"...","props":{"text":"新值"}}`
- `find`/`replace` 必须放在 `props` 内（非顶层栏位）
- 使用 `--force` 确保遇错继续执行

### 步骤 4：验证

```powershell
officecli view file.docx text --max-lines 200
officecli close file.docx
```

---

## 常用命令速查

| 任务 | 命令 |
|------|------|
| 查看大纲 | `view <file> outline` |
| 查看全文 | `view <file> text` |
| 取得 JSON | `get <file> <path> --depth N --json` |
| 段落替换文字 | `set <file> <path> --find X --replace Y` |
| 修改储存格 | `set <file> <cell-path> --prop text="值"` |
| 批量修改 | `batch <file> --input changes.json --force --json` |
| 关闭 resident | `close <file>` |

---

## 稳定路径定址

使用稳定 ID 而非位置索引（位置索引在插入/删除时会偏移）：

| 类型 | 格式 |
|------|------|
| Word 段落 | `/body/p[@paraId=XXXXXXXX]` |
| PPT 形状 | `/slide[N]/shape[@id=XXXXXXXX]` |
| 表格储存格 | `/body/tbl[N]/tr[N]/tc[N]` |

---

## 何时不要用 OfficeCLI

- 从零生成文档（用 GanttChart Pro / python-docx）
- 报批 + BQ 合并 PDF 操作（用 material-approval-pipeline）
- 复杂 Python 逻辑（用 S3/S4 生态）
- 甘特图生成（用 GanttChart Pro v15.0）
