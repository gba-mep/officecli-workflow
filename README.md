# OfficeCLI Workflow · 批量改 Office 文档

<div align="center">

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://microsoft.com/powershell)
[![Part of the MEP Automation Toolkit](https://img.shields.io/badge/Toolkit-MEP%20automation-1565C0?logo=github&logoColor=white)](https://github.com/gba-mep)

**OfficeCLI batch document modification workflow**

Modify existing .docx/.xlsx/.pptx files with batch JSON mode and stable path addressing.

[快速开始](#快速开始) · [文件结构](#文件结构)

</div>

---

> AI-native Office CLI 实战工作流。用于**修改现有** .docx / .xlsx / .pptx 文档——改日期、替换文字、更新表格、find/replace。

## 解决什么问题

修改现有 Office 文档时常见痛点：
- 用 python-docx/openpyxl 改，格式/字体/图片全部走样
- 手动打开改几十个文件，重复机械劳动
- 多个修改操作并行执行，file-lock 竞争随机失败
- Git Bash 下路径 `/` 被自动转换，命令完全不work

**OfficeCLI Workflow** 提供一套经过实战验证的操作流程，解决以上所有问题。

## 核心特性

### ⚡ 单体二进制，零依赖
`officecli.exe` 单一文件，不需要安装 Python / Node / 任何 runtime。

### 🚨 PowerShell 硬性要求
Git Bash 会将 `/` 自动转换为 `C:/Program Files/Git/`，导致文档根路径 `/` 失效。
**所有 officecli 命令必须通过 PowerShell 调用。**

```powershell
# ✅ 正确
& "<OFFICECLI_DIR>/officecli.exe" add doc.pptx / --type slide --prop "title=Hello"

# ❌ 错误（Git Bash）
officecli add doc.pptx / --type slide    # / 被转换
```

### 🔄 标准 4 步工作流

| 步骤 | 动作 | 目的 |
|:---|:---|:---|
| 1 | `view` / `get` 读取结构 | 确认 paraId、表格坐标 |
| 2 | `close` 释放 file-lock | view 会自动启动 resident 锁档 |
| 3 | `batch --input changes.json` 批次修改 | 一次执行所有修改，避免竞争 |
| 4 | 验证结果 | 读取修改后内容确认正确 |

### 📦 Batch JSON 批次模式
不要平行执行多个 `set` 命令（file-lock 竞争会随机失败）。所有修改打包成 JSON，一次过执行：

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

### 🎯 稳定路径寻址
- 段落用 `paraId`（不会因为插入/删除段落而移位）
- 表格用索引路径 `/body/tbl[1]/tr[3]/tc[4]`
- 形状/图片用 shape ID

## 适用 vs 不适用

| ✅ 适用场景 | ❌ 不适用场景 |
|:---|:---|
| 改几个字 / 改日期 / 改数字 | 从零生成文档 |
| 表格内少量单元格更新 | 复杂格式/排版逻辑 |
| find/replace 全文替换 | PDF 操作 |
| 批量修改多份同模板文件 | 新增大量内容 |

不适用的场景，请用 Python 生态对应技能（python-docx / openpyxl / 材料报批 / 甘特图等）。

## 快速开始

1. 确保使用 **PowerShell**（不要用 Git Bash）
2. 读 [DOCUMENTATION.md](DOCUMENTATION.md) 熟悉完整工作流
3. 先用 `view` / `get` 看清楚文档结构
4. 用 `close` 释放锁定
5. 写 batch JSON，一次过执行所有修改
6. 再次 `view` 验证结果

## 文件结构

```
officecli-workflow/
├── README.md       # 本文件（GitHub 预览页）
└── DOCUMENTATION.md        # 完整技能文档（工作流 + 常见坑 + 实战案例）
```

## 常见坑位提醒

1. **Git Bash 路径问题** — 必须用 PowerShell
2. **忘记 close** — view/get 后档案被 resident 锁住，修改会失败
3. **并行 set 命令** — 要用 batch 模式，不要逐条 call
4. **find/replace 位置** — 要放在 `props` 内，不是顶层栏位
5. **paraId 变化** — 如果内容被完全替换，paraId 可能改变

---

## License

MIT License — feel free to use, modify, and share.

---

## Related repositories

Part of the **[MEP & construction document automation toolkit](https://github.com/gba-mep)** — open-source tools built from real jobsite workflows.

- **Handbook** — [ai-agent-manual](https://github.com/gba-mep/ai-agent-manual) (8-level AI cultivation for engineers)
- **Document generation** — [material-approval-pipeline](https://github.com/gba-mep/material-approval-pipeline) · [material-submittal-generator](https://github.com/gba-mep/material-submittal-generator) · [excel-template-filler](https://github.com/gba-mep/excel-template-filler) · [python-docx-photo-grid](https://github.com/gba-mep/python-docx-photo-grid) · [daily-construction-log](https://github.com/gba-mep/daily-construction-log)
- **Engineering calculation** — [lighting-lux-calculator](https://github.com/gba-mep/lighting-lux-calculator) · [ups-discharge-time-calculator](https://github.com/gba-mep/ups-discharge-time-calculator) · [gantt-chart-pro](https://github.com/gba-mep/gantt-chart-pro) · [electrical-test-report-generator](https://github.com/gba-mep/electrical-test-report-generator)
- **CAD & drawings** — [electrical-panel-label-plates](https://github.com/gba-mep/electrical-panel-label-plates)
- **Data & OCR** — [ocr-skill](https://github.com/gba-mep/ocr-skill) · [VBA-Macro-Reader-v2.0.0](https://github.com/gba-mep/VBA-Macro-Reader-v2.0.0)
- **Compliance & AI ops** — [confined-space-planner](https://github.com/gba-mep/confined-space-planner) · [路由规则](https://github.com/gba-mep/路由规则) · [consulting-services](https://github.com/gba-mep/consulting-services)
