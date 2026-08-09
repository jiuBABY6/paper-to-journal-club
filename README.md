# Paper to Journal Club

本目录可独立作为 GitHub 仓库发布，并作为本地或团队 Marketplace 安装的 Codex 插件。它把论文先转成可追溯的证据包，再设计组会叙事，最后直接控制 Windows Microsoft PowerPoint 生成原生可编辑的 `.pptx`。运行时不使用 Node.js。

## 已实现能力

- 读取 TXT、Markdown、TeX 和 PDF；正式发布包通过自包含 `paper-parser.exe` 处理 PDF，并限制输入大小、页数、图片资产数和解析时长。
- 识别章节、候选结论、图号、逐页原文证据和可导出的论文图片资产，生成 `evidence_pack`。
- 默认强制生成研究背景、创新点、研究方法、实验数据、局限性和未来研究方向六个组会模块；每页保留来源 ID、证据状态和内容来源模式。
- 审核证据回链、必备模块覆盖率、标题、takeaway、要点密度和“汇报者讨论”标识；存在硬错误时不会启动 PowerPoint。
- 通过 Windows PowerPoint COM 创建 16:9、可编辑文本框和形状组成的 PPTX，并用 PowerPoint 原生渲染器导出预览图、审计文字溢出和画布越界。
- 提供 `powerpoint_status`、`inspect_powerpoint` 与 `audit_editable_pptx`：只有当前窗口检查使用 COM 时才会声明已连接当前 PowerPoint；生成始终新建后台演示文稿，不会修改用户当前打开的文件。
- PowerShell MCP 服务，不需要用户安装 Node.js、npm、Python 或 .NET。

## 本地演示

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\plugins\paper-to-journal-club\scripts\run-demo.ps1
```

该命令会基于 `examples/sample-paper.md` 生成 `examples/sample-deck-spec.json` 并打印审核结果。

要从 JSON 生成 PPTX，需要已安装 Windows Microsoft PowerPoint。插件只控制 PowerPoint，不适配或验证 WPS：

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\plugins\paper-to-journal-club\scripts\generate-editable-pptx.ps1 `
  -DeckSpecPath .\plugins\paper-to-journal-club\examples\sample-deck-spec.json `
  -OutputPath .\plugins\paper-to-journal-club\examples\my-journal-club.pptx
```

## Codex 中的建议工作流

1. 调用 `powerpoint_status`，确认 PowerPoint COM、当前窗口和可用能力；无当前文件时生成器会新建演示文稿。
2. 调用 `analyse_paper`，检查候选结论、逐页证据和图像资产。
3. 如需插入论文图，先审阅真实导出的图片资产，再通过 `figure_asset_selection` 显式传入 `figure id -> 图片路径`；插件不会仅按页码或图号猜测绑定。
4. 调用 `design_journal_club_deck`，让用户确认包含六模块的逐页故事线。
5. 调用 `audit_journal_club_deck`；所有硬错误必须先修复。
6. 调用 `generate_editable_pptx`，提供一个新的 `.pptx` 输出路径；如需替换已有文件，必须显式传 `overwrite=true`。
7. 使用返回的质量报告和预览图，或调用 `audit_editable_pptx` 重新打开文件做只读审计；不要仅凭脚本成功就认为版式正确。

## 功能边界

复杂原始科研图默认应以紧裁剪的“原子图片”保留，图注、箭头、标题和强调框则应在 PPT 中重建为独立对象。该原型不伪造图片含义，也不会在没有文本证据时虚构论文结论。

## 发布构建

发布维护者需要在有 .NET 8 SDK 的构建机运行 `plugins/paper-to-journal-club/scripts/build-paper-parser.ps1`，并将生成的 `assets/paper-parser.exe` 随插件发行。这个构建步骤只发生在发布阶段；普通用户不需要 .NET。GitHub Release、本地 Marketplace、验签和安装步骤见 [PUBLISHING.md](PUBLISHING.md)。
