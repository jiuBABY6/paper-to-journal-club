# Paper to Journal Club 安装与使用

这是一个面向 **Windows + Microsoft PowerPoint 桌面版** 的本地 Codex 插件。它会将论文整理为可追溯的证据包和组会叙事，并生成以原生文本、形状、表格和图片组成的可编辑 PowerPoint。

## 运行条件

- 64 位 Windows；
- 已安装并可正常启动的 Microsoft PowerPoint 桌面版；
- 已安装 Codex 桌面版或 Codex CLI；
- 完整、已验证的本发行包。

不需要安装 Node.js、npm、Python、Microsoft Word 或 .NET SDK。PDF 文本提取由发行包内预编译的 `paper-parser.exe` 完成。

## 离线安装

1. 将 `paper-to-journal-club-marketplace-<version>.zip` 解压到一个稳定目录，例如 `C:\CodexMarketplaces\PaperToJournalClub`。安装后不要删除或移动该目录，因为 Codex 会将其作为本地 Marketplace 源。
2. 在解压后的**包根目录**执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\install.ps1
```

3. 安装器会校验 `SHA256SUMS.txt`、PDF 解析器、Windows 位数和 PowerPoint COM 注册状态，然后登记并安装：

```text
paper-to-journal-club@paper-to-journal-club-tools
```

4. 完全退出并重新打开 Codex，再新建一个任务。

`RemoteSigned` 不会绕过企业组策略。若组织禁止未签名脚本，请让 IT 管理员审核并部署经过代码签名的发行包；安装器不会修改系统执行策略。

## 首次使用

上传论文 PDF 后，可在 Codex 中使用如下提示词：

```text
[@paper-to-journal-club](plugin://paper-to-journal-club@paper-to-journal-club-tools)
使用 Paper to Journal Club，将我上传的论文制作成一份 15 分钟中文组会汇报。
必须包括：研究背景、创新点、研究方法、实验数据、局限性和未来研究方向。
先调用 powerpoint_status；先解析论文并输出可核对的逐页提纲，保留关键结论的图号、章节或原文证据。
我确认提纲后，再在 Microsoft PowerPoint 中新建后台 16:9 可编辑 PPTX；优先使用原生文字、形状、箭头、流程图和表格。
仅把无法可靠重绘的原始科研图作为紧裁剪图片插入。完成后保存 PPTX，并导出预览图检查版式。
```

请始终先核对提纲，尤其是图号、统计信息、局限性和未来方向。插件不会把模型推断伪装为论文事实。

## 常见问题

**找不到 PowerPoint。** 请安装 Microsoft PowerPoint 桌面版并至少启动一次；网页版不提供本插件所需的 COM 自动化接口。

**找不到 Codex CLI。** 更新或重新安装 Codex 桌面版，并在新的 PowerShell 窗口中确认 `codex --version` 可运行。

**提示解析器缺失或无效。** 说明下载或解压不完整。请重新获取带 `SHA256SUMS.txt` 的正式 Release ZIP，不要尝试用本机 Node.js、Python 或 .NET SDK 替代。

**PowerShell 提示文件来自互联网。** 优先使用发布者提供的 Authenticode 签名版本。若你已核对正式 ZIP 的 SHA-256 并确认可信，可在解压后的发行包根目录执行：

```powershell
Get-ChildItem -Recurse -File | Unblock-File
```

不要对来源不明的目录执行该命令。

## 隐私

PDF 解析和 PowerPoint 自动化均在本机完成，不会额外上传论文或演示稿到独立服务器。上传给 Codex 的文件和对话内容仍受你的 Codex、模型提供方和组织策略约束。详见插件目录中的 `PRIVACY.md`。
