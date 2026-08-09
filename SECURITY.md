# 安全说明

## 支持范围

本插件只支持 Windows 64 位和 Microsoft PowerPoint 桌面版。它通过本机 PowerShell 启动 MCP 服务，并通过 Office COM 创建本地 PPTX；不开放网络监听端口。

## 安全使用建议

- 仅安装来自受信任发布渠道、且能通过发布者在独立渠道公布的 SHA-256 校验值核验的发行包；包内 `SHA256SUMS.txt` 只能发现意外损坏，不能单独构成发布者身份保证；
- 引导安装器只接受指定 GitHub 仓库的 HTTPS 正式 Release，拒绝草稿、预发布和非 Release 下载路径；它会先校验 ZIP 的外部 SHA-256，再执行受限解压；
- 解压后的 `SHA256SUMS.txt` 做双向集合校验：每个清单条目都必须存在且哈希一致，且除清单自身外不存在未列出的文件；发行包只允许一个固定的本地 Marketplace 条目及一个固定参数的 PowerShell MCP server；
- 安装器限制 ZIP 条目数、单项解压大小、总实际解压量和压缩比，并拒绝路径穿越、绝对路径、重解析点与大小写冲突路径；
- GitHub Release 工作流将第三方 Actions 固定到完整提交 SHA、禁止 checkout 持久化令牌，并拒绝覆写已经存在的 Release；
- 对公开或企业分发，建议使用可信代码签名证书签署 EXE 和 PowerShell 脚本，并由发布者在独立发布页公布签名和哈希；
- 不要在未经审阅的目录中执行 `install.ps1`；
- 论文、PPTX、预览和临时资产路径默认只能位于“桌面”“文档”“下载”或插件专用临时目录；需要资料盘时，只能在启动 Codex 前通过 `PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS` 声明必要的窄目录，切勿将 `C:\`、整个用户目录或 `AppData` 作为允许根；
- `paper-parser.exe` 缺失时，Word PDF COM 回退默认关闭；它只有在启动 Codex 前显式设置 `PAPER_TO_JOURNAL_CLUB_ALLOW_WORD_PDF_FALLBACK=1` 时才会启用，因为该回退无法可靠终止卡住的恶意或损坏 PDF；
- 不要将含未公开研究数据、个人信息或凭据的论文和日志公开分享；
- 若企业策略限制 PowerShell 或 Office COM，请由 IT 管理员批准后使用。

代码签名目前为维护者可选的发布预检，不会强制普通用户安装证书或降低 Windows 执行策略。未签名发行包应由维护者明确披露，并由用户在可信渠道复核 Release、签名（如有）和外部 SHA-256。

## 漏洞报告

请通过插件发布页或维护仓库的私密联系渠道向发布者报告安全问题。报告中不要附带真实的敏感论文、令牌或个人信息；请提供可复现的最小化示例、插件版本和 Windows/PowerPoint 版本。
