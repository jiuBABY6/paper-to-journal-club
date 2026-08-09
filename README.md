# Paper to Journal Club

将一篇学术论文制作成可编辑、可导出、适合组会汇报的 Microsoft PowerPoint 演示稿。

Paper to Journal Club 是一个运行在 Windows 本机的 Codex 插件：它先从论文构建可追溯的证据包，再生成组会叙事和逐页提纲，经过内容审核后，由本机 Microsoft PowerPoint 创建原生可编辑的 PPTX。运行时不依赖 Node.js、npm、Python 或 .NET。

> 这是“GitHub Release + 本地 Codex Marketplace”分发的插件，不是公网 MCP 服务，也不会自动进入 OpenAI 公共 Plugins Directory。

## 适合谁

- 需要把论文快速转为 10–20 分钟组会汇报的科研人员；
- 希望保留可编辑文字、形状、箭头、表格和图注，而不是只得到一张图片的用户；
- 需要核对结论、图号、章节和原文证据的导师、学生或实验室成员。

## 你会得到什么

默认生成一份 16:9 的组会演示稿，并覆盖以下六个必备模块：

1. 研究背景与知识缺口；
2. 创新点与学术贡献；
3. 研究方法与实验设计；
4. 实验数据与核心发现；
5. 局限性与批判性评价；
6. 未来研究方向。

生成过程会输出：

- 可核对的 evidence pack：论文章节、候选结论、图号、页码和原文摘录；
- 可修改的逐页 deck spec：在生成 PowerPoint 前审阅故事线；
- 原生可编辑的 PPTX：文字、形状、箭头、流程图和表格都可继续修改；
- 每页 PNG 预览与质量报告：检查文字溢出、画布越界和可编辑对象结构。

可下载查看一个示例演示稿：[sample-journal-club.pptx](plugins/paper-to-journal-club/examples/sample-journal-club.pptx)。对应的可审阅内容规格见 [sample-deck-spec.json](plugins/paper-to-journal-club/examples/sample-deck-spec.json)。

## 核心功能

- 支持 PDF、TXT、Markdown 和 TeX 论文输入；
- 正式发行包自带 paper-parser.exe，普通用户无需安装解析环境；
- 默认生成六个组会模块；缺少来源或必备模块时会阻止生成 PPT；
- 每个关键结论可保留来源章节、图号、页码或原文摘录；
- 复杂原始科研图以紧裁剪图片插入；标题、箭头、注释和强调框尽量使用 PowerPoint 原生对象重建；
- 使用 Windows PowerPoint COM 创建、保存、导出和重新打开审计演示稿；
- 生成时始终新建后台演示文稿，不会覆盖你当前打开的 PowerPoint；
- 支持 powerpoint_status、inspect_powerpoint 和 audit_editable_pptx，如实说明是当前窗口检查还是文件只读审计。

## 使用前需要准备

- 64 位 Windows；
- 已安装并可正常启动的 Microsoft PowerPoint 桌面版；
- Codex 桌面版或 Codex CLI；
- 网络安装时能够访问 GitHub Release。

不支持 macOS、网页版 PowerPoint 或 WPS 演示。普通用户不需要 Git、Node.js、npm、Python 或 .NET SDK。

## 第一次使用：只需 3 步

### 第 1 步：安装插件

用户体验可以和 Scientific Illustrator 一样：把下面整段话直接发送给 Codex。不同之处是，本插件优先让 Codex 使用经过 SHA-256 校验的 GitHub Release 安装器，而不是直接检出开发仓库；安装器会把解压后的发行包根目录注册为**用户本机**的 Codex Marketplace。

~~~text
请安装 Paper to Journal Club：
https://github.com/jiuBABY6/paper-to-journal-club

请先读取该仓库的 GitHub Releases，选择最新的非草稿正式版本标签；如果没有正式 Release，请停止并告诉我，不要直接安装 main 或其他开发分支。
下载该 Release 中的 paper-to-journal-club-bootstrap.ps1 及其同名 .sha256 文件，核验 SHA-256 一致后，使用明确的 ReleaseTag 运行安装器。
安装器会下载并再次核验 Marketplace 发行包，将其根目录（含 .agents 和 plugins）注册为 Codex Marketplace，然后安装
paper-to-journal-club@paper-to-journal-club-tools。

如果需要下载、注册 Marketplace 或写入本机配置，请先向我说明并请求确认。
不要使用 ExecutionPolicy Bypass。完成后报告实际安装的版本，并提醒我完全重启 Codex。
~~~

首次安装时，Codex 可能要求你确认网络访问、PowerShell 执行或本机配置写入；这是正常的安全确认。若 Codex 无法代为执行本机安装，请按下方步骤手动运行同一套 Release 安装器。

#### 更可审计的方式：从 GitHub Release 安装

1. 打开 [Releases 页面](https://github.com/jiuBABY6/paper-to-journal-club/releases)，下载同一版本的：

   - paper-to-journal-club-bootstrap.ps1
   - paper-to-journal-club-bootstrap.ps1.sha256

2. 在 PowerShell 中核对下载脚本的 SHA-256：

~~~powershell
Get-FileHash -LiteralPath ./paper-to-journal-club-bootstrap.ps1 -Algorithm SHA256
Get-Content -LiteralPath ./paper-to-journal-club-bootstrap.ps1.sha256
~~~

3. 确认哈希一致后，在下载目录运行。下面以 v1.0.1 为例；安装新版本时替换标签即可：

~~~powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ./paper-to-journal-club-bootstrap.ps1 `
  -RepositoryUrl 'https://github.com/jiuBABY6/paper-to-journal-club' `
  -ReleaseTag 'v1.0.1'
~~~

安装器会下载并校验 Marketplace ZIP，检查 PowerPoint，然后自动执行本地 Marketplace 注册和插件安装。默认安装目录为：

~~~text
%LOCALAPPDATA%/Codex/marketplaces/paper-to-journal-club/current
~~~

安装完成后，完全退出并重新打开 Codex，再新建一个任务。

> 如果企业策略标记了已验证的下载脚本，可在确认 Release 来源和哈希后执行 Unblock-File -LiteralPath ./paper-to-journal-club-bootstrap.ps1。不要使用 ExecutionPolicy Bypass，也不要对来源不明的文件解除阻止。

#### 离线安装

管理员可下载同一 Release 中的 Marketplace ZIP 和对应 .sha256 文件，在离线环境核验后解压 ZIP。进入**解压后的包根目录**，运行：

~~~powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ./install.ps1
~~~

不要把 Codex Marketplace 指向 plugins/paper-to-journal-club 子目录；正确的根目录同时包含 .agents 和 plugins。

### 第 2 步：重启 Codex，并打开 Microsoft PowerPoint

安装完成后：

1. 完全退出并重新打开 Codex；
2. 新建一个任务；
3. 打开准备查看或编辑的 **Microsoft PowerPoint 桌面版**。

不需要预先新建演示文稿：插件会在后台创建一个新的 PPTX，不会覆盖你当前打开的演示文稿。仅支持 Windows Microsoft PowerPoint；不支持 WPS、macOS 或网页版 PowerPoint。

### 第 3 步：上传论文并复制提示词

1. 将论文 PDF 放在“桌面”“文档”或“下载”目录，然后在 Codex 中上传该文件；
2. 新建任务并粘贴下面的提示词；
3. 先审阅插件输出的逐页提纲和证据，再确认生成 PowerPoint；
4. 打开生成的 PPTX，结合预览图和质量报告进行最后核对。

> **文件访问边界。** 为防止恶意论文或提示词诱导插件读取任意本机文件，插件默认只读写“桌面”“文档”“下载”及自身临时目录。若你的论文或输出必须存放在资料盘，请在**完全退出 Codex 前**将可信资料根写入当前 Windows 用户环境变量，然后重启 Codex：
>
> ~~~powershell
> [Environment]::SetEnvironmentVariable(
>   'PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS',
>   'D:\Research;E:\SharedPapers',
>   'User'
> )
> ~~~
>
> 请只填确实存放论文/输出的窄目录；不要填写 `C:\`、整个用户目录或 `AppData`。Codex 的某些附件缓存路径可能不在默认目录内：遇到“approved user data directory”错误时，将论文另存到上述目录或配置资料根，而不是放宽到整个磁盘。

~~~text
[@paper-to-journal-club](plugin://paper-to-journal-club@paper-to-journal-club-tools)
使用 Paper to Journal Club，将我上传的论文制作成一份 15 分钟中文组会汇报。
必须包括：研究背景、创新点、研究方法、实验数据、局限性和未来研究方向。

先调用 powerpoint_status；先解析论文并输出可核对的逐页提纲。
每项关键结论都要保留图号、章节、页码或原文证据；不要把模型推断写成论文事实。
在我确认提纲后，再新建后台 16:9 可编辑 PowerPoint。
优先使用原生文本、形状、箭头、流程图和表格；仅把无法可靠重绘的原始科研图作为紧裁剪图片插入。
完成后保存 PPTX，导出预览图，并重新打开文件进行质量审计。
~~~

## 推荐工作流

1. powerpoint_status：确认本机 PowerPoint COM、当前窗口状态和可用能力；
2. analyse_paper：解析论文，得到章节、图号、证据和图片资产；
3. design_journal_club_deck：生成六模块逐页故事线；
4. 审阅提纲，并在需要时用 figure_asset_selection 显式选择已检查过的论文图片；
5. audit_journal_club_deck：修复所有硬错误；
6. generate_editable_pptx：保存到一个新的 PPTX 路径；
7. audit_editable_pptx：重新打开文件并核对预览图与结构质量。

## 编辑、导出与效果说明

生成后的 PPTX 可直接用 Microsoft PowerPoint 打开、编辑和另存为。常规内容会保留为原生对象：

- 文本框、标题、要点和脚注；
- 色块、圆角矩形、箭头、连接线和流程图；
- 表格、强调框和图注；
- 原始科研图作为独立图片对象插入。

插件优先保证“可编辑 + 可追溯”，而不是把整页扁平化为图片。复杂显微图、热图或论文原图本身仍可能以图片形式保留；这能避免虚构或错误重绘数据含义。

## 重要限制

- 论文候选结论默认可能标记为“需要复核”；汇报前必须核对原始文字、统计信息和因果表述；
- 缺失原文证据、缺失必备模块或来源不清的内容会阻止 PPT 生成；
- 插件只控制 Windows 桌面版 Microsoft PowerPoint，不支持 WPS、macOS 或网页版 PowerPoint；
- 请在上传论文、分享日志或发送演示稿前确认保密、版权和数据授权要求。

## 常见问题与升级

**安装后在 Codex 中看不到插件。** 完整退出 Codex 并重新打开，然后新建一个任务。若仍未出现，请重新运行 Release 安装命令，并确认 Codex CLI 可以运行。

**提示找不到 PowerPoint。** 需要安装并至少启动一次 Microsoft PowerPoint 桌面版；网页版和 WPS 不提供本插件需要的 Windows COM 自动化接口。

**企业策略阻止 PowerShell。** 安装器不会绕过企业安全策略。请让 IT 管理员审核 Release 哈希或部署经过 Authenticode 签名的版本。

**为什么提示 bundled PDF parser is unavailable。** 正式发行包应自带受验证的 `paper-parser.exe`。插件默认不会使用 Word 打开 PDF 作为自动回退，因为 Word COM 无法可靠终止卡住的恶意或损坏 PDF。维护者排障时才可在启动 Codex 前设置 `PAPER_TO_JOURNAL_CLUB_ALLOW_WORD_PDF_FALLBACK=1`；普通用户不应启用它。

**如何升级。** 从 Releases 页面选择新的固定版本标签，再次运行同一安装命令。安装器会保留此前的 current 版本目录，便于人工回退和排障。

**需要反馈问题。** 请在 [GitHub Issues](https://github.com/jiuBABY6/paper-to-journal-club/issues) 提供插件版本、Windows/PowerPoint 版本和不含敏感论文内容的最小复现步骤。

## 隐私与安全

论文解析和 PowerPoint 自动化在本机完成；插件不运营独立的论文或演示稿上传服务器，也不会主动收集遥测。上传到 Codex 的论文、提示词和结果仍受你的 Codex、模型提供方及组织策略约束。详见 [PRIVACY.md](PRIVACY.md) 和 [SECURITY.md](SECURITY.md)。

## 许可证

本项目采用 [MIT License](LICENSE)。第三方组件和许可证说明见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

## 发布者与开发者

发布、版本标签、SHA-256、代码签名、离线部署与 GitHub Actions 的维护说明见 [PUBLISHING.md](PUBLISHING.md)。发布者应在正式发行前更新插件清单中的作者/开发者名称和支持渠道。
