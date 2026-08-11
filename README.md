# 🧪 Paper to Journal Club

将一篇学术论文制作成可编辑、可导出、适合组会汇报的 Microsoft PowerPoint 演示稿。

> ✨ 从论文证据出发，先生成可审阅的故事线，再用本机 PowerPoint 创建原生可编辑 PPTX。

**Windows · Microsoft PowerPoint · Codex 本地 Marketplace**

[🔖 最新正式版本](https://github.com/jiuBABY6/paper-to-journal-club/releases/latest) · [🐛 反馈问题](https://github.com/jiuBABY6/paper-to-journal-club/issues) · [📜 更新记录](CHANGELOG.md)

Paper to Journal Club 运行在 Windows 本机：它先从论文构建可追溯的证据包，再生成组会叙事和逐页提纲；经你确认后，由本机 Microsoft PowerPoint 创建原生可编辑的 PPTX。普通用户运行时不需要安装 Node.js、npm、Python 或 .NET SDK。

> 🔒 本项目采用“GitHub Release + 本地 Codex Marketplace”分发。它不是公网 MCP 服务，也不会自动进入 OpenAI 公共插件目录。

## 🎯 适合谁

- 想把论文快速转为 10–20 分钟组会汇报的科研人员；
- 希望保留可编辑文字、形状、箭头、表格和图注，而不是只得到一张扁平图片的用户；
- 需要核对结论、图号、章节和原文证据的导师、学生与实验室成员。

## 📦 你会得到什么

默认生成一份 16:9 组会演示稿，覆盖六个必备模块：

1. 🌱 研究背景与知识缺口；
2. 💡 创新点与学术贡献；
3. 🔬 研究方法与实验设计；
4. 📊 实验数据与核心发现；
5. ⚖️ 局限性与批判性评价；
6. 🔭 未来研究方向。

生成过程中还会输出：

- 🧾 **Evidence pack**：论文章节、候选结论、图号、页码和原文摘录；
- 🗂️ **Deck spec**：生成 PowerPoint 前可修改、可审阅的逐页提纲；
- 🖼️ **可编辑 PPTX**：文字、形状、箭头、流程图和表格都可继续修改；
- ✅ **预览与质量报告**：逐页 PNG 预览、文字溢出、画布越界和对象结构检查。

可先查看示例：[sample-journal-club.pptx](plugins/paper-to-journal-club/examples/sample-journal-club.pptx)；对应的可审阅内容规格见 [sample-deck-spec.json](plugins/paper-to-journal-club/examples/sample-deck-spec.json)。

## ✨ 核心能力

- 支持 PDF、TXT、Markdown 和 TeX 论文输入；
- 正式发行包自带 `paper-parser.exe`，普通用户无需配置解析环境；
- 缺少来源或六个必备模块时，会阻止生成 PPT，避免“看起来完整、实际不可追溯”；
- 每个关键结论可保留来源章节、图号、页码或原文摘录；
- 复杂原始科研图可作为紧裁剪图片插入；标题、箭头、注释和强调框尽量用 PowerPoint 原生对象构建；
- 使用 Windows PowerPoint COM 创建、保存、导出和重新打开审计演示稿；
- 始终创建新的后台演示文稿，不会覆盖你当前打开的 PowerPoint；
- 支持 `powerpoint_status`、`inspect_powerpoint` 和 `audit_editable_pptx`，如实说明当前窗口检查或文件只读审计的结果。

## ✅ 使用前准备

| 需要 | 说明 |
|---|---|
| 🪟 64 位 Windows | 当前仅支持 Windows。 |
| 📽️ Microsoft PowerPoint 桌面版 | 已安装且至少成功启动过一次。 |
| 🤖 Codex 桌面版 | 用于浏览个人 Marketplace、安装并调用插件。 |
| 🌐 GitHub Release 网络访问 | 仅在在线安装或升级时需要。 |

不支持 macOS、网页版 PowerPoint 或 WPS 演示。普通用户不需要 Git、Node.js、npm、Python 或 .NET SDK。

## 🚀 安装插件：选择一种方式

| 方式 | 适合谁 | 特点 |
|---|---|---|
| 🤖 让 Codex 安装 | 大多数用户 | 最省事；Codex 会在操作前请求确认。 |
| 🔐 从 Release 手动安装 | 希望逐项核验的用户 | 先比对脚本 SHA-256，再运行固定版本安装器。 |
| ⌨️ 命令行下载并安装 | 熟悉 PowerShell 的用户 | 不需要打开浏览器，仍会校验 SHA-256。 |
| 📴 离线安装 | 内网或受限网络环境 | 可在联网机器下载并核验后转移 ZIP。 |

> 🧭 **安装器会自动选择最省事的安全路径。** 在你确认安装后，它会验证受信的 Codex CLI，并尝试自动完成安装；CLI 自身可能写入少量临时数据。只有 CLI 无法执行（例如 WindowsApps 显示“拒绝访问”）或自动安装失败且已确认回滚时，安装器才会安全回退到个人 Marketplace。若 CLI 已尝试登记但最终状态无法可靠确认，安装器会停止并报告错误，而不会擅自删除或覆盖你的配置。回退后完全重启 Codex，在 **Plugins Directory** 中点击一次 **Install** 即可；普通用户不需要在 PowerShell 中安装或手动调用 Codex CLI。

### 🤖 方式 1：把安装说明发送给 Codex（推荐）

新建一个普通 Codex 任务，把下面整段话完整发送给 Codex：

```text
请安装 Paper to Journal Club：
https://github.com/jiuBABY6/paper-to-journal-club

请先读取该仓库的 GitHub Releases，选择最新的非草稿正式版本标签；如果没有正式 Release，请停止并告诉我，不要直接安装 main 或其他开发分支。
下载该 Release 中的 paper-to-journal-club-bootstrap.ps1 及其同名 .sha256 文件，核验 SHA-256 一致后，使用明确的 ReleaseTag 运行安装器。
在我确认安装后，安装器必须验证受信 Codex CLI 的实体路径和签名，并只使用已验证 CLI 的绝对路径在同一安装事务中查询 Marketplace、尝试自动完成 Paper to Journal Club 的安装；CLI 自身可能写入临时数据。若 CLI 不可执行、Marketplace 查询或自动安装失败且可确认回滚，请不要修改 WindowsApps 权限、不要反复调用失败的 CLI，而是自动部署到当前用户的个人 Marketplace；若 CLI 已尝试登记但最终状态无法可靠确认，请停止并明确报告错误，不要自动删除或回退未知配置。

如果需要下载、自动安装、部署个人 Marketplace 或写入本机配置，请先向我说明并请求确认。
不要使用 ExecutionPolicy Bypass。完成后报告实际采用的安装路径和版本：自动安装成功时提醒我完全重启 Codex；个人 Marketplace 回退时提醒我重启后在 Plugins Directory 中点击 Install。
```

首次安装时，Codex 可能要求你确认网络访问、PowerShell 执行或本机配置写入；这是正常的安全确认。若 Codex 无法代为执行本机安装，请使用下面任一手动方式。

### 🔐 方式 2：浏览器下载 Release 安装器

1. 打开 [Releases 页面](https://github.com/jiuBABY6/paper-to-journal-club/releases)，选择一个固定的正式版本，例如 `v1.0.2`。
2. 下载同一 Release 中的：

   - `paper-to-journal-club-bootstrap.ps1`
   - `paper-to-journal-club-bootstrap.ps1.sha256`

3. 在下载目录打开 PowerShell，核对脚本哈希：

```powershell
Get-FileHash -LiteralPath ./paper-to-journal-club-bootstrap.ps1 -Algorithm SHA256
Get-Content -LiteralPath ./paper-to-journal-club-bootstrap.ps1.sha256
```

4. 确认两处 SHA-256 一致后运行。以下以 `v1.0.2` 为例；安装新版时只替换标签：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ./paper-to-journal-club-bootstrap.ps1 `
  -RepositoryUrl 'https://github.com/jiuBABY6/paper-to-journal-club' `
  -ReleaseTag 'v1.0.2'
```

安装器会再次校验 Marketplace ZIP、检查 PowerPoint，然后验证受信 Codex CLI 并尝试自动安装。发行包缓存目录为：

```text
%LOCALAPPDATA%\Codex\marketplaces\paper-to-journal-club\current
```

运行结果会明确告诉你下一步：

- **自动安装成功**：完全退出并重新打开 Codex，再新建一个任务；不需要手动打开插件目录。
- **个人 Marketplace 回退**：完全退出并重新打开 Codex，打开 **Plugins Directory**，选择个人 Marketplace，点击 **Paper to Journal Club → Install**，再新建一个任务。

> 🛡️ 如果企业策略标记了已核验的下载脚本，请在确认 Release 来源与 SHA-256 后执行 `Unblock-File -LiteralPath ./paper-to-journal-club-bootstrap.ps1`。不要使用 `ExecutionPolicy Bypass`，也不要对来源不明的文件解除阻止。

### ⌨️ 方式 3：用 PowerShell 下载、校验并安装固定版本

下方命令只下载 `v1.0.2` 的正式 Release。将第一行的标签替换为你要安装的正式版本；不要改为 `main`。

```powershell
$tag = 'v1.0.2'
$downloadDirectory = Join-Path $env:USERPROFILE 'Downloads\Paper-to-Journal-Club'
$releaseBase = "https://github.com/jiuBABY6/paper-to-journal-club/releases/download/$tag"
$scriptPath = Join-Path $downloadDirectory 'paper-to-journal-club-bootstrap.ps1'
$checksumPath = "$scriptPath.sha256"

New-Item -ItemType Directory -Path $downloadDirectory -Force | Out-Null
Invoke-WebRequest -Uri "$releaseBase/paper-to-journal-club-bootstrap.ps1" -OutFile $scriptPath
Invoke-WebRequest -Uri "$releaseBase/paper-to-journal-club-bootstrap.ps1.sha256" -OutFile $checksumPath

$expected = ((Get-Content -LiteralPath $checksumPath -Raw).Trim() -split '\s+')[0].ToLowerInvariant()
$actual = (Get-FileHash -LiteralPath $scriptPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actual -ne $expected) { throw '安装器 SHA-256 不匹配，已停止安装。' }

powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $scriptPath `
  -RepositoryUrl 'https://github.com/jiuBABY6/paper-to-journal-club' `
  -ReleaseTag $tag
```

### 📴 方式 4：离线 / 受限网络安装

1. 在可访问 GitHub 的设备上，从同一 Release 下载 Marketplace ZIP 及其 `.sha256` 文件。
2. 校验 ZIP 的 SHA-256 后，将两者转移到离线 Windows 设备。
3. 解压 ZIP，进入**解压后的包根目录**（该目录同时包含 `.agents` 与 `plugins`）。
4. 在该目录运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ./install.ps1
```

离线安装器也会在你确认后验证受信 CLI 并尝试自动安装；CLI 不可执行或自动安装失败且可确认回滚时，会回退到个人 Marketplace。若 CLI 已尝试登记但最终状态无法可靠确认，安装器会停止而不是擅自处理未知配置。解压后的包根目录只用于运行包内 `install.ps1` 和完整性验证；不要手动修改 Codex 配置或 WindowsApps 权限。

## 🧭 安装后：第一次使用只需 3 步

### 1. 🔄 完全重启 Codex

安装结束后，完全退出并重新打开 Codex：

- 若安装结果显示“自动安装成功”，直接新建一个任务；
- 若安装结果显示“个人 Marketplace 回退”，打开 **Plugins Directory**，选择个人 Marketplace，找到 **Paper to Journal Club** 并点击 **Install**，再新建一个任务。

### 2. 📽️ 打开 Microsoft PowerPoint 桌面版

打开准备查看或编辑的 PowerPoint。你不需要预先新建演示文稿：插件会在后台创建新的 PPTX，不会覆盖当前打开的演示文稿。

### 3. 📄 上传论文并发送提示词

1. 将论文 PDF 放在“桌面”“文档”或“下载”目录，再在 Codex 中上传该文件；
2. 新建任务并粘贴下面的提示词；
3. 先审阅逐页提纲和证据，再确认生成 PowerPoint；
4. 打开生成的 PPTX，结合预览图和质量报告做最后核对。

> 🔒 **文件访问边界：** 为防止恶意论文或提示词诱导插件读取任意本机文件，插件默认只读写“桌面”“文档”“下载”及自身临时目录。若论文或输出必须存放在资料盘，请在**完全退出 Codex 前**将可信资料根写入当前 Windows 用户环境变量，然后重启 Codex：
>
> ```powershell
> [Environment]::SetEnvironmentVariable(
>   'PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS',
>   'D:\Research;E:\SharedPapers',
>   'User'
> )
> ```
>
> 只填写确实存放论文或输出的窄目录；不要填写 `C:\`、整个用户目录或 `AppData`。如果 Codex 的附件缓存不在默认目录而提示“approved user data directory”，请将论文另存到上述目录或配置资料根，而不是放宽到整个磁盘。

```text
[@paper-to-journal-club](plugin://paper-to-journal-club@paper-to-journal-club-tools)
使用 Paper to Journal Club，将我上传的论文制作成一份 15 分钟中文组会汇报。
必须包括：研究背景、创新点、研究方法、实验数据、局限性和未来研究方向。

先调用 powerpoint_status；先解析论文并输出可核对的逐页提纲。
每项关键结论都要保留图号、章节、页码或原文证据；不要把模型推断写成论文事实。
在我确认提纲后，再新建后台 16:9 可编辑 PowerPoint。
优先使用原生文本、形状、箭头、流程图和表格；仅把无法可靠重绘的原始科研图作为紧裁剪图片插入。
完成后保存 PPTX，导出预览图，并重新打开文件进行质量审计。
```

如果第一行插件命令没有被识别，请在 Codex 输入框的插件菜单中选择 **Paper to Journal Club**，再发送后续提示词。

## 🎬 推荐工作流

1. `powerpoint_status`：确认本机 PowerPoint COM、当前窗口状态和可用能力；
2. `analyse_paper`：解析论文，得到章节、图号、证据和图片资产；
3. `design_journal_club_deck`：生成六模块逐页故事线；
4. 审阅提纲，并在需要时用 `figure_asset_selection` 显式选择已检查过的论文图片；
5. `audit_journal_club_deck`：修复所有硬错误；
6. `generate_editable_pptx`：保存到一个新的 PPTX 路径；
7. `audit_editable_pptx`：重新打开文件并核对预览图与结构质量。

## ✏️ 编辑、导出与效果说明

生成后的 PPTX 可直接用 Microsoft PowerPoint 打开、编辑和另存为。常规内容会保留为原生对象：

- 文本框、标题、要点和脚注；
- 色块、圆角矩形、箭头、连接线和流程图；
- 表格、强调框和图注；
- 原始科研图作为独立图片对象插入。

插件优先保证“可编辑 + 可追溯”，而不是把整页扁平化为图片。复杂显微图、热图或论文原图本身仍可能以图片形式保留；这样可以避免虚构或错误重绘数据含义。

## ⚠️ 重要限制

- 论文候选结论可能标记为“需要复核”；汇报前必须核对原始文字、统计信息和因果表述；
- 缺失原文证据、缺失必备模块或来源不清的内容会阻止 PPT 生成；
- 插件只控制 Windows 桌面版 Microsoft PowerPoint，不支持 WPS、macOS 或网页版 PowerPoint；
- 请在上传论文、分享日志或发送演示稿前确认保密、版权和数据授权要求。

## 🧩 常见问题与升级

**安装后在 Codex 中看不到插件？** 先查看安装器最后的结果：若显示“自动安装成功”，完全重启 Codex 后直接新建任务；若显示“个人 Marketplace 回退”，在 **Plugins Directory** 的个人 Marketplace 中点击 Install。若仍未出现，请重新运行同一版本的 Release 安装命令。

**结果显示 CLI 不可执行或 WindowsApps 拒绝访问？** 这是安装器触发个人 Marketplace 回退的条件，不表示发行包校验失败。不要修改 WindowsApps 权限、不要用管理员权限强行执行；完全重启 Codex 后，在 **Plugins Directory** 的个人 Marketplace 中点击 Install 即可。

**提示找不到 PowerPoint？** 需要安装并至少启动一次 Microsoft PowerPoint 桌面版；网页版和 WPS 不提供本插件需要的 Windows COM 自动化接口。

**企业策略阻止 PowerShell？** 安装器不会绕过企业安全策略。请让 IT 管理员审核 Release 哈希，或部署经过 Authenticode 签名的版本。

**为什么提示 `bundled PDF parser is unavailable`？** 正式发行包应自带受验证的 `paper-parser.exe`。插件默认不会使用 Word 打开 PDF 作为自动回退，因为 Word COM 无法可靠终止卡住的恶意或损坏 PDF。维护者排障时才可在启动 Codex 前设置 `PAPER_TO_JOURNAL_CLUB_ALLOW_WORD_PDF_FALLBACK=1`；普通用户不应启用它。

**如何升级？** 到 [Releases 页面](https://github.com/jiuBABY6/paper-to-journal-club/releases)选择新的固定版本标签，再运行同一安装命令并替换 `ReleaseTag`。安装器会保留此前的 `current` 版本目录，便于人工回退和排障。

**需要反馈问题？** 请在 [GitHub Issues](https://github.com/jiuBABY6/paper-to-journal-club/issues) 提供插件版本、Windows/PowerPoint 版本，以及不含敏感论文内容的最小复现步骤。

## 🔐 隐私与安全

论文解析和 PowerPoint 自动化在本机完成；插件不运营独立的论文或演示稿上传服务器，也不会主动收集遥测。上传到 Codex 的论文、提示词和结果仍受你的 Codex、模型提供方及组织策略约束。详见 [PRIVACY.md](PRIVACY.md) 和 [SECURITY.md](SECURITY.md)。

## ⚖️ 许可证

本项目采用 [MIT License](LICENSE)。第三方组件和许可证说明见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

## 🛠️ 发布与维护

发布、版本标签、SHA-256、代码签名、离线部署与 GitHub Actions 的维护说明见 [PUBLISHING.md](PUBLISHING.md)。
