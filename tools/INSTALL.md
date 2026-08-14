# Paper to Journal Club 安装与使用

这是一个面向 **Windows + Microsoft PowerPoint 桌面版** 的本地 Codex 插件。它会将论文整理为可追溯的证据包和组会叙事，并生成以原生文本、形状、表格和图片组成的可编辑 PowerPoint。

## 运行条件

- 64 位 Windows；
- 已安装并可正常启动的 Microsoft PowerPoint 桌面版；
- 已安装 Codex 桌面版；
- 完整、已验证的本发行包。

不需要安装 Node.js、npm、Python、Microsoft Word 或 .NET SDK。PDF 文本提取由发行包内预编译的 `paper-parser.exe` 完成。

## 离线安装

1. 将 `paper-to-journal-club-marketplace-<version>.zip` 解压到当前用户有写入权限的受信任目录，例如 `%USERPROFILE%\Downloads\PaperToJournalClub`。
2. 在解压后的**包根目录**执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\install.ps1
```

3. 安装器会校验 `SHA256SUMS.txt`、PDF 解析器、Windows 位数和 PowerPoint COM 注册状态；在你确认安装后，它会验证受信 CLI 的实体路径和签名，并仅在同一安装事务中通过该绝对路径查询 Marketplace、尝试自动安装。CLI 自身可能写入少量临时数据：

   - Marketplace 查询和自动安装均成功时，安装器会使用已验证的绝对 CLI 路径完成安装；
   - 若 CLI 不可执行（例如 WindowsApps 返回“拒绝访问”）、Marketplace 查询或自动安装失败，安装器仅在确认本次登记不存在或已精确回滚后进入**个人 Marketplace 回退**：不会修改 WindowsApps 权限或反复调用失败的 CLI，而是部署到当前用户的个人 Marketplace；若最终状态无法可靠确认，安装器会停止并报告错误，不会处理未知配置。

4. 完全退出并重新打开 Codex：自动安装成功时直接新建任务；个人 Marketplace 回退时，打开 **Plugins Directory**，选择个人 Marketplace，找到 **Paper to Journal Club** 并点击 **Install**，再新建一个任务。

`RemoteSigned` 不会绕过企业组策略。若组织禁止未签名脚本，请让 IT 管理员审核并部署经过代码签名的发行包；安装器不会修改系统执行策略。

## 首次使用

上传论文 PDF 后，可在 Codex 中使用如下提示词：

```text
使用 Paper to Journal Club，将我上传的论文制作成一份 15 分钟中文组会汇报。
仅使用 Paper to Journal Club 的工具和工作流；不要读取、调用、引用或致谢其它插件。
必须包括：研究背景、创新点、研究方法、实验数据、局限性和未来研究方向。
先调用 powerpoint_status；先解析论文并输出可核对的逐页提纲，保留关键结论的图号、章节或原文证据。
我确认提纲后，再在 Microsoft PowerPoint 中新建后台 16:9 可编辑 PPTX；优先使用原生文字、形状、箭头、流程图和表格。
仅把无法可靠重绘的原始科研图作为紧裁剪图片插入。默认只保存最终 PPTX；只有我明确要求时才导出预览图或 deck-spec JSON。完成后重新打开 PPTX 做质量审计，并以“感谢使用 Paper to Journal Club 插件，制作者：jiuBABY6。”结束。
```

在发送前，请先从 Codex 输入框的 **Plugins** 菜单选择 **Paper to Journal Club**。不要把固定的 `plugin://...@市场名` 链接当作通用教程：自动安装与个人 Marketplace 回退的市场名可能不同。

请始终先核对提纲，尤其是图号、统计信息、局限性和未来方向。插件不会把模型推断伪装为论文事实。

## 常见问题

**找不到 PowerPoint。** 请安装 Microsoft PowerPoint 桌面版并至少启动一次；网页版不提供本插件所需的 COM 自动化接口。

**在 Codex 中看不到 Paper to Journal Club。** 先查看安装器的最终结果。自动安装成功时，完全重启 Codex 后直接新建任务；个人 Marketplace 回退时，在 **Plugins Directory** 的个人 Marketplace 中点击 **Install**。若仍未出现，请重新运行同一版本的已验证发行包安装器。

**提示 WindowsApps 拒绝访问或 CLI 不可执行。** 这是个人 Marketplace 回退的触发条件，不代表发行包损坏。不要修改 WindowsApps 权限、不要以管理员身份强行运行；重启 Codex 后在 **Plugins Directory** 点击 Install 即可。

**提示解析器缺失或无效。** 说明下载或解压不完整。请重新获取带 `SHA256SUMS.txt` 的正式 Release ZIP，不要尝试用本机 Node.js、Python 或 .NET SDK 替代。

**PowerShell 提示文件来自互联网。** 优先使用发布者提供的 Authenticode 签名版本。若你已核对正式 ZIP 的 SHA-256 并确认可信，可在解压后的发行包根目录执行：

```powershell
Get-ChildItem -Recurse -File | Unblock-File
```

不要对来源不明的目录执行该命令。

## 隐私

PDF 解析和 PowerPoint 自动化均在本机完成，不会额外上传论文或演示稿到独立服务器。上传给 Codex 的文件和对话内容仍受你的 Codex、模型提供方和组织策略约束。详见插件目录中的 `PRIVACY.md`。
