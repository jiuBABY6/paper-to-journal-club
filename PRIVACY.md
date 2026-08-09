# 隐私说明 / Privacy

## 中文

Paper to Journal Club 的 PDF 文本提取和 Microsoft PowerPoint 自动化在用户的 Windows 设备上完成：

- 插件不运营独立的论文、演示稿或图像上传服务器；
- `paper-parser.exe` 仅处理用户指定的本地论文文件；
- 为避免通过工具调用读取任意本机文件，论文和输出默认仅限于“桌面”“文档”“下载”及插件专用临时目录；用户如确有资料盘需求，须在启动 Codex 前以 `PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS` 显式声明窄目录，且不应将磁盘根目录、整个用户目录或 `AppData` 设为允许根；
- PDF 图片资产默认写入用户临时目录下的 `paper-to-journal-club` 专用子目录，供用户/Agent 审阅与插图；完成后可用 `cleanup_paper_assets` 并显式确认删除该目录；
- PowerPoint 操作通过本机 PowerShell 与 Microsoft Office COM 完成；
- 插件不主动收集遥测、账号信息或使用统计；
- 引导安装脚本只访问用户指定的 GitHub Release 以下载公开发行资产和校验文件，不会上传论文、演示稿或凭据；通过校验后，它只调用本机 `codex plugin` 命令注册和安装 Marketplace，且不会更改系统执行策略；
- 用户上传给 Codex 的论文、提示词、工具结果和生成内容仍受用户选择的 Codex、模型提供方及组织政策约束，这些服务不由本插件控制。

论文可能包含未公开数据、个人信息或受版权保护的材料。请在上传、分享生成的 PPTX 或向第三方发送日志前自行确认授权与保密要求。

## English

Paper to Journal Club performs PDF text extraction and Microsoft PowerPoint automation locally on the user's Windows device.

- The plugin does not operate a separate server for papers, presentations, or images.
- `paper-parser.exe` processes only the local paper selected by the user.
- To avoid arbitrary local-file access through tool calls, paper and output paths are limited by default to Desktop, Documents, Downloads, and the plugin's dedicated temp directory. Users with a research-data drive must explicitly set a narrow `PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS` directory before starting Codex; drive roots, an entire user profile, and `AppData` must not be used as approved roots.
- Extracted PDF image assets are stored by default in a dedicated `paper-to-journal-club` subdirectory of the user's temp folder for review and insertion; `cleanup_paper_assets` removes only that confirmed temporary directory.
- PowerPoint automation uses local PowerShell and Microsoft Office COM.
- The plugin does not intentionally collect telemetry, account data, or usage analytics.
- The bootstrap installer only requests the user-selected public GitHub Release to download release assets and checksums; it does not upload papers, presentations, or credentials. After verification, it invokes local `codex plugin` commands and does not modify the system execution policy.
- Papers, prompts, tool output, and generated content provided to Codex remain subject to the user's Codex, model-provider, and organization policies, which are outside this plugin's control.

Users remain responsible for permissions, confidentiality, and copyright before uploading or sharing research material.
