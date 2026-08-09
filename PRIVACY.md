# 隐私说明 / Privacy

## 中文

Paper to Journal Club 的 PDF 文本提取和 Microsoft PowerPoint 自动化在用户的 Windows 设备上完成：

- 插件不运营独立的论文、演示稿或图像上传服务器；
- `paper-parser.exe` 仅处理用户指定的本地论文文件；
- PDF 图片资产默认写入用户临时目录下的 `paper-to-journal-club` 专用子目录，供用户/Agent 审阅与插图；完成后可用 `cleanup_paper_assets` 并显式确认删除该目录；
- PowerPoint 操作通过本机 PowerShell 与 Microsoft Office COM 完成；
- 插件不主动收集遥测、账号信息或使用统计；
- 安装脚本只调用本机 `codex plugin` 命令注册和安装 Marketplace，不会更改系统执行策略；
- 用户上传给 Codex 的论文、提示词、工具结果和生成内容仍受用户选择的 Codex、模型提供方及组织政策约束，这些服务不由本插件控制。

论文可能包含未公开数据、个人信息或受版权保护的材料。请在上传、分享生成的 PPTX 或向第三方发送日志前自行确认授权与保密要求。

## English

Paper to Journal Club performs PDF text extraction and Microsoft PowerPoint automation locally on the user's Windows device.

- The plugin does not operate a separate server for papers, presentations, or images.
- `paper-parser.exe` processes only the local paper selected by the user.
- Extracted PDF image assets are stored by default in a dedicated `paper-to-journal-club` subdirectory of the user's temp folder for review and insertion; `cleanup_paper_assets` removes only that confirmed temporary directory.
- PowerPoint automation uses local PowerShell and Microsoft Office COM.
- The plugin does not intentionally collect telemetry, account data, or usage analytics.
- The installer only invokes local `codex plugin` commands and does not modify the system execution policy.
- Papers, prompts, tool output, and generated content provided to Codex remain subject to the user's Codex, model-provider, and organization policies, which are outside this plugin's control.

Users remain responsible for permissions, confidentiality, and copyright before uploading or sharing research material.
