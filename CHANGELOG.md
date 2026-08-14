# Changelog

## Unreleased

- 将组会叙事收紧为“研究问题 → 背景缺口 → 创新 → 研究设计与系统 → 实验依据 → 局限 → 下一步 → 结论”；审核器会阻断顺序混乱、结论未收尾或证据不完整的 deck spec。
- 非封面标题改为短陈述句并禁止问号和冒号；收紧标题、要点和结论文字长度，避免用缩小字号掩盖内容过载。
- 新增系统结构页与实验图表页专用版式：前者将原始系统图与论文方法解释并列，后者将原始图表与来源可追溯的比较、解释、限制并列。无可靠图片时改为全宽文字，不再生成空白“待插图”卡片。
- 新增 `render_paper_visual`：在 Windows 内置 PDF 渲染能力下，将已审核的 Figure/Table 来源页渲染为受限 PNG 候选。裁剪必须经用户确认；渲染图或表格不会自动插入，也不会据截图虚构数值或优劣结论。
- 删除未使用的通用占位卡、证据标签和旧文本辅助函数，减少生成规则冲突。

## 1.0.3

- 新增 `export_figure_assets=true`：按需将本次真正插入 PowerPoint 的合规论文原图复制到 `<PPT名>_assets/images`；默认不创建该目录，也不导出未使用的候选图片。
- 同页只有一个论文图号且只提取到一张合规栅格图片时，自动将实验结果图或系统结构图插入对应 PPT 页面；多图页、无图或关联不确定时不猜测，也不再创建图片占位卡。
- 默认只在用户指定的目录保存最终 PPTX；deck-spec JSON 和 PowerPoint PNG 预览改为显式请求，避免污染论文或汇报项目目录。
- 为内部生成工作文件使用受限临时目录，并在生成后清理；新增可选 `deck_spec_output_path`，供需要长期保留逐页提纲与证据引用的用户显式导出。
- 收紧 Paper to Journal Club 技能范围和默认提示词，明确只用于“论文到组会 PPT”流程，不用于科研插图重绘或通用 PPT 编辑；成功生成时返回本插件的署名信息。
- README 改为通过 Plugins 菜单选择插件，并说明安装临时下载、运行时缓存与辅助输出文件的区别。

## 1.0.2

- 将终端用户安装切换为官方个人 Marketplace 流程：安装器校验正式 Release 后，将已验证插件部署到当前用户目录并更新个人 Marketplace；重启 Codex 后由用户在 Plugins 页面点击 Install。
- 不再直接执行 WindowsApps 内部的 `codex.exe`，从而避免“Access is denied”阻断安装；不修改 `config.toml`，也不要求 Node.js、npm、Python、.NET SDK 或管理员权限。
- 为个人 Marketplace 部署增加 PowerShell 5.1 兼容、SHA-256 复核、重解析点拒绝、配置原子替换和中断恢复测试。

## 1.0.1

- 收紧本地 MCP 的文件访问、布尔确认和 JSON-RPC 输入边界，防止提示注入诱导读取任意本机文件或覆盖现有演示稿。
- 为临时目录、图片、PowerPoint/Word COM 和 PDF 解析增加链接、资源和自动化安全保护；Word PDF 回退默认关闭。
- 将随包解析器升级至 PdfPig 0.1.15，并要求发布 CI 重建和验证二进制的资源限制。
- 加固 Release ZIP、安装器和 GitHub Actions：严格校验完整清单、固定依赖、分离构建与发布权限，并拒绝覆盖既有 Release 资产。

## 1.0.0

- 提供无 Node.js 的 PowerShell MCP 运行路径，并通过 `RemoteSigned` 启动策略运行。
- 直接控制 Windows Microsoft PowerPoint，生成可编辑 PPTX、原生 PNG 预览和结构/版式质量报告。
- 默认强制研究背景、创新点、研究方法、实验数据、局限性和未来研究方向六个组会模块；缺少来源或必备模块时阻断生成。
- 新增 `powerpoint_status`、当前窗口只读检查、保存文件只读审计、协议版本协商和显式覆盖保护。
- PDF 解析器改为带文件、页数、图片资产和时间边界的 Windows x64 自包含 EXE，终端用户无需安装 Node.js、Python 或 .NET SDK。
- 增加本地/团队 Marketplace 打包、校验、安装、许可与第三方组件通知流程。
