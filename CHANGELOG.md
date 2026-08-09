# Changelog

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
