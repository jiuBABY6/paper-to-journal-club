# Notices

`assets/paper-parser.exe` 由 `parser/PaperParser.csproj` 使用 Microsoft .NET 8 发布为 Windows x64 自包含单文件程序。它在发行包中携带所需的 .NET 运行时组件，因此终端用户无需单独安装 .NET SDK 或运行时。

发布者在分发该文件时应遵守适用的 Microsoft .NET 许可、出口管制、代码签名和组织软件分发政策。解析器使用 PdfPig 0.1.13（Apache-2.0）提取 PDF 文本；其许可证与归属说明见 `THIRD_PARTY_NOTICES.md`。如后续加入依赖，必须同步更新该文件和第三方通知。
