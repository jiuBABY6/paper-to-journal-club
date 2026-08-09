# 发布资源

正式发布包应包含 `paper-parser.exe`。它由 `parser/PaperParser.csproj` 构建为 Windows x64 自包含单文件，终端用户不需要安装 Node.js、Python 或 .NET。

维护者在有 .NET 8 SDK 的构建机执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\scripts\build-paper-parser.ps1
```

生成物为 `assets/paper-parser.exe`。正式发行版必须携带该文件；若开发检出中缺失它，插件才会尝试已安装的 Microsoft Word 作为 PDF 文本提取回退。正式包不会依赖此回退。
