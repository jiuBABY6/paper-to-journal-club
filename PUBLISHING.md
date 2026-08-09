# GitHub 发布与本地 Marketplace 分发

Paper to Journal Club 是一个 **Windows 本地 Codex Marketplace 插件**：运行时由本机 PowerShell MCP 通过 COM 控制已安装的 Microsoft PowerPoint。它不是公网 MCP 服务，也不需要 Node.js、npm、Python、Git 或 .NET SDK 才能让普通用户使用。

本目录本身就是 Marketplace 根目录：

```text
paper-to-journal-club/
├── .agents/plugins/marketplace.json
├── plugins/paper-to-journal-club/
└── install.ps1
```

## 发布者首次准备

将 `demo/paper-to-journal-club` 单独作为一个 GitHub 仓库的根目录后再发布。这样本目录中的 `.github/workflows/release.yml` 才会被 GitHub Actions 识别。不要仅把它作为更大仓库中的子目录推送后，期待子目录内的工作流自动运行。

发布者需要具备：

- GitHub 仓库的 Release 发布权限；
- Windows x64 构建环境或 GitHub Actions 的 Windows runner；
- .NET 8 SDK，仅用于构建随发行包附带的 `paper-parser.exe`；
- 用于最终质量验收的 Windows 桌面版 Microsoft PowerPoint；
- 可选但强烈建议：组织代码签名证书。

终端用户不需要上述构建依赖，只需要 Windows x64、桌面版 Microsoft PowerPoint 和 Codex 桌面版/CLI。

## GitHub Release 构建流程

1. 更新 `plugins/paper-to-journal-club/.codex-plugin/plugin.json` 中的严格 SemVer 版本号。
2. 在干净的 Windows 环境中运行插件测试和发行包验证；没有本机 .NET SDK 时，先在 GitHub **Actions → 发布 Paper to Journal Club 本地 Marketplace → Run workflow** 中选择 `main` 运行手动预检。该预检会完整构建解析器、生成并离线验证 ZIP，但只上传临时 Actions 工件，**不会创建或修改 GitHub Release**。
3. 预检全绿后，创建与清单版本严格对应的 Git 标签，例如当前清单为 `1.0.1` 时，标签必须为 `v1.0.1`。
4. 推送该标签。`release.yml` 会构建 PDF 解析器、创建本地 Marketplace ZIP、生成 ZIP 的外部 SHA-256 文件，并创建同名 GitHub Release。

工作流固定 `actions/checkout` 与 `actions/setup-dotnet` 到完整提交 SHA，禁用 checkout 凭据持久化。构建/打包 job 仅有 `contents: read`，它只把四个已校验的最终资产交给独立的发布 job；后者不执行仓库源码，才获得创建 Release 所需的 `contents: write` 权限。SDK 也固定为 `8.0.418`。同一标签若已有 Release，工作流会失败，**绝不会**用 `--clobber` 覆写旧资产；修复必须发布新版本并使用新标签。

每次正式 Release 都必须由 CI 从当前标签源码重新构建 `paper-parser.exe`；打包器和包内验证器会实际调用 `paper-parser.exe --build-info`，要求 PdfPig 为 `0.1.15` 且资源预算完全匹配。不得复用旧的 1.0.0/0.1.13 二进制，即使它具有有效的 MZ 文件头。

### 首次生成并提交 NuGet 依赖锁

`parser/packages.lock.json` 是解析器可复现构建的一部分，**必须**由固定的 .NET SDK 8.0.418 和当前 `NuGet.Config` 真实 restore 生成；不要手写、复制旧 EXE 的信息或提交空锁文件。正式构建与打包会启用 `--locked-mode`，缺少锁文件或项目与锁不一致时会直接失败。

有受控 Windows 构建环境时，运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\tools\Generate-ParserPackageLock.ps1
git diff -- plugins/paper-to-journal-club/parser/packages.lock.json
git add plugins/paper-to-journal-club/parser/packages.lock.json
```

先审阅锁文件中的版本与 SHA-512 `contentHash`，再提交。若需要有意更新依赖，才使用 `-Force` 重新生成，并在单独的依赖更新提交中审阅变更。

没有本机 SDK 时，可在 GitHub 的 **Actions → 生成 Paper Parser NuGet 锁文件 → Run workflow** 手动运行只读工作流。它只上传 `packages.lock.json` artifact，不会提交或推送；发布者下载、审阅并手工提交该文件后，才可创建 Release 标签。

### GitHub 仓库保护（发布者必须在网页端完成）

代码不能替代 GitHub 仓库权限设置。请在仓库的 Rulesets/Tag protection 中为 `v*` 标签设置规则：只允许指定发布维护者创建、更新或删除匹配标签；同时保护默认分支，并限制能够修改 `.github/workflows/` 的人员。这样，未经审核的标签不能触发带发布权限的工作流。完成设置后，在发布前用一个非正式测试标签验证规则是否如预期拒绝未授权操作。

如需在推送标签前做本地发行验证，先安装 .NET 8 SDK，然后在本目录执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\plugins\paper-to-journal-club\scripts\build-paper-parser.ps1
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\tools\Build-ReleasePackage.ps1 -OutputDirectory .\dist
```

构建器会在 `dist/` 生成 ZIP 和同名目录，并自动运行包内 `verify-release.ps1`。`dist/` 是可再生成发行产物，不应提交到仓库。

发布的 Release 必须同时包含：

```text
paper-to-journal-club-marketplace-<version>.zip
paper-to-journal-club-marketplace-<version>.zip.sha256
paper-to-journal-club-bootstrap.ps1
paper-to-journal-club-bootstrap.ps1.sha256
```

安装器只接受成对存在的 ZIP 与 `.sha256` 文件，并拒绝草稿和预发布 Release。它先校验 ZIP，再以受限的流式解压器展开；解压后，`SHA256SUMS.txt` 会做**双向集合校验**：清单中的每个文件必须存在且哈希一致，目录中除清单自身外的每个文件也必须被清单列出。随后安装器只接受唯一的本地 Marketplace 条目和固定的 PowerShell MCP 参数。缺少校验文件、额外文件、额外 MCP server 或不安全 ZIP 都会被拒绝。引导安装器也有独立哈希，用户应在执行前校验它。

为防 ZIP 资源炸弹，当前安装器限制 ZIP 最多 256 个条目、单个文件最多 256 MiB、总实际解压量最多 1 GiB，且拒绝异常压缩比、路径穿越、绝对路径和大小写冲突路径。

## 普通用户的安装方式

发布者应在自己的 Release 页面、实验室内网或 README 中提供以下模板，并把尖括号中的内容替换为真实值：

1. 从同一 GitHub Release 下载 `paper-to-journal-club-bootstrap.ps1` 及其 `.sha256` 文件；
2. 比对引导脚本 SHA-256，确认它与发布者公布的值一致；

   ```powershell
   Get-FileHash -LiteralPath .\paper-to-journal-club-bootstrap.ps1 -Algorithm SHA256
   ```

3. 若 Windows 标记该已验证脚本来自互联网且发布者未提供受信任签名，可先执行 `Unblock-File -LiteralPath .\paper-to-journal-club-bootstrap.ps1`；
4. 在脚本所在目录执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\paper-to-journal-club-bootstrap.ps1 `
  -RepositoryUrl '<发布者的 GitHub HTTPS 仓库地址>' `
  -ReleaseTag 'v<发行版本>'
```

`install.ps1` 会执行以下工作：

1. 从指定仓库的指定标签读取 GitHub Release；
2. 下载 ZIP 和同 Release 的 `.sha256` 文件；
3. 校验 SHA-256，拒绝损坏或不匹配的 ZIP；
4. 解压到当前用户的稳定目录，默认位置为：

   ```text
   %LOCALAPPDATA%\Codex\marketplaces\paper-to-journal-club\current
   ```

5. 调用包内安装器执行离线结构校验、PowerPoint 前置检查，并运行：

   ```text
   codex plugin marketplace add <本地 Marketplace 根目录>
   codex plugin add paper-to-journal-club@paper-to-journal-club-tools
   ```

升级时使用新的固定 Release 标签再次运行同一命令。安装器不会删除旧版本：会把旧的 `current` 目录保留为 `previous-...`，以便排障或人工回退。

## 离线或受管网络安装

如果单位不允许安装器访问 GitHub API，管理员可以从已验证的 GitHub Release 下载 ZIP 和对应 `.sha256` 文件，在离线环境校验 SHA-256 后解压。随后在 **解压包根目录** 中运行包内的 `install.ps1`。

不要把 Marketplace 根目录指向 `plugins/paper-to-journal-club`；正确的根目录同时包含 `.agents` 和 `plugins` 两个目录。

## RemoteSigned、下载标记与代码签名

MCP 使用 `RemoteSigned`，不会使用 `ExecutionPolicy Bypass`。浏览器下载的 ZIP 或安装脚本可能带有 Windows 的 Mark-of-the-Web：

- 优先发布经过 Authenticode 签名的 PowerShell 脚本和 `paper-parser.exe`；
- 若组织策略允许且用户已核验 Release 来源和 SHA-256，可在 **已验证的解压目录** 内执行 `Get-ChildItem -Recurse -File | Unblock-File`；
- 不要为了绕过企业策略修改系统执行策略，也不要对来源不明的目录执行 `Unblock-File`。

企业环境应让 IT 管理员审查并部署受信任签名证书。若组策略禁止未签名脚本，安装器会如实失败，而不是尝试绕过控制。

代码签名是公开或企业分发的强烈建议，但当前构建器不强制维护者拥有证书：只有显式传入 `-CodeSigningCertificateThumbprint` 时，才会预检当前用户证书存储并签署运行时脚本与 EXE；未提供该参数时，构建输出会如实标明未启用签名。SHA-256 只能保证与已获得的哈希值一致，不能独立证明发布者身份，因此应在可信 Release 页面、签名或独立发布渠道中核对哈希。

## 发布前检查清单

- GitHub Release 标签与插件清单版本匹配；
- Release 同时有 ZIP 与对应 `.zip.sha256`；
- ZIP 内含 `.agents/plugins/marketplace.json`、`plugins/paper-to-journal-club`、自包含 `paper-parser.exe` 和 `SHA256SUMS.txt`；
- `SHA256SUMS.txt` 与发行目录通过双向集合校验，且 ZIP 解压限制、Marketplace/MCP 严格 allowlist 均通过；
- CI 使用完整 SHA 固定 Actions，且不存在 `--clobber` 或可变版本标签覆盖旧 Release；
- `verify-release.ps1` 在不安装 Codex 和不启动 PowerPoint 的 CI 环境中通过；
- 在干净 Windows x64 机器上验证安装、PowerPoint 生成、重新打开审计和导出预览；
- `plugin.json` 中的作者、支持渠道、隐私政策与许可证信息已替换为真实发布主体；
- 发布者已说明本插件是本地 Marketplace 分发，不承诺自动进入第三方公开插件目录。

## 安装器自检与故障排查

安装器支持不联网的参数和布局检查：

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\install.ps1 `
  -RepositoryUrl '<发布者的 GitHub HTTPS 仓库地址>' `
  -ReleaseTag 'v<发行版本>' `
  -WhatIf
```

如只想下载、解压并验证包结构而暂不注册到 Codex，可在真实安装命令中追加：

```powershell
-SkipCodexInstall -SkipPowerPointCheck
```

这些开关仅适合 CI、维护者排障或管理员预部署；普通科研用户应保留默认检查，并在安装后重启 Codex、开启一个新任务再调用插件。
