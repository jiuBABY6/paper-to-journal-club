<#
  不联网的安装器布局回归测试。

  此测试不下载 GitHub Release，也不启动 PowerPoint 或注册 Codex；它验证本地 Marketplace
  根目录、GitHub 发布材料，以及 install.ps1 的 -WhatIf 安全路径。
#>
[CmdletBinding()]
param(
    [string]$RepositoryRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $RepositoryRoot = Split-Path -Parent $PSScriptRoot
}
$RepositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot -ErrorAction Stop).Path

function Assert-True {
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-FileExists {
    param([Parameter(Mandatory)][string]$Path)

    Assert-True -Condition (Test-Path -LiteralPath $Path -PathType Leaf) -Message "缺少文件：$Path"
}

$pluginName = 'paper-to-journal-club'
$marketplaceName = 'paper-to-journal-club-tools'
$installPowerShellPath = Join-Path $RepositoryRoot 'install.ps1'
$publishingPath = Join-Path $RepositoryRoot 'PUBLISHING.md'
$workflowPath = Join-Path $RepositoryRoot '.github\workflows\release.yml'
$lockWorkflowPath = Join-Path $RepositoryRoot '.github\workflows\generate-parser-lock.yml'
$marketplacePath = Join-Path $RepositoryRoot '.agents\plugins\marketplace.json'
$pluginRoot = Join-Path $RepositoryRoot "plugins\$pluginName"
$releaseToolPaths = @(
    (Join-Path $RepositoryRoot 'tools\Build-ReleasePackage.ps1'),
    (Join-Path $RepositoryRoot 'tools\Install-PaperToJournalClub.ps1'),
    (Join-Path $RepositoryRoot 'tools\Test-ReleasePackage.ps1'),
    (Join-Path $RepositoryRoot 'tools\Generate-ParserPackageLock.ps1'),
    (Join-Path $RepositoryRoot 'tools\INSTALL.md')
)

foreach ($path in @($installPowerShellPath, $publishingPath, $workflowPath, $lockWorkflowPath, $marketplacePath, (Join-Path $pluginRoot '.codex-plugin\plugin.json'), (Join-Path $pluginRoot '.mcp.json')) + $releaseToolPaths) {
    Assert-FileExists -Path $path
}

# Windows PowerShell 5.1 需要 UTF-8 BOM 才能可靠读取中文字符串和注释。
$installBytes = [System.IO.File]::ReadAllBytes($installPowerShellPath)
Assert-True -Condition ($installBytes.Length -ge 3 -and $installBytes[0] -eq 0xEF -and $installBytes[1] -eq 0xBB -and $installBytes[2] -eq 0xBF) -Message 'install.ps1 必须采用 UTF-8 BOM 编码。'

foreach ($scriptPath in @($installPowerShellPath) + @($releaseToolPaths | Where-Object { $_ -like '*.ps1' })) {
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
    $parseErrorMessages = @($parseErrors | ForEach-Object { $_.Message })
    Assert-True -Condition ($parseErrorMessages.Count -eq 0) -Message ("$scriptPath 存在 PowerShell 语法错误：{0}" -f ($parseErrorMessages -join '; '))
}

$marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-True -Condition ($marketplace.name -eq $marketplaceName) -Message 'Marketplace 名称不正确。'
$entries = @($marketplace.plugins | Where-Object { $_.name -eq $pluginName })
Assert-True -Condition ($entries.Count -eq 1) -Message 'Marketplace 必须恰好包含一个目标插件条目。'
Assert-True -Condition ($entries[0].source.source -eq 'local' -and $entries[0].source.path -eq "./plugins/$pluginName") -Message 'Marketplace 必须指向插件的相对本地路径。'
Assert-True -Condition (@($marketplace.PSObject.Properties.Name | Sort-Object) -join ',' -ceq 'interface,name,plugins') -Message 'Marketplace 顶层字段必须为严格允许清单。'
Assert-True -Condition (@($marketplace.plugins).Count -eq 1) -Message 'Marketplace 不得包含额外插件条目。'
Assert-True -Condition (@($entries[0].PSObject.Properties.Name | Sort-Object) -join ',' -ceq 'category,name,policy,source') -Message 'Marketplace 插件字段必须为严格允许清单。'

$mcpPath = Join-Path $pluginRoot '.mcp.json'
$mcp = Get-Content -LiteralPath $mcpPath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-True -Condition (@($mcp.PSObject.Properties.Name) -join ',' -ceq 'mcpServers') -Message 'MCP 顶层字段必须为严格允许清单。'
Assert-True -Condition (@($mcp.mcpServers.PSObject.Properties.Name) -join ',' -ceq $pluginName) -Message 'MCP 不得包含额外 server。'
$mcpServer = $mcp.mcpServers.$pluginName
Assert-True -Condition ($mcpServer.command -eq 'powershell.exe' -and $mcpServer.cwd -eq '.') -Message 'MCP 必须使用固定的本地 PowerShell 入口。'
$expectedMcpArguments = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', './scripts/paper-to-journal-club-server.ps1')
Assert-True -Condition ((@($mcpServer.args) -join "`n") -ceq ($expectedMcpArguments -join "`n")) -Message 'MCP 参数必须严格固定，不能追加其他命令。'

$pluginManifest = Get-Content -LiteralPath (Join-Path $pluginRoot '.codex-plugin\plugin.json') -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-True -Condition (@($pluginManifest.PSObject.Properties.Name | Sort-Object) -join ',' -ceq 'author,description,interface,keywords,license,mcpServers,name,skills,version') -Message 'plugin.json 顶层字段必须为严格允许清单。'
Assert-True -Condition ($pluginManifest.skills -eq './skills/' -and $pluginManifest.mcpServers -eq './.mcp.json') -Message 'plugin.json 只能声明本地 skills 与固定 MCP 文件。'
Assert-True -Condition (@($pluginManifest.interface.PSObject.Properties.Name | Sort-Object) -join ',' -ceq 'capabilities,category,defaultPrompt,developerName,displayName,longDescription,shortDescription') -Message 'plugin.json interface 字段必须为严格允许清单。'

$installSource = [System.IO.File]::ReadAllText($installPowerShellPath, [System.Text.UTF8Encoding]::new($true))
Assert-True -Condition ($installSource -match 'RepositoryUrl' -and $installSource -match 'ReleaseTag' -and $installSource -match 'SupportsShouldProcess') -Message 'install.ps1 必须要求仓库地址和固定 Release 标签，并支持 -WhatIf。'
Assert-True -Condition ($installSource -notmatch '(?i)ExecutionPolicy\s+Bypass') -Message 'install.ps1 不得引入 ExecutionPolicy Bypass。'
foreach ($requiredFragment in @('Test-ChecksumManifest', 'Assert-SafeZipArchive', 'Expand-SafeZipArchive', 'Assert-AllowedMarketplaceAndMcpConfiguration', 'MaximumExpandedBytes')) {
    Assert-True -Condition ($installSource.Contains($requiredFragment)) -Message "install.ps1 缺少发行链安全检查：$requiredFragment"
}
Assert-True -Condition ($installSource -notmatch '(?m)^\s*Expand-Archive\b') -Message 'install.ps1 不得绕过受限 ZIP 解压器直接调用 Expand-Archive。'

$publishingSource = Get-Content -LiteralPath $publishingPath -Raw -Encoding UTF8
Assert-True -Condition ($publishingSource -match '\.zip\.sha256' -and $publishingSource -match 'bootstrap\.ps1' -and $publishingSource -match 'RemoteSigned') -Message 'PUBLISHING.md 必须说明外部 SHA-256、引导安装器与 RemoteSigned。'

$workflowSource = Get-Content -LiteralPath $workflowPath -Raw -Encoding UTF8
foreach ($requiredFragment in @('actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5', 'actions/setup-dotnet@67a3573c9a986a3f9c594539f4ab511d57bb3ce9', 'actions/upload-artifact@ea165f8d65b6e75b540449e92b4886f43607fa02', 'actions/download-artifact@d3f86a106a0bac45b974a628896c90dbdf5c8093', 'Get-FileHash', 'bootstrap.ps1', 'gh release create')) {
    Assert-True -Condition ($workflowSource.Contains($requiredFragment)) -Message "release.yml 缺少预期发布步骤：$requiredFragment"
}
Assert-True -Condition ($workflowSource.Contains('tools/Build-ReleasePackage.ps1')) -Message 'release.yml 必须调用 tools/Build-ReleasePackage.ps1。'
Assert-True -Condition (-not $workflowSource.Contains('release/Build-ReleasePackage.ps1')) -Message 'release.yml 不得调用迁移前的 release/Build-ReleasePackage.ps1。'
Assert-True -Condition (-not $workflowSource.Contains('--clobber')) -Message 'release.yml 不得覆写已有 Release 资产。'
Assert-True -Condition ($workflowSource.Contains('persist-credentials: false')) -Message 'release.yml 必须禁止 checkout 持久化 GitHub 凭据。'
Assert-True -Condition ($workflowSource.Contains("dotnet-version: '8.0.418'")) -Message 'release.yml 必须固定使用经过审核的 .NET SDK 8.0.418。'
Assert-True -Condition ($workflowSource.Contains('build-and-package:') -and $workflowSource.Contains('create-release:')) -Message 'release.yml 必须将只读构建与可写发布拆分为两个 job。'

$lockWorkflowSource = Get-Content -LiteralPath $lockWorkflowPath -Raw -Encoding UTF8
foreach ($requiredFragment in @('workflow_dispatch:', 'permissions:', 'contents: read', 'actions/checkout@34e114876b0b11c390a56381ad16ebd13914f8d5', 'actions/setup-dotnet@67a3573c9a986a3f9c594539f4ab511d57bb3ce9', 'actions/upload-artifact@ea165f8d65b6e75b540449e92b4886f43607fa02', 'Generate-ParserPackageLock.ps1', "dotnet-version: '8.0.418'")) {
    Assert-True -Condition ($lockWorkflowSource.Contains($requiredFragment)) -Message "generate-parser-lock.yml 缺少受控锁文件生成步骤：$requiredFragment"
}

$parserProjectSource = Get-Content -LiteralPath (Join-Path $pluginRoot 'parser\PaperParser.csproj') -Raw -Encoding UTF8
$parserTestSource = Get-Content -LiteralPath (Join-Path $pluginRoot 'tests\parser-resource-limit-tests.ps1') -Raw -Encoding UTF8
Assert-True -Condition ($parserProjectSource.Contains('<RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>') -and $parserProjectSource.Contains('<RestoreLockedMode>true</RestoreLockedMode>')) -Message '解析器项目必须启用 NuGet lock file 与 locked mode。'
Assert-True -Condition ($parserTestSource.Contains('Test-ParserPackageLock') -and $parserTestSource.Contains('RequireLock')) -Message '解析器测试必须验证受控 NuGet 锁文件。'

# -WhatIf 必须在请求网络或创建目录前退出，因而可用于 CI 中的安全预检。
$unusedDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("paper-to-journal-club-install-test-{0}" -f [Guid]::NewGuid().ToString('N'))
$whatIfOutput = & $installPowerShellPath -RepositoryUrl 'https://github.com/test-owner/test-repository.git' -ReleaseTag 'v1.0.0' -InstallDirectory $unusedDirectory -WhatIf
$whatIf = $whatIfOutput | ConvertFrom-Json
Assert-True -Condition ($whatIf.pass -eq $true -and $whatIf.action -eq 'would-download-and-install-release' -and $whatIf.node_required -eq $false) -Message 'install.ps1 的 -WhatIf 输出不符合预期。'
Assert-True -Condition (-not (Test-Path -LiteralPath $unusedDirectory)) -Message '-WhatIf 不得创建安装目录。'

# 安装目录即使尚未真正安装内容，也不能穿过 junction/symlink。此处使用实际 Junction
# 验证 -WhatIf 会在访问网络或创建文件前拒绝它；清理时先删除链接本身，绝不递归跟随链接。
$junctionTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("paper-to-journal-club-junction-test-{0}" -f [Guid]::NewGuid().ToString('N'))
$junctionTarget = Join-Path $junctionTestRoot 'target'
$junctionInstallDirectory = Join-Path $junctionTestRoot 'install-link'
try {
    New-Item -ItemType Directory -Force -Path $junctionTarget | Out-Null
    New-Item -ItemType Junction -Path $junctionInstallDirectory -Target $junctionTarget -ErrorAction Stop | Out-Null
    $junctionError = $null
    try {
        & $installPowerShellPath -RepositoryUrl 'https://github.com/test-owner/test-repository.git' -ReleaseTag 'v1.0.0' -InstallDirectory $junctionInstallDirectory -WhatIf | Out-Null
    } catch {
        $junctionError = $_
    }
    Assert-True -Condition ($null -ne $junctionError -and $junctionError.Exception.Message -match '重解析点') -Message '安装器必须在 -WhatIf 阶段拒绝包含 junction/symlink 的 InstallDirectory。'
} finally {
    if (Test-Path -LiteralPath $junctionInstallDirectory) {
        Remove-Item -LiteralPath $junctionInstallDirectory -Force
    }
    if (Test-Path -LiteralPath $junctionTarget) {
        Remove-Item -LiteralPath $junctionTarget -Recurse -Force
    }
    if (Test-Path -LiteralPath $junctionTestRoot) {
        Remove-Item -LiteralPath $junctionTestRoot -Force
    }
}

Write-Host 'PASS: install-layout-tests.ps1'
