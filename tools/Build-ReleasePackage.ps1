<#
  将本仓库打包为可由 GitHub Release 分发的本地 Codex Marketplace。

  源仓库与发行包均采用同一布局：
  <根目录>/.agents/plugins/marketplace.json
  <根目录>/plugins/paper-to-journal-club/

  终端用户不需要 Node.js、npm、Python 或 .NET SDK；.NET 仅在发布时用于预编译
  随包提供的 paper-parser.exe。
#>
[CmdletBinding()]
param(
    # 留空时输出到仓库根目录的 dist/；该目录已由 .gitignore 排除。
    [string]$OutputDirectory,

    # 只供维护者检查打包流程。含此开关的包不完整，严禁分发给用户。
    [switch]$AllowMissingParser,

    # 只构建目录并离线验证，不额外生成 ZIP。
    [switch]$SkipArchive,

    # 可选：使用当前用户证书存储中的代码签名证书签署运行时脚本和 EXE。
    [string]$CodeSigningCertificateThumbprint
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'
$RepositoryRoot = Split-Path -Parent $PSScriptRoot
$PluginRoot = Join-Path $RepositoryRoot "plugins\$PluginName"
$MarketplaceSource = Join-Path $RepositoryRoot '.agents\plugins\marketplace.json'
$PackageTester = Join-Path $PSScriptRoot 'Test-ReleasePackage.ps1'
$PackageInstaller = Join-Path $PSScriptRoot 'Install-PaperToJournalClub.ps1'
$PackageReadme = Join-Path $PSScriptRoot 'INSTALL.md'

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $RepositoryRoot 'dist'
}

function Copy-RequiredItem {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        throw "发布所需文件不存在：$Source"
    }
    Copy-Item -LiteralPath $Source -Destination $Destination -Recurse -Force
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$TargetPath
    )

    $baseUri = [System.Uri]::new((Join-Path $BasePath ''))
    $targetUri = [System.Uri]::new($TargetPath)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString()).Replace('\', '/')
}

function Write-ChecksumManifest {
    param([Parameter(Mandatory)][string]$Root)

    $hashLines = foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File | Sort-Object FullName) {
        if ($file.Name -eq 'SHA256SUMS.txt') {
            continue
        }
        $relative = Get-RelativePath -BasePath $Root -TargetPath $file.FullName
        $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        "$hash  *$relative"
    }
    Set-Content -LiteralPath (Join-Path $Root 'SHA256SUMS.txt') -Value $hashLines -Encoding UTF8
}

function Set-ReleaseSignature {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Thumbprint
    )

    $normalizedThumbprint = $Thumbprint -replace '\s', ''
    $certificate = Get-ChildItem -Path "Cert:\CurrentUser\My\$normalizedThumbprint" -ErrorAction SilentlyContinue
    if ($null -eq $certificate -or -not $certificate.HasPrivateKey) {
        throw "未在 Cert:\CurrentUser\My\$normalizedThumbprint 找到具有私钥的代码签名证书。"
    }

    # 签署所有实际会在用户电脑执行的脚本和 EXE，而不是只签署安装器。
    $filesToSign = @(
        Get-ChildItem -LiteralPath $Root -Recurse -File |
            Where-Object { $_.Extension -in @('.ps1', '.exe') } |
            Select-Object -ExpandProperty FullName
    )
    foreach ($file in $filesToSign) {
        $result = Set-AuthenticodeSignature -FilePath $file -Certificate $certificate
        if ($result.Status -ne 'Valid') {
            throw "代码签名失败：${file}；$($result.StatusMessage)"
        }
    }
}

function Remove-StagedBuildCache {
    param([Parameter(Mandatory)][string]$StagedPluginRoot)

    # 即使开发者遗漏清理，本函数也不会让构建缓存进入正式发行包。
    $assetsRoot = [System.IO.Path]::GetFullPath((Join-Path $StagedPluginRoot 'assets')).TrimEnd('\')
    $cacheRoot = [System.IO.Path]::GetFullPath((Join-Path $assetsRoot 'parser-publish'))
    $requiredPrefix = "$assetsRoot\"
    if (-not $cacheRoot.StartsWith($requiredPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "拒绝删除预期 assets 目录之外的缓存路径：$cacheRoot"
    }
    if (Test-Path -LiteralPath $cacheRoot) {
        Remove-Item -LiteralPath $cacheRoot -Recurse -Force
    }
}

$manifestPath = Join-Path $PluginRoot '.codex-plugin\plugin.json'
$parserPath = Join-Path $PluginRoot 'assets\paper-parser.exe'
$requiredInputs = @(
    $manifestPath,
    $MarketplaceSource,
    $PackageTester,
    $PackageInstaller,
    $PackageReadme,
    (Join-Path $PluginRoot '.mcp.json'),
    (Join-Path $PluginRoot 'scripts'),
    (Join-Path $PluginRoot 'skills'),
    (Join-Path $PluginRoot 'parser\PaperParser.csproj'),
    (Join-Path $PluginRoot 'parser\Program.cs'),
    (Join-Path $PluginRoot 'parser\NuGet.Config')
)
foreach ($requiredPath in $requiredInputs) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "发布所需文件不存在：$requiredPath"
    }
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($manifest.name -ne $PluginName) {
    throw "预期插件名称为 '$PluginName'，实际为 '$($manifest.name)'。"
}
$safeVersion = ([string]$manifest.version) -replace '[^0-9A-Za-z.-]', '-'
if ([string]::IsNullOrWhiteSpace($safeVersion)) {
    throw '插件清单必须包含版本号。'
}

if (-not (Test-Path -LiteralPath $parserPath -PathType Leaf) -and -not $AllowMissingParser) {
    throw "随包 PDF 解析器不存在：$parserPath`n请先在发布机运行 .\plugins\$PluginName\scripts\build-paper-parser.ps1。"
}

$resolvedOutput = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $resolvedOutput | Out-Null
$packageDirectoryName = "$PluginName-marketplace-$safeVersion"
$packageRoot = Join-Path $resolvedOutput $packageDirectoryName
$archivePath = Join-Path $resolvedOutput "$packageDirectoryName.zip"
if (Test-Path -LiteralPath $packageRoot) {
    throw "发行目录已存在：$packageRoot。请使用新的输出目录，或确认后删除该旧构建。"
}
if (-not $SkipArchive -and (Test-Path -LiteralPath $archivePath)) {
    throw "发行 ZIP 已存在：$archivePath。请使用新的输出目录，或确认后删除该旧构建。"
}

New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot '.agents\plugins') | Out-Null
$stagedPluginRoot = Join-Path $packageRoot "plugins\$PluginName"
New-Item -ItemType Directory -Force -Path $stagedPluginRoot | Out-Null

# Marketplace 清单保持与源码完全一致，避免安装器和 CI 出现两套配置。
Copy-RequiredItem -Source $MarketplaceSource -Destination (Join-Path $packageRoot '.agents\plugins\marketplace.json')

# 仅复制 MCP 运行时，不复制 examples、tests、.git、构建中间产物或旧发行目录。
foreach ($runtimeItem in @('.codex-plugin', '.mcp.json', 'scripts', 'skills')) {
    Copy-RequiredItem -Source (Join-Path $PluginRoot $runtimeItem) -Destination $stagedPluginRoot
}

$stagedAssetsRoot = Join-Path $stagedPluginRoot 'assets'
New-Item -ItemType Directory -Force -Path $stagedAssetsRoot | Out-Null
foreach ($assetFile in @('paper-parser.exe', 'README.md')) {
    $assetPath = Join-Path $PluginRoot "assets\$assetFile"
    if (Test-Path -LiteralPath $assetPath -PathType Leaf) {
        Copy-RequiredItem -Source $assetPath -Destination $stagedAssetsRoot
    }
}

# 保留可审阅的解析器源码，以便发布者和安全团队重建 EXE；排除 bin/ 和 obj/。
$stagedParserRoot = Join-Path $stagedPluginRoot 'parser'
New-Item -ItemType Directory -Force -Path $stagedParserRoot | Out-Null
foreach ($parserSourceFile in @('PaperParser.csproj', 'Program.cs', 'NuGet.Config')) {
    Copy-RequiredItem -Source (Join-Path $PluginRoot "parser\$parserSourceFile") -Destination $stagedParserRoot
}

# 合规文件在仓库根维护，随包放到实际插件目录中。
foreach ($document in @('README.md', 'LICENSE', 'NOTICE.md', 'PRIVACY.md', 'SECURITY.md', 'CHANGELOG.md', 'THIRD_PARTY_NOTICES.md')) {
    Copy-RequiredItem -Source (Join-Path $RepositoryRoot $document) -Destination $stagedPluginRoot
}

Remove-StagedBuildCache -StagedPluginRoot $stagedPluginRoot
foreach ($forbiddenRelativePath in @('parser\bin', 'parser\obj', 'release', 'examples', 'tests', '.git', 'assets\parser-publish')) {
    $forbiddenPath = Join-Path $stagedPluginRoot $forbiddenRelativePath
    if (Test-Path -LiteralPath $forbiddenPath) {
        throw "拒绝将开发产物放入发行包：$forbiddenPath"
    }
}

# 包内安装器和验证器必须与运行时放在同一根目录，供离线安装与 GitHub 引导安装器共同使用。
Copy-RequiredItem -Source $PackageInstaller -Destination (Join-Path $packageRoot 'install.ps1')
Copy-RequiredItem -Source $PackageTester -Destination (Join-Path $packageRoot 'verify-release.ps1')
Copy-RequiredItem -Source $PackageReadme -Destination (Join-Path $packageRoot 'README.md')

if ($CodeSigningCertificateThumbprint) {
    Set-ReleaseSignature -Root $packageRoot -Thumbprint $CodeSigningCertificateThumbprint
}

if (-not $AllowMissingParser) {
    Write-ChecksumManifest -Root $packageRoot
    & (Join-Path $packageRoot 'verify-release.ps1') -MarketplaceRoot $packageRoot
} else {
    Write-Warning '因指定 -AllowMissingParser，已生成不完整测试包；请勿分发或安装。'
}

if (-not $SkipArchive) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($packageRoot, $archivePath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
}

[pscustomobject]@{
    package_root = $packageRoot
    archive_path = if ($SkipArchive) { $null } else { $archivePath }
    marketplace = $MarketplaceName
    plugin = $PluginName
    version = $manifest.version
    parser_included = Test-Path -LiteralPath (Join-Path $stagedPluginRoot 'assets\paper-parser.exe')
    node_required = $false
} | ConvertTo-Json -Depth 4
