<#
  已解压发行包的离线安装器。

  仅接受 Build-ReleasePackage.ps1 生成的完整 Marketplace 根目录：
  <发行根>/.agents/plugins/marketplace.json
  <发行根>/plugins/paper-to-journal-club/

  普通用户不需要 Node.js、npm、Python 或 .NET SDK。PDF 解析器由随包的
  paper-parser.exe 提供，演示稿由本机 Microsoft PowerPoint COM 创建。
#>
[CmdletBinding()]
param(
    # 留空时把本脚本所在目录视为已解压发行包根目录。
    [string]$MarketplaceRoot,

    # 仅用于 CI、发布前验证或管理员预部署；普通用户应保留 PowerPoint 检查。
    [switch]$SkipPowerPointCheck,

    # 仅检查完整性，不登记 Marketplace 或安装到 Codex。
    [switch]$SkipCodexInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'

if ([string]::IsNullOrWhiteSpace($MarketplaceRoot)) {
    $MarketplaceRoot = $PSScriptRoot
}

function Assert-NativeSuccess {
    param([Parameter(Mandatory)][string]$Action)

    if ($LASTEXITCODE -ne 0) {
        throw "$Action 失败，退出码为 $LASTEXITCODE。"
    }
}

function Test-PortableExecutable {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $false
    }
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        if ($stream.Length -lt 2) {
            return $false
        }
        return $stream.ReadByte() -eq 0x4D -and $stream.ReadByte() -eq 0x5A
    } finally {
        $stream.Dispose()
    }
}

function Test-Checksums {
    param([Parameter(Mandatory)][string]$Root)

    $checksumPath = Join-Path $Root 'SHA256SUMS.txt'
    if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
        throw "发行包缺少 SHA256SUMS.txt：$checksumPath"
    }

    $invalid = @()
    foreach ($line in Get-Content -LiteralPath $checksumPath -Encoding UTF8) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $parts = $line -split '\s+\*?', 2
        if ($parts.Count -ne 2 -or $parts[0] -notmatch '^[A-Fa-f0-9]{64}$') {
            $invalid += "SHA-256 清单格式错误：$line"
            continue
        }
        $relativePath = $parts[1].Trim()
        if ([System.IO.Path]::IsPathRooted($relativePath) -or $relativePath -match '(^|[\\/])\.\.([\\/]|$)') {
            $invalid += "SHA-256 清单含不安全路径：$relativePath"
            continue
        }
        $targetPath = Join-Path $Root ($relativePath -replace '/', '\\')
        if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
            $invalid += "清单文件不存在：$relativePath"
            continue
        }
        $actual = (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash
        if ($actual -ne $parts[0].ToUpperInvariant()) {
            $invalid += "SHA-256 不匹配：$relativePath"
        }
    }
    if ($invalid.Count -gt 0) {
        throw "发行包完整性校验失败：`n$($invalid -join [Environment]::NewLine)"
    }
}

if ($env:OS -ne 'Windows_NT') {
    throw 'Paper to Journal Club 目前仅支持 Windows，因为它通过 PowerPoint COM 控制桌面版 Microsoft PowerPoint。'
}
if (-not [Environment]::Is64BitOperatingSystem) {
    throw '随包 PDF 解析器需要 64 位 Windows。'
}

$resolvedRoot = (Resolve-Path -LiteralPath $MarketplaceRoot -ErrorAction Stop).Path
$marketplaceFile = Join-Path $resolvedRoot '.agents\plugins\marketplace.json'
$pluginRoot = Join-Path $resolvedRoot "plugins\$PluginName"
$parserPath = Join-Path $pluginRoot 'assets\paper-parser.exe'
$mcpPath = Join-Path $pluginRoot '.mcp.json'
$verifierPath = Join-Path $resolvedRoot 'verify-release.ps1'

foreach ($requiredPath in @($marketplaceFile, $pluginRoot, $parserPath, $mcpPath, $verifierPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "发行包缺少所需文件：$requiredPath。请取得完整的正式 Release ZIP。"
    }
}
if (-not (Test-PortableExecutable -Path $parserPath)) {
    throw "随包 PDF 解析器不存在或无效：$parserPath。不要以本机 Node.js、Python、.NET SDK 或 Word 回退代替它。"
}

$mcp = Get-Content -LiteralPath $mcpPath -Raw -Encoding UTF8 | ConvertFrom-Json
$mcpServer = $mcp.mcpServers.$PluginName
if ($null -eq $mcpServer -or $mcpServer.command -ne 'powershell.exe') {
    throw '发行包没有使用预期的 PowerShell MCP 入口。'
}
$mcpArguments = @($mcpServer.args) -join ' '
if ($mcpArguments -match '(?i)(^|[\\/\s])node(\.exe)?([\\/\s]|$)') {
    throw '发行包意外依赖 Node.js，拒绝安装。'
}
if ($mcpArguments -match '(?i)(^|[\\/\s])bypass([\\/\s]|$)') {
    throw '发行包不得使用 ExecutionPolicy Bypass。'
}

# 先校验哈希，再执行包内结构验证，避免两套安装规则发生漂移。
Test-Checksums -Root $resolvedRoot
& $verifierPath -MarketplaceRoot $resolvedRoot

if (-not $SkipPowerPointCheck) {
    $powerPointClass = 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID'
    if (-not (Test-Path -LiteralPath $powerPointClass)) {
        throw '未检测到 Microsoft PowerPoint 桌面版。请安装并启动一次 PowerPoint 后重新运行安装器。'
    }
}

Write-Host '发行包完整性和运行前置检查通过。'
Write-Host '终端用户不需要 Node.js、npm、Python 或 .NET SDK。'

if ($SkipCodexInstall) {
    Write-Host '按请求跳过 Codex Marketplace 注册。'
    return
}

if (-not (Get-Command codex -ErrorAction SilentlyContinue)) {
    throw '未找到 Codex CLI。请安装或更新 Codex 桌面版/CLI，然后重新运行安装器。'
}

codex plugin marketplace add $resolvedRoot
Assert-NativeSuccess '注册 Paper to Journal Club Marketplace'

codex plugin add "$PluginName@$MarketplaceName"
Assert-NativeSuccess '安装 Paper to Journal Club 插件'

Write-Host "已安装 $PluginName@$MarketplaceName"
Write-Host "请保留该发行目录：Codex 会将其用作本地 Marketplace 源 ($resolvedRoot)。"
Write-Host '请重启 Codex，并新建一个任务后首次使用插件。'
