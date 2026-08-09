<#
  对已经构建好的本地 Marketplace 发行目录做离线验证。

  本脚本不会启动 PowerPoint、不会联网，也不会注册 Codex，因此可直接在 CI 和发布机使用。
  它验证 Marketplace 布局、MCP 入口、随包解析器、脚本语法、禁止的开发产物和 SHA-256 清单。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$MarketplaceRoot,

    # 仅供构建中尚未生成 SHA256SUMS.txt 的临时检查；正式发行不得使用。
    [switch]$SkipChecksumCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'

function Get-RequiredProperty {
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Context
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        throw "$Context 缺少必填字段 '$Name'。"
    }
    return $property.Value
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

function Test-ChecksumManifest {
    param([Parameter(Mandatory)][string]$Root)

    $checksumPath = Join-Path $Root 'SHA256SUMS.txt'
    if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
        throw "未找到 SHA-256 清单：$checksumPath"
    }
    $lines = @(Get-Content -LiteralPath $checksumPath -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) {
        throw 'SHA-256 清单为空。'
    }

    foreach ($line in $lines) {
        $parts = $line -split '\s+\*?', 2
        if ($parts.Count -ne 2 -or $parts[0] -notmatch '^[A-Fa-f0-9]{64}$') {
            throw "SHA-256 清单格式错误：$line"
        }
        $relativePath = $parts[1].Trim()
        if ([System.IO.Path]::IsPathRooted($relativePath) -or $relativePath -match '(^|[\\/])\.\.([\\/]|$)') {
            throw "SHA-256 清单含不安全路径：$relativePath"
        }
        $filePath = Join-Path $Root ($relativePath -replace '/', '\\')
        if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
            throw "清单所列文件不存在：$relativePath"
        }
        $actual = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash
        if ($actual -ne $parts[0].ToUpperInvariant()) {
            throw "SHA-256 不匹配：$relativePath"
        }
    }
}

function Test-PowerShellSyntax {
    param([Parameter(Mandatory)][string]$Root)

    $syntaxErrors = @()
    $utf8WithoutBom = [System.Text.UTF8Encoding]::new($false)
    foreach ($scriptFile in Get-ChildItem -LiteralPath $Root -Recurse -Filter '*.ps1' -File) {
        $sourceText = [System.IO.File]::ReadAllText($scriptFile.FullName, $utf8WithoutBom)
        $tokens = $null
        $parseErrors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($sourceText, [ref]$tokens, [ref]$parseErrors)
        foreach ($parseError in $parseErrors) {
            $syntaxErrors += "{0}:{1}:{2} {3}" -f $scriptFile.FullName, $parseError.Extent.StartLineNumber, $parseError.Extent.StartColumnNumber, $parseError.Message
        }
    }
    if ($syntaxErrors.Count -gt 0) {
        throw "PowerShell AST 校验失败：`n$($syntaxErrors -join [Environment]::NewLine)"
    }
}

$root = (Resolve-Path -LiteralPath $MarketplaceRoot -ErrorAction Stop).Path
$marketplacePath = Join-Path $root '.agents\plugins\marketplace.json'
if (-not (Test-Path -LiteralPath $marketplacePath -PathType Leaf)) {
    throw "未找到 Marketplace 清单：$marketplacePath"
}

$marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
if ((Get-RequiredProperty $marketplace 'name' 'Marketplace 清单') -ne $MarketplaceName) {
    throw "Marketplace 名称必须为 '$MarketplaceName'。"
}
$interface = Get-RequiredProperty $marketplace 'interface' 'Marketplace 清单'
if ([string]::IsNullOrWhiteSpace([string](Get-RequiredProperty $interface 'displayName' 'Marketplace interface'))) {
    throw 'Marketplace interface.displayName 不能为空。'
}

$entry = @($marketplace.plugins | Where-Object { $_.name -eq $PluginName })
if ($entry.Count -ne 1) {
    throw "Marketplace 必须恰好包含一个 '$PluginName' 条目。"
}
$entry = $entry[0]
$source = Get-RequiredProperty $entry 'source' 'Marketplace 插件条目'
$policy = Get-RequiredProperty $entry 'policy' 'Marketplace 插件条目'
if ((Get-RequiredProperty $source 'source' 'Marketplace source') -ne 'local') {
    throw 'Marketplace source.source 必须为 local。'
}
if ((Get-RequiredProperty $source 'path' 'Marketplace source') -ne "./plugins/$PluginName") {
    throw "Marketplace source.path 必须为 './plugins/$PluginName'。"
}
if ((Get-RequiredProperty $policy 'installation' 'Marketplace policy') -notin @('NOT_AVAILABLE', 'AVAILABLE', 'INSTALLED_BY_DEFAULT')) {
    throw 'Marketplace policy.installation 无效。'
}
if ((Get-RequiredProperty $policy 'authentication' 'Marketplace policy') -notin @('ON_INSTALL', 'ON_USE')) {
    throw 'Marketplace policy.authentication 无效。'
}
[void](Get-RequiredProperty $entry 'category' 'Marketplace 插件条目')

$pluginRoot = Join-Path $root "plugins\$PluginName"
$manifestPath = Join-Path $pluginRoot '.codex-plugin\plugin.json'
$mcpPath = Join-Path $pluginRoot '.mcp.json'
$serverPath = Join-Path $pluginRoot 'scripts\paper-to-journal-club-server.ps1'
$parserPath = Join-Path $pluginRoot 'assets\paper-parser.exe'
$requiredPaths = @(
    $pluginRoot,
    $manifestPath,
    $mcpPath,
    $serverPath,
    $parserPath,
    (Join-Path $pluginRoot 'skills'),
    (Join-Path $pluginRoot 'THIRD_PARTY_NOTICES.md'),
    (Join-Path $pluginRoot 'parser\NuGet.Config')
)
foreach ($requiredPath in $requiredPaths) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "发行包缺少所需文件：$requiredPath"
    }
}

foreach ($forbiddenRelativePath in @('parser\bin', 'parser\obj', 'release', 'examples', 'tests', '.git', 'assets\parser-publish')) {
    $forbiddenPath = Join-Path $pluginRoot $forbiddenRelativePath
    if (Test-Path -LiteralPath $forbiddenPath) {
        throw "发行包包含禁止的开发产物：$forbiddenPath"
    }
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ((Get-RequiredProperty $manifest 'name' '插件清单') -ne $PluginName) {
    throw "插件清单名称必须为 '$PluginName'。"
}
$version = [string](Get-RequiredProperty $manifest 'version' '插件清单')
if ($version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$') {
    throw "插件版本不是严格 SemVer：$version"
}
[void](Get-RequiredProperty $manifest 'description' '插件清单')
$author = Get-RequiredProperty $manifest 'author' '插件清单'
[void](Get-RequiredProperty $author 'name' '插件作者')
$pluginInterface = Get-RequiredProperty $manifest 'interface' '插件清单'
foreach ($propertyName in @('displayName', 'shortDescription', 'longDescription', 'developerName', 'category', 'capabilities')) {
    [void](Get-RequiredProperty $pluginInterface $propertyName '插件 interface')
}

$mcp = Get-Content -LiteralPath $mcpPath -Raw -Encoding UTF8 | ConvertFrom-Json
$server = $mcp.mcpServers.$PluginName
if ($null -eq $server) {
    throw "未找到 MCP server '$PluginName'。"
}
if ($server.command -ne 'powershell.exe') {
    throw 'MCP server 必须由 powershell.exe 启动。'
}
$mcpArguments = @($server.args) -join ' '
if ($mcpArguments -match '(?i)(^|[\\/\s])node(\.exe)?([\\/\s]|$)') {
    throw 'MCP server 不得依赖 Node.js。'
}
if ($mcpArguments -match '(?i)(^|[\\/\s])bypass([\\/\s]|$)') {
    throw 'MCP server 不得使用 ExecutionPolicy Bypass。'
}
if (-not (Test-PortableExecutable -Path $parserPath)) {
    throw '随包 paper-parser.exe 缺失或不是有效的 Windows 可执行文件。'
}

Test-PowerShellSyntax -Root $root
if (-not $SkipChecksumCheck) {
    Test-ChecksumManifest -Root $root
}

[pscustomobject]@{
    pass = $true
    marketplace = $MarketplaceName
    plugin = $PluginName
    version = $version
    parser = $parserPath
    node_required = $false
} | ConvertTo-Json -Depth 4
