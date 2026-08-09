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

function Assert-ExactPropertyNames {
    <# 离线安装器只接受唯一且固定的 Marketplace/MCP 描述，避免注册额外入口。 #>
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string[]]$AllowedNames,
        [Parameter(Mandatory)][string]$Context
    )

    $actualNames = @($Object.PSObject.Properties | ForEach-Object { $_.Name })
    $unexpectedNames = @($actualNames | Where-Object { $AllowedNames -notcontains $_ })
    $missingNames = @($AllowedNames | Where-Object { $actualNames -notcontains $_ })
    if ($unexpectedNames.Count -gt 0 -or $missingNames.Count -gt 0) {
        throw "$Context 不符合允许字段清单。"
    }
}

function Assert-AllowedMarketplaceAndMcpConfiguration {
    <#
      在调用包内验证器或 Codex CLI 前执行纯数据校验。此处不允许任意 Marketplace 条目、
      远程 source、第二个 MCP server 或多余 PowerShell 参数。
    #>
    param([Parameter(Mandatory)][string]$Root)

    $marketplacePath = Join-Path $Root '.agents\plugins\marketplace.json'
    $mcpPath = Join-Path $Root "plugins\$PluginName\.mcp.json"
    $pluginManifestPath = Join-Path $Root "plugins\$PluginName\.codex-plugin\plugin.json"
    try {
        $marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $mcp = Get-Content -LiteralPath $mcpPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $pluginManifest = Get-Content -LiteralPath $pluginManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "无法解析 Marketplace 或 MCP JSON：$($_.Exception.Message)"
    }

    Assert-ExactPropertyNames -Object $marketplace -AllowedNames @('name', 'interface', 'plugins') -Context 'Marketplace 清单'
    if ($marketplace.name -ne $MarketplaceName) {
        throw "Marketplace 名称必须为 '$MarketplaceName'。"
    }
    Assert-ExactPropertyNames -Object $marketplace.interface -AllowedNames @('displayName') -Context 'Marketplace interface'
    if ([string]::IsNullOrWhiteSpace([string]$marketplace.interface.displayName)) {
        throw 'Marketplace interface.displayName 不能为空。'
    }
    $entries = @($marketplace.plugins)
    if ($entries.Count -ne 1) {
        throw 'Marketplace 只能包含一个插件条目。'
    }
    $entry = $entries[0]
    Assert-ExactPropertyNames -Object $entry -AllowedNames @('name', 'source', 'policy', 'category') -Context 'Marketplace 插件条目'
    Assert-ExactPropertyNames -Object $entry.source -AllowedNames @('source', 'path') -Context 'Marketplace source'
    Assert-ExactPropertyNames -Object $entry.policy -AllowedNames @('installation', 'authentication') -Context 'Marketplace policy'
    if ($entry.name -ne $PluginName -or $entry.source.source -ne 'local' -or $entry.source.path -ne "./plugins/$PluginName" -or $entry.policy.installation -ne 'AVAILABLE' -or $entry.policy.authentication -ne 'ON_INSTALL' -or [string]::IsNullOrWhiteSpace([string]$entry.category)) {
        throw 'Marketplace 条目不在发行包允许清单内。'
    }

    Assert-ExactPropertyNames -Object $mcp -AllowedNames @('mcpServers') -Context 'MCP 配置'
    Assert-ExactPropertyNames -Object $mcp.mcpServers -AllowedNames @($PluginName) -Context 'MCP server 清单'
    $server = $mcp.mcpServers.$PluginName
    Assert-ExactPropertyNames -Object $server -AllowedNames @('command', 'args', 'cwd') -Context 'MCP server'
    $expectedArguments = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', './scripts/paper-to-journal-club-server.ps1')
    $actualArguments = @($server.args | ForEach-Object { [string]$_ })
    if ($server.command -ne 'powershell.exe' -or $server.cwd -ne '.' -or $actualArguments.Count -ne $expectedArguments.Count) {
        throw 'MCP server 不在允许清单内。'
    }
    for ($index = 0; $index -lt $expectedArguments.Count; $index++) {
        if ($actualArguments[$index] -cne $expectedArguments[$index]) {
            throw 'MCP server 启动参数不在允许清单内。'
        }
    }

    # plugin.json 同样可能声明新的技能或 MCP 配置，因此也固定其字段和本地路径。
    Assert-ExactPropertyNames -Object $pluginManifest -AllowedNames @('name', 'version', 'description', 'author', 'license', 'keywords', 'skills', 'mcpServers', 'interface') -Context '插件清单'
    Assert-ExactPropertyNames -Object $pluginManifest.author -AllowedNames @('name') -Context '插件作者'
    Assert-ExactPropertyNames -Object $pluginManifest.interface -AllowedNames @('displayName', 'shortDescription', 'longDescription', 'developerName', 'category', 'capabilities', 'defaultPrompt') -Context '插件 interface'
    if ($pluginManifest.name -ne $PluginName -or [string]::IsNullOrWhiteSpace([string]$pluginManifest.version) -or [string]::IsNullOrWhiteSpace([string]$pluginManifest.description) -or [string]::IsNullOrWhiteSpace([string]$pluginManifest.author.name) -or $pluginManifest.license -ne 'MIT' -or $pluginManifest.skills -ne './skills/' -or $pluginManifest.mcpServers -ne './.mcp.json' -or @($pluginManifest.keywords).Count -eq 0 -or @($pluginManifest.interface.capabilities).Count -eq 0 -or @($pluginManifest.interface.defaultPrompt).Count -eq 0) {
        throw '插件清单不在允许的本地技能/MCP 配置范围内。'
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

function ConvertTo-CanonicalRelativePath {
    <# 将清单路径收敛到相对、无路径穿越的 ZIP 路径表示。 #>
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or $Path.IndexOf([char]0) -ge 0) {
        throw 'SHA-256 清单包含空路径或 NUL 字符。'
    }
    $normalized = $Path.Replace('\', '/').Trim()
    if ($normalized.StartsWith('/') -or $normalized.EndsWith('/') -or $normalized.Contains(':')) {
        throw "SHA-256 清单含不安全路径：$Path"
    }
    $segments = @($normalized -split '/')
    if ($segments.Count -eq 0 -or @($segments | Where-Object { $_ -in @('', '.', '..') }).Count -gt 0) {
        throw "SHA-256 清单含不安全路径：$Path"
    }
    return $normalized
}

function Get-RelativePath {
    <# 兼容 Windows PowerShell 5.1：其 .NET Framework 不提供 Path.GetRelativePath。 #>
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$TargetPath
    )

    $baseUri = [System.Uri]::new((Join-Path $BasePath ''))
    $targetUri = [System.Uri]::new($TargetPath)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString()).Replace('\', '/')
}

function Test-Checksums {
    param([Parameter(Mandatory)][string]$Root)

    $checksumPath = Join-Path $Root 'SHA256SUMS.txt'
    if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
        throw "发行包缺少 SHA256SUMS.txt：$checksumPath"
    }

    $manifestEntries = @{}
    foreach ($line in @(Get-Content -LiteralPath $checksumPath -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $match = [System.Text.RegularExpressions.Regex]::Match($line, '^(?<hash>[A-Fa-f0-9]{64})  \*(?<path>.+)$')
        if (-not $match.Success) {
            throw "SHA-256 清单格式错误：$line"
        }
        $relativePath = ConvertTo-CanonicalRelativePath -Path $match.Groups['path'].Value
        if ($manifestEntries.ContainsKey($relativePath)) {
            throw "SHA-256 清单包含重复路径：$relativePath"
        }
        $manifestEntries[$relativePath] = $match.Groups['hash'].Value.ToUpperInvariant()
    }
    if ($manifestEntries.Count -eq 0) {
        throw 'SHA-256 清单为空。'
    }

    $actualEntries = @{}
    foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
        if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "发行包不得包含重解析点文件：$($file.FullName)"
        }
        $relativePath = Get-RelativePath -BasePath $Root -TargetPath $file.FullName
        if ($relativePath -eq 'SHA256SUMS.txt') {
            continue
        }
        $relativePath = ConvertTo-CanonicalRelativePath -Path $relativePath
        if ($actualEntries.ContainsKey($relativePath)) {
            throw "发行包包含规范化后重复的文件路径：$relativePath"
        }
        $actualEntries[$relativePath] = $file.FullName
    }

    # 清单与文件列表都必须完全一致，防止攻击者在已解压目录塞入未校验的新脚本或配置。
    $missingEntries = @($manifestEntries.Keys | Where-Object { -not $actualEntries.ContainsKey($_) })
    $unexpectedEntries = @($actualEntries.Keys | Where-Object { -not $manifestEntries.ContainsKey($_) })
    if ($missingEntries.Count -gt 0 -or $unexpectedEntries.Count -gt 0) {
        $details = @()
        if ($missingEntries.Count -gt 0) {
            $details += "清单列出但实际缺失：$($missingEntries -join ', ')"
        }
        if ($unexpectedEntries.Count -gt 0) {
            $details += "实际存在但未列入清单：$($unexpectedEntries -join ', ')"
        }
        throw "发行包 SHA-256 清单双向集合校验失败：$($details -join '；')"
    }

    foreach ($relativePath in $manifestEntries.Keys) {
        $actual = (Get-FileHash -LiteralPath $actualEntries[$relativePath] -Algorithm SHA256).Hash.ToUpperInvariant()
        if ($actual -ne $manifestEntries[$relativePath]) {
            throw "SHA-256 不匹配：$relativePath"
        }
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

# 先双向校验所有文件的哈希和白名单配置，再执行已校验的包内结构验证器。
Test-Checksums -Root $resolvedRoot
Assert-AllowedMarketplaceAndMcpConfiguration -Root $resolvedRoot
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
