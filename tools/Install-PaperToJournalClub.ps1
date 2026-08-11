<#
  已解压发行包的离线安装器。

  仅接受 Build-ReleasePackage.ps1 生成的完整 Marketplace 根目录：
  <发行根>/.agents/plugins/marketplace.json
  <发行根>/plugins/paper-to-journal-club/

  普通用户不需要 Node.js、npm、Python 或 .NET SDK。PDF 解析器由随包的
  paper-parser.exe 提供，演示稿由本机 Microsoft PowerPoint COM 创建。

  通过包校验后，本脚本优先使用真实、可执行的 Codex CLI 完成无感安装。
  若 CLI 不存在、来自 WindowsApps 执行别名、权限被拒绝或不具备插件命令，
  则安全复制插件到当前用户的 ~/.codex/plugins/paper-to-journal-club，并更新
  默认个人 Marketplace ~/.agents/plugins/marketplace.json。回退路径不修改
  config.toml，也不会要求用户调整 WindowsApps 下的应用权限。
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # 留空时把本脚本所在目录视为已解压发行包根目录。
    [string]$MarketplaceRoot,

    # 仅用于 CI、发布前验证或管理员预部署；普通用户应保留 PowerPoint 检查。
    [switch]$SkipPowerPointCheck,

    # 仅检查完整性，不部署插件或登记当前用户的个人 Marketplace。
    # SkipCodexInstall 为旧版本自动化调用保留的兼容别名。
    [Alias('SkipCodexInstall')]
    [switch]$SkipPersonalMarketplaceDeployment
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'

if ([string]::IsNullOrWhiteSpace($MarketplaceRoot)) {
    $MarketplaceRoot = $PSScriptRoot
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
      在调用包内验证器或部署到个人 Marketplace 前执行纯数据校验。此处不允许任意 Marketplace 条目、
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

    # 不使用 Get-ChildItem -Recurse：它可能跟随目录联接点。安全遍历会逐项拒绝
    # junction、符号链接和其他重解析点，并同时计算实际文件的 SHA-256。
    $actualEntries = Get-DirectoryDigestMap -Root $Root -Purpose '发行包目录'
    if (-not $actualEntries.ContainsKey('SHA256SUMS.txt')) {
        throw '发行包目录缺少 SHA256SUMS.txt。'
    }
    [void]$actualEntries.Remove('SHA256SUMS.txt')

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
        if ($actualEntries[$relativePath] -ne $manifestEntries[$relativePath]) {
            throw "SHA-256 不匹配：$relativePath"
        }
    }

    # 返回经校验的清单，供复制阶段继续核对，缩短“校验后被替换”的攻击窗口。
    Write-Output -NoEnumerate $manifestEntries
}

function Get-ExistingPathAttributes {
    <# 不解析链接目标，直接读取目录项属性；不存在的路径以 Exists=false 返回。 #>
    param([Parameter(Mandatory)][string]$Path)

    try {
        return [pscustomobject]@{
            Exists = $true
            Attributes = [System.IO.File]::GetAttributes($Path)
        }
    } catch [System.IO.FileNotFoundException] {
        return [pscustomobject]@{ Exists = $false; Attributes = $null }
    } catch [System.IO.DirectoryNotFoundException] {
        return [pscustomobject]@{ Exists = $false; Attributes = $null }
    }
}

function Assert-NoReparsePointInExistingPath {
    <#
      逐级拒绝符号链接、junction 和其他重解析点。个人 Marketplace 位于用户目录，
      因此不能仅检查末级路径，避免部署或配置写入被重定向到目录外。
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose,
        [switch]$RequireExisting
    )

    try {
        $fullPath = [System.IO.Path]::GetFullPath($Path)
    } catch {
        throw "$Purpose 不是有效的文件系统路径：$($_.Exception.Message)"
    }
    $pathRoot = [System.IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($pathRoot)) {
        throw "$Purpose 缺少有效的路径根目录：$Path"
    }

    $rootAttributes = Get-ExistingPathAttributes -Path $pathRoot
    if (-not $rootAttributes.Exists) {
        throw "$Purpose 的路径根目录不存在或无法访问：$pathRoot"
    }
    if (($rootAttributes.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "$Purpose 的路径根目录是重解析点，已拒绝：$pathRoot"
    }

    $relativePart = $fullPath.Substring($pathRoot.Length)
    $segments = @($relativePart.Split([char[]]@([char]'\', [char]'/'), [System.StringSplitOptions]::RemoveEmptyEntries))
    $currentPath = $pathRoot
    for ($index = 0; $index -lt $segments.Count; $index++) {
        $currentPath = [System.IO.Path]::Combine($currentPath, $segments[$index])
        $attributes = Get-ExistingPathAttributes -Path $currentPath
        if (-not $attributes.Exists) {
            continue
        }
        if (($attributes.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "$Purpose 的路径包含重解析点，已拒绝：$currentPath"
        }
        if ($index -lt ($segments.Count - 1) -and -not [System.IO.Directory]::Exists($currentPath)) {
            throw "$Purpose 的中间路径不是目录：$currentPath"
        }
    }

    $finalAttributes = Get-ExistingPathAttributes -Path $fullPath
    if ($RequireExisting -and -not $finalAttributes.Exists) {
        throw "$Purpose 不存在：$fullPath"
    }
    return $fullPath
}

function Assert-NoReparsePointsInDirectoryTree {
    <# 用队列遍历，避免 PowerShell 的 -Recurse 跟随不可信 junction。 #>
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Purpose
    )

    $rootPath = Assert-NoReparsePointInExistingPath -Path $Root -Purpose $Purpose -RequireExisting
    if (-not [System.IO.Directory]::Exists($rootPath)) {
        throw "$Purpose 必须是目录：$rootPath"
    }

    $pendingDirectories = New-Object 'System.Collections.Generic.Queue[string]'
    $pendingDirectories.Enqueue($rootPath)
    while ($pendingDirectories.Count -gt 0) {
        $directoryPath = $pendingDirectories.Dequeue()
        $directoryPath = Assert-NoReparsePointInExistingPath -Path $directoryPath -Purpose $Purpose -RequireExisting
        foreach ($child in @(Get-ChildItem -LiteralPath $directoryPath -Force -ErrorAction Stop)) {
            $attributes = Get-ExistingPathAttributes -Path $child.FullName
            if (-not $attributes.Exists) {
                throw "$Purpose 在检查期间发生并发路径变化：$($child.FullName)"
            }
            if (($attributes.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "$Purpose 包含重解析点，已拒绝：$($child.FullName)"
            }
            if ([System.IO.Directory]::Exists($child.FullName)) {
                $pendingDirectories.Enqueue($child.FullName)
            }
        }
    }
    return $rootPath
}

function Ensure-SafeDirectory {
    <# 创建目录前后均验证路径链，降低用户目录中 junction 的写入风险。 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose
    )

    $fullPath = Assert-NoReparsePointInExistingPath -Path $Path -Purpose $Purpose
    $attributes = Get-ExistingPathAttributes -Path $fullPath
    if ($attributes.Exists -and -not [System.IO.Directory]::Exists($fullPath)) {
        throw "$Purpose 已存在但不是目录：$fullPath"
    }
    [System.IO.Directory]::CreateDirectory($fullPath) | Out-Null
    $fullPath = Assert-NoReparsePointInExistingPath -Path $fullPath -Purpose $Purpose -RequireExisting
    if (-not [System.IO.Directory]::Exists($fullPath)) {
        throw "$Purpose 创建后不是目录：$fullPath"
    }
    return $fullPath
}

function New-UniqueChildPath {
    <# 返回未创建的随机子路径；父目录必须是已经验证的实体目录。 #>
    param(
        [Parameter(Mandatory)][string]$Parent,
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Purpose
    )

    $parentPath = Assert-NoReparsePointInExistingPath -Path $Parent -Purpose "$Purpose 的父目录" -RequireExisting
    if (-not [System.IO.Directory]::Exists($parentPath)) {
        throw "$Purpose 的父路径不是目录：$parentPath"
    }
    for ($attempt = 0; $attempt -lt 16; $attempt++) {
        $candidate = Join-Path $parentPath ("{0}{1}" -f $Prefix, [Guid]::NewGuid().ToString('N'))
        $candidate = Assert-NoReparsePointInExistingPath -Path $candidate -Purpose $Purpose
        if (-not (Get-ExistingPathAttributes -Path $candidate).Exists) {
            return $candidate
        }
    }
    throw "$Purpose 无法生成未占用的随机路径。"
}

function Get-DirectoryDigestMap {
    <# 对复制前后的插件树建立相对路径 → SHA-256 映射，确保本地副本没有漏项或损坏。 #>
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Purpose
    )

    $rootPath = Assert-NoReparsePointsInDirectoryTree -Root $Root -Purpose $Purpose
    $files = @{}
    $pendingDirectories = New-Object 'System.Collections.Generic.Queue[string]'
    $pendingDirectories.Enqueue($rootPath)
    while ($pendingDirectories.Count -gt 0) {
        $directoryPath = $pendingDirectories.Dequeue()
        foreach ($child in @(Get-ChildItem -LiteralPath $directoryPath -Force -ErrorAction Stop)) {
            $attributes = Get-ExistingPathAttributes -Path $child.FullName
            if (-not $attributes.Exists -or ($attributes.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "$Purpose 在枚举时遇到不安全目录项：$($child.FullName)"
            }
            if ([System.IO.Directory]::Exists($child.FullName)) {
                $pendingDirectories.Enqueue($child.FullName)
                continue
            }
            if (-not [System.IO.File]::Exists($child.FullName)) {
                throw "$Purpose 包含非普通文件或目录的目录项：$($child.FullName)"
            }
            $relativePath = ConvertTo-CanonicalRelativePath -Path (Get-RelativePath -BasePath $rootPath -TargetPath $child.FullName)
            if ($files.ContainsKey($relativePath)) {
                throw "$Purpose 包含规范化后重复的文件路径：$relativePath"
            }
            $files[$relativePath] = (Get-FileHash -LiteralPath $child.FullName -Algorithm SHA256).Hash.ToUpperInvariant()
        }
    }
    return $files
}

function Assert-DirectoryCopyMatches {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        # 来自 SHA256SUMS.txt 的插件子树清单；存在时，源目录本身也必须仍与它一致。
        [hashtable]$ExpectedSourceHashes
    )

    $sourceMap = Get-DirectoryDigestMap -Root $Source -Purpose '发行插件源目录'
    if ($null -ne $ExpectedSourceHashes) {
        $sourceMissing = @($ExpectedSourceHashes.Keys | Where-Object { -not $sourceMap.ContainsKey($_) })
        $sourceUnexpected = @($sourceMap.Keys | Where-Object { -not $ExpectedSourceHashes.ContainsKey($_) })
        if ($sourceMissing.Count -gt 0 -or $sourceUnexpected.Count -gt 0) {
            throw '待部署发行插件与经校验的 SHA-256 清单文件集合不一致。'
        }
        foreach ($relativePath in $ExpectedSourceHashes.Keys) {
            if ($sourceMap[$relativePath] -ne $ExpectedSourceHashes[$relativePath]) {
                throw "待部署发行插件 SHA-256 不匹配：$relativePath"
            }
        }
    }
    $destinationMap = Get-DirectoryDigestMap -Root $Destination -Purpose '个人插件副本目录'
    $missing = @($sourceMap.Keys | Where-Object { -not $destinationMap.ContainsKey($_) })
    $unexpected = @($destinationMap.Keys | Where-Object { -not $sourceMap.ContainsKey($_) })
    if ($missing.Count -gt 0 -or $unexpected.Count -gt 0) {
        throw '个人插件副本与已验证发行插件的文件集合不一致。'
    }
    foreach ($relativePath in $sourceMap.Keys) {
        if ($sourceMap[$relativePath] -ne $destinationMap[$relativePath]) {
            throw "个人插件副本 SHA-256 不匹配：$relativePath"
        }
    }
}

function Get-PluginChecksumMap {
    <# 从完整发行清单中提取插件目录的相对路径 → SHA-256 映射。 #>
    param([Parameter(Mandatory)][hashtable]$PackageChecksums)

    $prefix = "plugins/$PluginName/"
    $pluginChecksums = @{}
    foreach ($path in $PackageChecksums.Keys) {
        $pathText = [string]$path
        if (-not $pathText.StartsWith($prefix, [System.StringComparison]::Ordinal)) {
            continue
        }
        $relativePluginPath = $pathText.Substring($prefix.Length)
        if ([string]::IsNullOrWhiteSpace($relativePluginPath) -or $pluginChecksums.ContainsKey($relativePluginPath)) {
            throw '发行包 SHA-256 清单包含无效或重复的插件路径。'
        }
        $pluginChecksums[$relativePluginPath] = [string]$PackageChecksums[$path]
    }
    if ($pluginChecksums.Count -eq 0) {
        throw '发行包 SHA-256 清单未包含插件文件。'
    }
    Write-Output -NoEnumerate $pluginChecksums
}

function Assert-ManagedPersonalPluginRoot {
    <# 只覆盖同名且确实声明为本插件的旧目录，避免替换用户的无关文件。 #>
    param([Parameter(Mandatory)][string]$Path)

    $root = Assert-NoReparsePointsInDirectoryTree -Root $Path -Purpose '现有个人插件目录'
    $manifestPath = Join-Path $root '.codex-plugin\plugin.json'
    $manifestPath = Assert-NoReparsePointInExistingPath -Path $manifestPath -Purpose '现有个人插件清单' -RequireExisting
    if (-not [System.IO.File]::Exists($manifestPath)) {
        throw "拒绝替换缺少插件清单的目录：$root"
    }
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "拒绝替换无法解析的个人插件清单：$($_.Exception.Message)"
    }
    if ([string]$manifest.name -ne $PluginName) {
        throw "拒绝替换不属于 $PluginName 的目录：$root"
    }
    return $root
}

function Remove-SafeDirectory {
    <# 删除随机候选或已验证的旧副本前，先确认整棵树不包含重解析点。 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose
    )

    if (-not (Get-ExistingPathAttributes -Path $Path).Exists) {
        return
    }
    $safePath = Assert-NoReparsePointsInDirectoryTree -Root $Path -Purpose $Purpose
    Remove-Item -LiteralPath $safePath -Recurse -Force -ErrorAction Stop
}

function Get-CurrentUserProfilePath {
    $userProfile = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
    if ([string]::IsNullOrWhiteSpace($userProfile)) {
        $userProfile = $env:USERPROFILE
    }
    if ([string]::IsNullOrWhiteSpace($userProfile) -or -not [System.IO.Path]::IsPathRooted($userProfile)) {
        throw '无法确定当前 Windows 用户目录，无法部署个人 Marketplace。'
    }
    return (Assert-NoReparsePointInExistingPath -Path $userProfile -Purpose '当前 Windows 用户目录' -RequireExisting)
}

function Restore-InterruptedPersonalPluginDeployment {
    <#
      目录移动不是 File.Replace 那样的单文件原子替换。若上次安装在 old → previous 与
      candidate → target 之间异常终止，这里仅在目标缺失、且恰好存在一个经验证的旧副本时恢复它。
      多个候选或未知目录一律停止，绝不猜测或删除用户数据。
    #>
    param([Parameter(Mandatory)][string]$PluginsDirectory)

    $pluginsDirectory = Assert-NoReparsePointInExistingPath -Path $PluginsDirectory -Purpose '个人插件目录' -RequireExisting
    $targetPath = Join-Path $pluginsDirectory $PluginName
    if ((Get-ExistingPathAttributes -Path $targetPath).Exists) {
        return $targetPath
    }

    $escapedName = [System.Text.RegularExpressions.Regex]::Escape($PluginName)
    $previousPattern = "^\.$escapedName\.previous-[0-9a-f]{32}$"
    $previousCandidates = @(
        Get-ChildItem -LiteralPath $pluginsDirectory -Force -Directory -ErrorAction Stop |
            Where-Object { $_.Name -match $previousPattern }
    )
    if ($previousCandidates.Count -eq 0) {
        return $targetPath
    }
    if ($previousCandidates.Count -ne 1) {
        throw '检测到多个未完成的个人插件旧副本；为避免覆盖用户数据，已停止。请保留这些目录并联系维护者。'
    }

    $previousPath = Assert-ManagedPersonalPluginRoot -Path $previousCandidates[0].FullName
    if ((Get-ExistingPathAttributes -Path $targetPath).Exists) {
        throw '恢复个人插件旧副本时目标目录已出现，已停止以避免覆盖。'
    }
    [System.IO.Directory]::Move($previousPath, $targetPath)
    return (Assert-ManagedPersonalPluginRoot -Path $targetPath)
}

function Deploy-VerifiedPluginCopy {
    <#
      先复制到随机候选目录、逐文件比对哈希，再用目录移动替换同名旧副本。
      若复制或替换失败，保留旧副本或候选用于排障，绝不覆盖未知目录。
    #>
    param(
        [Parameter(Mandatory)][string]$SourcePluginRoot,
        [Parameter(Mandatory)][string]$UserProfilePath,
        [Parameter(Mandatory)][hashtable]$ExpectedPluginChecksums
    )

    $source = Assert-NoReparsePointsInDirectoryTree -Root $SourcePluginRoot -Purpose '待部署发行插件目录'
    $pluginsDirectory = Ensure-SafeDirectory -Path (Join-Path $UserProfilePath '.codex\plugins') -Purpose '个人插件目录'
    $target = Restore-InterruptedPersonalPluginDeployment -PluginsDirectory $pluginsDirectory
    $targetAttributes = Get-ExistingPathAttributes -Path $target
    if ($targetAttributes.Exists) {
        $target = Assert-ManagedPersonalPluginRoot -Path $target
    }

    $candidate = New-UniqueChildPath -Parent $pluginsDirectory -Prefix ".${PluginName}.incoming-" -Purpose '个人插件候选目录'
    $previous = $null
    $activated = $false
    try {
        Copy-Item -LiteralPath $source -Destination $candidate -Recurse -Force -ErrorAction Stop
        $candidate = Assert-NoReparsePointsInDirectoryTree -Root $candidate -Purpose '个人插件候选目录'
        Assert-DirectoryCopyMatches -Source $source -Destination $candidate -ExpectedSourceHashes $ExpectedPluginChecksums

        if ($targetAttributes.Exists) {
            $previous = New-UniqueChildPath -Parent $pluginsDirectory -Prefix ".${PluginName}.previous-" -Purpose '个人插件旧副本'
            [System.IO.Directory]::Move($target, $previous)
            $previous = Assert-ManagedPersonalPluginRoot -Path $previous
        }
        [System.IO.Directory]::Move($candidate, $target)
        $activated = $true
        $target = Assert-ManagedPersonalPluginRoot -Path $target
        Assert-DirectoryCopyMatches -Source $source -Destination $target -ExpectedSourceHashes $ExpectedPluginChecksums
    } catch {
        $deploymentError = $_
        if ($activated -and (Get-ExistingPathAttributes -Path $target).Exists) {
            try { Remove-SafeDirectory -Path $target -Purpose '失败的个人插件副本' } catch { Write-Warning "未删除失败的个人插件副本：$($_.Exception.Message)" }
        }
        if ($null -ne $previous -and (Get-ExistingPathAttributes -Path $previous).Exists -and -not (Get-ExistingPathAttributes -Path $target).Exists) {
            try { [System.IO.Directory]::Move($previous, $target) } catch { Write-Warning "未恢复此前的个人插件副本：$($_.Exception.Message)" }
        }
        if ((Get-ExistingPathAttributes -Path $candidate).Exists) {
            try { Remove-SafeDirectory -Path $candidate -Purpose '未激活的个人插件候选目录' } catch { Write-Warning "未删除个人插件候选目录：$($_.Exception.Message)" }
        }
        throw $deploymentError
    }

    if ($null -ne $previous -and (Get-ExistingPathAttributes -Path $previous).Exists) {
        # 旧副本已经验证为同名插件；成功后清理，避免累积多个不可见版本目录。
        Remove-SafeDirectory -Path $previous -Purpose '已替换的个人插件旧副本'
    }
    return $target
}

function New-PersonalMarketplaceEntry {
    return [pscustomobject]@{
        name = $PluginName
        source = [pscustomobject]@{
            source = 'local'
            path = "./.codex/plugins/$PluginName"
        }
        policy = [pscustomobject]@{
            installation = 'AVAILABLE'
            authentication = 'ON_INSTALL'
        }
        category = 'Productivity'
    }
}

function Assert-PersonalMarketplaceEntry {
    param([Parameter(Mandatory)]$Entry)

    if ($null -eq $Entry -or [string]$Entry.name -ne $PluginName -or $null -eq $Entry.source -or $null -eq $Entry.policy -or $Entry.source.source -ne 'local' -or $Entry.source.path -ne "./.codex/plugins/$PluginName" -or $Entry.policy.installation -ne 'AVAILABLE' -or $Entry.policy.authentication -ne 'ON_INSTALL' -or $Entry.category -ne 'Productivity') {
        throw '个人 Marketplace 中的 Paper to Journal Club 条目不符合预期。'
    }
}

function Update-PersonalMarketplaceFile {
    <#
      只替换本插件的同名条目，保留用户现有的其他 Marketplace 条目和元数据。
      配置文件使用同目录临时文件 + 备份移动，避免中途写出半截 JSON。
    #>
    param([Parameter(Mandatory)][string]$UserProfilePath)

    $agentsDirectory = Ensure-SafeDirectory -Path (Join-Path $UserProfilePath '.agents\plugins') -Purpose '个人 Marketplace 目录'
    $marketplacePath = Join-Path $agentsDirectory 'marketplace.json'
    $marketplaceAttributes = Get-ExistingPathAttributes -Path $marketplacePath
    $backupPath = $null

    if ($marketplaceAttributes.Exists) {
        $marketplacePath = Assert-NoReparsePointInExistingPath -Path $marketplacePath -Purpose '现有个人 Marketplace 文件' -RequireExisting
        if (-not [System.IO.File]::Exists($marketplacePath)) {
            throw "个人 Marketplace 路径不是普通文件：$marketplacePath"
        }
        try {
            $marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
            throw "无法解析现有个人 Marketplace JSON；为避免覆盖用户配置，已停止：$($_.Exception.Message)"
        }
        if ($null -eq $marketplace -or $marketplace -is [System.Array]) {
            throw '现有个人 Marketplace JSON 必须是对象；为避免覆盖用户配置，已停止。'
        }
    } else {
        $marketplace = [pscustomobject]@{
            name = 'personal'
            interface = [pscustomobject]@{ displayName = '个人插件' }
            plugins = @()
        }
    }

    if ($null -eq $marketplace.PSObject.Properties['name']) {
        $marketplace | Add-Member -NotePropertyName 'name' -NotePropertyValue 'personal'
    } elseif ([string]::IsNullOrWhiteSpace([string]$marketplace.name)) {
        throw '现有个人 Marketplace 缺少有效 name；为避免覆盖用户配置，已停止。'
    }
    if ($null -eq $marketplace.PSObject.Properties['interface']) {
        $marketplace | Add-Member -NotePropertyName 'interface' -NotePropertyValue ([pscustomobject]@{ displayName = '个人插件' })
    } elseif ($null -eq $marketplace.interface -or $marketplace.interface -is [string] -or $marketplace.interface -is [System.Array]) {
        throw '现有个人 Marketplace 的 interface 不是对象；为避免覆盖用户配置，已停止。'
    } elseif ($null -eq $marketplace.interface.PSObject.Properties['displayName']) {
        $marketplace.interface | Add-Member -NotePropertyName 'displayName' -NotePropertyValue '个人插件'
    }
    if ($null -eq $marketplace.PSObject.Properties['plugins']) {
        $marketplace | Add-Member -NotePropertyName 'plugins' -NotePropertyValue @()
    } elseif ($null -eq $marketplace.plugins -or $marketplace.plugins -isnot [System.Array]) {
        # ConvertFrom-Json 对 JSON 数组会产生 Object[]；对象、字符串或 null 都不能被
        # 安全地当作插件列表重写，以免损坏用户手工维护的配置。
        throw '现有个人 Marketplace 的 plugins 必须是 JSON 数组；为避免覆盖用户配置，已停止。'
    }

    $mergedEntries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in @($marketplace.plugins)) {
        if ($null -eq $entry -or $entry -is [string] -or $entry -is [System.ValueType]) {
            throw '现有个人 Marketplace 包含无法安全保留的插件条目；已停止写入。'
        }
        if ([string]$entry.name -ne $PluginName) {
            [void]$mergedEntries.Add($entry)
        }
    }
    [void]$mergedEntries.Add((New-PersonalMarketplaceEntry))
    # Windows PowerShell 5.1 对 Generic List<T>.ToArray() 的绑定在部分宿主中会失败。
    # 手动构造 object[]，确保单个条目也会被序列化为 JSON 数组而非对象。
    # New-Object 'object[]' <count> 在 Windows PowerShell 5.1 会生成带 value/Count
    # 属性的包装对象，而不是预期的数组。用 Array.CreateInstance 明确取得 object[]。
    $mergedEntryArray = [System.Array]::CreateInstance([object], [int]$mergedEntries.Count)
    for ($entryIndex = 0; $entryIndex -lt $mergedEntries.Count; $entryIndex++) {
        $mergedEntryArray[$entryIndex] = $mergedEntries[$entryIndex]
    }
    $marketplace.plugins = $mergedEntryArray

    $temporaryPath = New-UniqueChildPath -Parent $agentsDirectory -Prefix '.marketplace.json.incoming-' -Purpose '个人 Marketplace 临时文件'
    try {
        # 不允许 ConvertTo-Json 静默截断过深的用户元数据；一旦出现深度警告，保留原文件并停止。
        $serializationWarnings = @()
        $json = $marketplace | ConvertTo-Json -Depth 32 -WarningVariable serializationWarnings
        if (@($serializationWarnings).Count -gt 0) {
            throw '现有个人 Marketplace 的嵌套层级过深，无法在不截断内容的前提下安全更新。'
        }
        # CreateNew 防止随机临时名在检查后被其他进程抢占或覆盖。
        $payload = [System.Text.UTF8Encoding]::new($false).GetBytes($json)
        $temporaryStream = $null
        try {
            $temporaryStream = [System.IO.File]::Open($temporaryPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            $temporaryStream.Write($payload, 0, $payload.Length)
            $temporaryStream.Flush()
        } finally {
            if ($null -ne $temporaryStream) { $temporaryStream.Dispose() }
        }
        $temporaryPath = Assert-NoReparsePointInExistingPath -Path $temporaryPath -Purpose '个人 Marketplace 临时文件' -RequireExisting
        if (-not [System.IO.File]::Exists($temporaryPath)) {
            throw '个人 Marketplace 临时路径不是普通文件。'
        }
        $verification = Get-Content -LiteralPath $temporaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $targetEntry = @($verification.plugins | Where-Object { [string]$_.name -eq $PluginName })
        if ($targetEntry.Count -ne 1) {
            throw '个人 Marketplace 临时文件未包含唯一的目标插件条目。'
        }
        Assert-PersonalMarketplaceEntry -Entry $targetEntry[0]

        if ($marketplaceAttributes.Exists) {
            $backupPath = New-UniqueChildPath -Parent $agentsDirectory -Prefix '.marketplace.json.paper-to-journal-club-backup-' -Purpose '个人 Marketplace 备份文件'
            # 同卷 File.Replace 是原子替换：即使进程中断，也会保留原文件或完整新文件，
            # 同时将旧文件写入备份路径，避免两次 Move 之间暂时丢失 marketplace.json。
            [System.IO.File]::Replace($temporaryPath, $marketplacePath, $backupPath, $true)
        } else {
            [System.IO.File]::Move($temporaryPath, $marketplacePath)
        }
        $marketplacePath = Assert-NoReparsePointInExistingPath -Path $marketplacePath -Purpose '更新后的个人 Marketplace 文件' -RequireExisting
    } catch {
        $updateError = $_
        if ($null -ne $backupPath -and (Get-ExistingPathAttributes -Path $backupPath).Exists -and -not (Get-ExistingPathAttributes -Path $marketplacePath).Exists) {
            try { [System.IO.File]::Move($backupPath, $marketplacePath) } catch { Write-Warning "未恢复个人 Marketplace 备份：$($_.Exception.Message)" }
        }
        if ((Get-ExistingPathAttributes -Path $temporaryPath).Exists) {
            try { [System.IO.File]::Delete($temporaryPath) } catch { Write-Warning "未删除个人 Marketplace 临时文件：$($_.Exception.Message)" }
        }
        throw $updateError
    }

    return [pscustomobject]@{
        Path = $marketplacePath
        BackupPath = $backupPath
    }
}

function Test-TrustedCodexCliSignature {
    <#
      PATH 中的同名 exe 不能仅凭文件名信任。仅接受 Windows 已验证有效、且签名
      Subject 或 Organization 明确以 OpenAI 开头的 Authenticode 签名。签发者轮换
      时允许 OpenAI 的新证书，不把某一张证书 thumbprint 写死在安装器中。
    #>
    param([Parameter(Mandatory)][string]$Path)

    try {
        $signature = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or $null -eq $signature.SignerCertificate) {
            return $false
        }
        $subject = [string]$signature.SignerCertificate.Subject
        $simpleName = [string]$signature.SignerCertificate.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
        # 当前官方签名为 OpenAI OpCo, LLC；同时兼容未来合法的 OpenAI 实体名称。
        return ($simpleName -match '(?i)^OpenAI(?:\s|,|$)' -or $subject -match '(?i)(^|,\s*)(CN|O)\s*=\s*"?OpenAI(?:\s|,|$)')
    } catch {
        return $false
    }
}

function Get-ExecutableCodexCliPath {
    <#
      仅接受 PowerShell 识别为 Application 的实体文件，明确排除 Microsoft Store
      WindowsApps 执行别名。后者常可被 Get-Command 找到，却会在子进程启动时
      报 Access is denied；把它当成可用 CLI 只会让首次安装变得更复杂。
    #>
    $candidates = @(Get-Command -Name 'codex' -CommandType Application -All -ErrorAction SilentlyContinue)
    foreach ($candidate in $candidates) {
        if ($candidate.CommandType -ne [System.Management.Automation.CommandTypes]::Application) {
            continue
        }
        $candidatePath = [string]$candidate.Path
        if ([string]::IsNullOrWhiteSpace($candidatePath)) {
            $candidatePath = [string]$candidate.Source
        }
        if ([string]::IsNullOrWhiteSpace($candidatePath)) {
            continue
        }
        try {
            $candidatePath = Assert-NoReparsePointInExistingPath -Path $candidatePath -Purpose '候选 Codex CLI' -RequireExisting
        } catch {
            continue
        }

        # 同时覆盖 Program Files 和用户 AppData 下的 WindowsApps alias 目录。
        if ($candidatePath -match '(?i)(^|\\)WindowsApps(\\|$)') {
            continue
        }
        if ([System.IO.Path]::GetFileName($candidatePath) -ine 'codex.exe' -or -not [System.IO.File]::Exists($candidatePath)) {
            continue
        }
        try {
            $attributes = [System.IO.File]::GetAttributes($candidatePath)
            if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                continue
            }
        } catch {
            continue
        }
        if (-not (Test-TrustedCodexCliSignature -Path $candidatePath)) {
            continue
        }
        return $candidatePath
    }
    return $null
}

function ConvertTo-WindowsProcessArgument {
    <# 为 ProcessStartInfo.Arguments 生成 Windows CreateProcess 规则的单个安全参数。 #>
    param([Parameter(Mandatory)][string]$Argument)

    if ($Argument.Length -eq 0) {
        return '""'
    }
    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq [char]'\') {
            $backslashCount++
            continue
        }
        if ($character -eq [char]'"') {
            [void]$builder.Append([char]'\', ($backslashCount * 2) + 1)
            [void]$builder.Append('"')
            $backslashCount = 0
            continue
        }
        if ($backslashCount -gt 0) {
            [void]$builder.Append([char]'\', $backslashCount)
            $backslashCount = 0
        }
        [void]$builder.Append($character)
    }
    if ($backslashCount -gt 0) {
        [void]$builder.Append([char]'\', $backslashCount * 2)
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Invoke-CodexCliCommand {
    <#
      所有 CLI 调用都通过绝对路径启动，并分离 stdout/stderr：Marketplace list 的
      JSON 不能被 CLI 自身的临时目录告警污染。捕获 Access is denied、损坏的可执行
      文件和非零退出码，让调用方可以无中断地转向个人 Marketplace 回退路径。
      Codex CLI 自身可能创建临时数据，因此任何调用都只会发生在用户确认的安装事务内。
    #>
    param(
        [Parameter(Mandatory)][string]$ExecutablePath,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    try {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ExecutablePath
        $startInfo.Arguments = (@($Arguments | ForEach-Object { ConvertTo-WindowsProcessArgument -Argument $_ }) -join ' ')
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        if (-not $process.Start()) {
            throw '无法启动已验证的 Codex CLI。'
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(120000)) {
            try { $process.Kill() } catch { }
            throw 'Codex CLI 在 120 秒内未结束，已停止等待。'
        }
        $rawOutput = $stdoutTask.GetAwaiter().GetResult().Trim()
        $errorOutput = $stderrTask.GetAwaiter().GetResult().Trim()
        $summary = if ([string]::IsNullOrWhiteSpace($errorOutput)) { $rawOutput } else { $errorOutput }
        if ($summary.Length -gt 500) {
            $summary = $summary.Substring(0, 500) + '…'
        }
        return [pscustomobject]@{
            Succeeded = ($process.ExitCode -eq 0)
            ExitCode = [int]$process.ExitCode
            Summary = $summary
            Output = $rawOutput
            ErrorOutput = $errorOutput
        }
    } catch {
        return [pscustomobject]@{
            Succeeded = $false
            ExitCode = -1
            Summary = $_.Exception.Message
            Output = ''
            ErrorOutput = $_.Exception.Message
        }
    }
}

function Get-CodexMarketplaceInventory {
    <#
      事务开始后从 CLI 的 JSON 列表中取得 Marketplace 状态。无法获得可验证的
      JSON 时宁可放弃 CLI 自动安装并走个人 Marketplace，绝不猜测某条目是否是
      用户先前登记的条目，更不会据此删除它。
    #>
    param([Parameter(Mandatory)][string]$ExecutablePath)

    $result = Invoke-CodexCliCommand -ExecutablePath $ExecutablePath -Arguments ([string[]]@('plugin', 'marketplace', 'list', '--json'))
    if (-not $result.Succeeded) {
        return [pscustomobject]@{
            Succeeded = $false
            Existing = $false
            TargetEntry = $null
            Summary = $result.Summary
        }
    }

    try {
        $parsed = $result.Output | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return [pscustomobject]@{
            Succeeded = $false
            Existing = $false
            TargetEntry = $null
            Summary = 'CLI Marketplace 列表不是可验证的 JSON。'
        }
    }

    if ($parsed -is [System.Array]) {
        $entries = @($parsed)
    } elseif ($null -ne $parsed.PSObject.Properties['marketplaces'] -and $parsed.marketplaces -is [System.Array]) {
        $entries = @($parsed.marketplaces)
    } elseif ($null -ne $parsed.PSObject.Properties['items'] -and $parsed.items -is [System.Array]) {
        $entries = @($parsed.items)
    } elseif ($null -ne $parsed.PSObject.Properties['data'] -and $parsed.data -is [System.Array]) {
        $entries = @($parsed.data)
    } else {
        return [pscustomobject]@{
            Succeeded = $false
            Existing = $false
            TargetEntry = $null
            Summary = 'CLI Marketplace JSON 未提供可验证的条目数组。'
        }
    }

    $targetEntries = @($entries | Where-Object {
            $null -ne $_ -and $_ -isnot [string] -and $_ -isnot [System.ValueType] -and
            $null -ne $_.PSObject.Properties['name'] -and [string]$_.name -eq $MarketplaceName
        })
    if ($targetEntries.Count -gt 1) {
        return [pscustomobject]@{
            Succeeded = $false
            Existing = $true
            TargetEntry = $null
            Summary = 'CLI Marketplace 列表包含多个同名目标条目，无法安全自动处理。'
        }
    }
    return [pscustomobject]@{
        Succeeded = $true
        Existing = ($targetEntries.Count -eq 1)
        TargetEntry = if ($targetEntries.Count -eq 1) { $targetEntries[0] } else { $null }
        Summary = ''
    }
}

function Test-CodexMarketplaceEntryMatchesRoot {
    <#
      rollback 前要求列表中的条目同时具备本插件名称和本次已验证发行根路径。CLI 若
      不公开路径字段，便不能证明该条目属于本次事务，因此不会执行 remove。
    #>
    param(
        [Parameter(Mandatory)]$Entry,
        [Parameter(Mandatory)][string]$VerifiedMarketplaceRoot
    )

    if ($null -eq $Entry -or [string]$Entry.name -ne $MarketplaceName) {
        return $false
    }
    $candidatePath = $null
    if ($null -ne $Entry.PSObject.Properties['path']) {
        $candidatePath = [string]$Entry.path
    } elseif ($null -ne $Entry.PSObject.Properties['source'] -and $null -ne $Entry.source -and $null -ne $Entry.source.PSObject.Properties['path']) {
        $candidatePath = [string]$Entry.source.path
    } elseif ($null -ne $Entry.PSObject.Properties['marketplaceRoot']) {
        $candidatePath = [string]$Entry.marketplaceRoot
    } elseif ($null -ne $Entry.PSObject.Properties['root']) {
        # 当前官方 CLI 的 list --json 使用 root 字段。
        $candidatePath = [string]$Entry.root
    }
    if ([string]::IsNullOrWhiteSpace($candidatePath)) {
        return $false
    }
    try {
        $expected = [System.IO.Path]::GetFullPath($VerifiedMarketplaceRoot).TrimEnd([char]'\', [char]'/')
        $actual = [System.IO.Path]::GetFullPath($candidatePath).TrimEnd([char]'\', [char]'/')
        return $actual.Equals($expected, [System.StringComparison]::OrdinalIgnoreCase)
    } catch {
        return $false
    }
}

function Remove-ThisRunCodexMarketplaceAndVerify {
    <#
      只在“当前列表仍恰好指向本次已验证根目录”时删除。任何 list 失败、条目消失、
      同名条目改指向其他根目录或 remove 后状态不明确，均视为事务状态未知并终止；
      调用方绝不能在这种情况下继续个人 Marketplace 回退。
    #>
    param(
        [Parameter(Mandatory)][string]$ExecutablePath,
        [Parameter(Mandatory)][string]$VerifiedMarketplaceRoot,
        [Parameter(Mandatory)][string]$Context
    )

    $beforeRemoval = Get-CodexMarketplaceInventory -ExecutablePath $ExecutablePath
    if (-not $beforeRemoval.Succeeded) {
        throw "$Context 后无法读取 Codex Marketplace 状态；为避免留下未知登记，已停止，未执行个人 Marketplace 回退。"
    }
    if (-not $beforeRemoval.Existing) {
        # 已由可靠 JSON 列表证明没有残留条目，无需 remove，调用方可安全回退。
        return $false
    }
    if (-not (Test-CodexMarketplaceEntryMatchesRoot -Entry $beforeRemoval.TargetEntry -VerifiedMarketplaceRoot $VerifiedMarketplaceRoot)) {
        throw "$Context 后同名 Marketplace 未精确指向本次发行根；拒绝 remove，已停止。"
    }

    $removal = Invoke-CodexCliCommand -ExecutablePath $ExecutablePath -Arguments ([string[]]@('plugin', 'marketplace', 'remove', $MarketplaceName))
    if (-not $removal.Succeeded) {
        throw "$Context 后无法回滚本次 Marketplace 登记；为避免留下未知状态，已停止，未执行个人 Marketplace 回退。"
    }
    $afterRemoval = Get-CodexMarketplaceInventory -ExecutablePath $ExecutablePath
    if (-not $afterRemoval.Succeeded -or $afterRemoval.Existing) {
        throw "$Context 后无法确认 Marketplace 已被移除；为避免留下未知状态，已停止，未执行个人 Marketplace 回退。"
    }
    return $true
}

function Try-InstallWithCodexCli {
    <#
      CLI 安装成功时保持零摩擦体验。只有在 JSON 列表已确认目标 Marketplace 不存在
      时才会 add。一旦 add 被尝试，后续状态必须明确：目标确实不存在，或能精确证明
      指向本次发行根、remove 成功且最终列表确认不存在。状态不确定时 fail-closed，
      不继续个人 Marketplace 回退，更不删除用户预先存在的同名条目。
    #>
    param(
        [Parameter(Mandatory)][string]$ExecutablePath,
        [Parameter(Mandatory)][string]$VerifiedMarketplaceRoot
    )

    $before = Get-CodexMarketplaceInventory -ExecutablePath $ExecutablePath
    if (-not $before.Succeeded) {
        return [pscustomobject]@{
            Succeeded = $false
            Stage = '查询 Marketplace 状态'
            Summary = $before.Summary
            RollbackSucceeded = $true
        }
    }
    if ($before.Existing) {
        return [pscustomobject]@{
            Succeeded = $false
            Stage = '检测到已有 Marketplace'
            Summary = "已存在名为 $MarketplaceName 的 Marketplace；为避免改变用户原有登记，改用个人 Marketplace。"
            RollbackSucceeded = $true
        }
    }

    # 从这里开始即使 add 返回非零也不能假定“没有写入”：CLI 可能已经完成登记后才报错。
    $registration = Invoke-CodexCliCommand -ExecutablePath $ExecutablePath -Arguments ([string[]]@('plugin', 'marketplace', 'add', $VerifiedMarketplaceRoot))
    $afterRegistration = Get-CodexMarketplaceInventory -ExecutablePath $ExecutablePath
    if (-not $afterRegistration.Succeeded) {
        throw '尝试登记 Codex Marketplace 后无法读取其最终状态；为避免留下未知登记，已停止，未执行个人 Marketplace 回退。'
    }
    if (-not $afterRegistration.Existing) {
        return [pscustomobject]@{
            Succeeded = $false
            Stage = '注册 Marketplace'
            Summary = if ($registration.Succeeded) { 'CLI 返回成功但最终列表未显示目标 Marketplace；已确认无残留，改用个人 Marketplace。' } else { "$($registration.Summary)；最终列表确认目标 Marketplace 不存在，改用个人 Marketplace。" }
            RollbackSucceeded = $true
        }
    }
    if (-not (Test-CodexMarketplaceEntryMatchesRoot -Entry $afterRegistration.TargetEntry -VerifiedMarketplaceRoot $VerifiedMarketplaceRoot)) {
        # 绝不能对 root 不匹配的同名条目 remove；它可能是并发新增的用户配置。
        throw '尝试登记 Codex Marketplace 后发现同名条目未精确指向本次发行根；拒绝删除未知条目，已停止，未执行个人 Marketplace 回退。'
    }
    if (-not $registration.Succeeded) {
        Remove-ThisRunCodexMarketplaceAndVerify -ExecutablePath $ExecutablePath -VerifiedMarketplaceRoot $VerifiedMarketplaceRoot -Context 'Codex Marketplace add 返回失败'
        return [pscustomobject]@{
            Succeeded = $false
            Stage = '注册 Marketplace'
            Summary = "$($registration.Summary)；已确认并回滚本次登记，改用个人 Marketplace。"
            RollbackSucceeded = $true
        }
    }

    $pluginInstall = Invoke-CodexCliCommand -ExecutablePath $ExecutablePath -Arguments ([string[]]@('plugin', 'add', "$PluginName@$MarketplaceName"))
    if (-not $pluginInstall.Succeeded) {
        Remove-ThisRunCodexMarketplaceAndVerify -ExecutablePath $ExecutablePath -VerifiedMarketplaceRoot $VerifiedMarketplaceRoot -Context 'Codex 插件安装失败'
        return [pscustomobject]@{
            Succeeded = $false
            Stage = '安装插件'
            Summary = "$($pluginInstall.Summary)；已确认并回滚本次 Marketplace 登记，改用个人 Marketplace。"
            RollbackSucceeded = $true
        }
    }

    return [pscustomobject]@{
        Succeeded = $true
        Stage = ''
        Summary = ''
        RollbackSucceeded = $true
    }
}

function Write-InstallationResult {
    <#
      引导安装器需要一个且仅一个机器可读的最终结果。所有过程信息均写到 Host，
      这里以单行 JSON 输出，避免上层脚本错误地把验证器日志当作安装状态。
    #>
    param(
        [Parameter(Mandatory)][string]$InstallationMode,
        [Parameter(Mandatory)][bool]$PluginInstalled,
        [Parameter(Mandatory)][bool]$CodexCliCalled,
        [Parameter(Mandatory)][string]$NextStep,
        [string]$MarketplacePath,
        [string]$Detail
    )

    [pscustomobject]@{
        pass = $true
        installation_mode = $InstallationMode
        plugin_installed = $PluginInstalled
        codex_cli_called = $CodexCliCalled
        next_step = $NextStep
        marketplace_path = $MarketplacePath
        detail = $Detail
        node_required = $false
    } | ConvertTo-Json -Depth 4 -Compress
}

if ($env:OS -ne 'Windows_NT') {
    throw 'Paper to Journal Club 目前仅支持 Windows，因为它通过 PowerPoint COM 控制桌面版 Microsoft PowerPoint。'
}
if (-not [Environment]::Is64BitOperatingSystem) {
    throw '随包 PDF 解析器需要 64 位 Windows。'
}

# 在读取清单、解析 JSON 或执行包内验证器前，先安全遍历整个已解压发行目录。
# 不使用 Resolve-Path，以免它先解引用攻击者放置的 junction/symlink。
$resolvedRoot = Assert-NoReparsePointsInDirectoryTree -Root $MarketplaceRoot -Purpose '已解压发行包根目录'
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
$packageChecksums = Test-Checksums -Root $resolvedRoot
Assert-AllowedMarketplaceAndMcpConfiguration -Root $resolvedRoot
# 执行包内验证器前再次检查，避免首次读取后文件被替换。
$packageChecksums = Test-Checksums -Root $resolvedRoot
# 验证器会输出自己的 JSON 报告；将其捕获，确保包内安装器最终只输出一个结构化结果。
$verifierReport = & $verifierPath -MarketplaceRoot $resolvedRoot
# 部署前保留一次刚刚复核过的清单，并让复制阶段直接对照该清单。
$packageChecksums = Test-Checksums -Root $resolvedRoot
$expectedPluginChecksums = Get-PluginChecksumMap -PackageChecksums $packageChecksums

if (-not $SkipPowerPointCheck) {
    $powerPointClass = 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID'
    if (-not (Test-Path -LiteralPath $powerPointClass)) {
        throw '未检测到 Microsoft PowerPoint 桌面版。请安装并启动一次 PowerPoint 后重新运行安装器。'
    }
}

Write-Host '发行包完整性和运行前置检查通过。'
Write-Host '终端用户不需要 Node.js、npm、Python 或 .NET SDK。'

if ($SkipPersonalMarketplaceDeployment) {
    Write-Host '按请求仅完成发行包验证，跳过 Codex CLI 安装和个人 Marketplace 部署。'
    Write-InstallationResult -InstallationMode 'verification-only' -PluginInstalled $false -CodexCliCalled $false -NextStep '已按请求跳过安装；移除此开关后重新运行安装器。' -Detail 'skip-install-requested'
    return
}

$codexCliCalled = $false
$cliFallbackDetail = ''
$codexCliPath = Get-ExecutableCodexCliPath
if ($null -ne $codexCliPath) {
    # CLI 会在其用户目录内管理临时数据，因此不再执行“看似只读”的 --help 探针；
    # 用户批准后直接执行带事务回滚的安装流程。
    if (-not $PSCmdlet.ShouldProcess(
        "$PluginName@$MarketplaceName",
        "使用已验证签名的 Codex CLI 注册已校验 Marketplace 并安装插件 ($codexCliPath)"
    )) {
        Write-InstallationResult -InstallationMode 'what-if' -PluginInstalled $false -CodexCliCalled $false -NextStep '确认后重新运行安装器，以尝试 CLI 安装或个人 Marketplace 回退。' -Detail 'cli-install-not-approved'
        return
    }

    $codexCliCalled = $true
    $cliInstallation = Try-InstallWithCodexCli -ExecutablePath $codexCliPath -VerifiedMarketplaceRoot $resolvedRoot
    if ($cliInstallation.Succeeded) {
        Write-Host "已通过 Codex CLI 安装 $PluginName@$MarketplaceName。"
        Write-Host '请完全退出并重新打开 Codex，然后新建一个任务即可使用插件。'
        Write-InstallationResult -InstallationMode 'codex-cli' -PluginInstalled $true -CodexCliCalled $true -NextStep '请完全退出并重新打开 Codex，然后新建一个任务即可使用插件。' -MarketplacePath $resolvedRoot -Detail 'trusted-cli-install-succeeded'
        return
    }
    $failureDetail = if ([string]::IsNullOrWhiteSpace([string]$cliInstallation.Summary)) { '未返回额外错误信息。' } else { $cliInstallation.Summary }
    $cliFallbackDetail = "CLI $($cliInstallation.Stage) 未成功：$failureDetail"
    if (-not $cliInstallation.RollbackSucceeded) {
        $cliFallbackDetail += '；未能确认 CLI Marketplace 回滚，个人 Marketplace 回退不会删除任何未知 CLI 条目。'
    }
    Write-Warning "$cliFallbackDetail 将自动改用个人 Marketplace 安装方式。"
} else {
    $cliFallbackDetail = '未找到受信、非 WindowsApps 的 OpenAI 签名 codex.exe。'
    Write-Host "$cliFallbackDetail 将自动改用个人 Marketplace 安装方式。"
}

$userProfilePath = Get-CurrentUserProfilePath
$personalPluginPath = Join-Path $userProfilePath ".codex\plugins\$PluginName"
$personalMarketplacePath = Join-Path $userProfilePath '.agents\plugins\marketplace.json'

# CLI 不可用时，这是不会编辑 config.toml 的安全回退路径。用户在重启后从
# Plugins Directory 明确点击 Install，保留最终安装确认。
if (-not $PSCmdlet.ShouldProcess(
    "$personalPluginPath 和 $personalMarketplacePath",
    '部署经校验的 Paper to Journal Club 个人 Marketplace'
)) {
    Write-InstallationResult -InstallationMode 'what-if' -PluginInstalled $false -CodexCliCalled $codexCliCalled -NextStep '确认后重新运行安装器，以部署个人 Marketplace。' -MarketplacePath $personalMarketplacePath -Detail $cliFallbackDetail
    return
}

try {
    $deployedPluginPath = Deploy-VerifiedPluginCopy -SourcePluginRoot $pluginRoot -UserProfilePath $userProfilePath -ExpectedPluginChecksums $expectedPluginChecksums
    $marketplaceUpdate = Update-PersonalMarketplaceFile -UserProfilePath $userProfilePath
} catch {
    throw "个人 Marketplace 部署失败：$($_.Exception.Message)"
}

Write-Host '已安全部署 Paper to Journal Club 的已验证副本。'
Write-Host "插件目录：$deployedPluginPath"
Write-Host "个人 Marketplace：$($marketplaceUpdate.Path)"
if ($null -ne $marketplaceUpdate.BackupPath) {
    Write-Host "已保留原个人 Marketplace 备份：$($marketplaceUpdate.BackupPath)"
}
Write-Host '请完全退出并重新打开 Codex，打开 Plugins，选择个人 Marketplace，找到 Paper to Journal Club 并点击 Install。'
Write-Host '完成 Install 后新建一个任务即可使用；无需手动运行 Codex CLI，也无需修改 WindowsApps 权限。'
Write-InstallationResult -InstallationMode 'personal-marketplace' -PluginInstalled $false -CodexCliCalled $codexCliCalled -NextStep '请完全退出并重新打开 Codex，然后在 Plugins Directory 的个人 Marketplace 中找到 Paper to Journal Club 并点击 Install。' -MarketplacePath $marketplaceUpdate.Path -Detail $cliFallbackDetail
