<#
  Paper to Journal Club 的 GitHub Release 自助安装器。

  此脚本不克隆开发仓库，也不会要求终端用户安装 Git、Node.js、npm、Python 或 .NET SDK。
  它从指定 GitHub 仓库的一个固定 Release 下载已经构建好的本地 Marketplace ZIP，先校验
  同一 Release 中的 SHA-256 文件，再把通过校验的包交给包内安装器完成发行包级别的验证。
  用户确认后，包内安装器会验证受信 Codex CLI 并尝试自动安装；CLI 自身可能写入临时数据。
  若 CLI 不可用或自动安装失败，安装器会回滚本次 Marketplace 注册并安全回退到个人 Marketplace。

  为避免把“最新版本”误装到科研用户的电脑，RepositoryUrl 和 ReleaseTag 都是必填参数。
  发布者应只为已经审核的不可变标签创建 Release，例如 v1.0.0。
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # 必须是发布者实际拥有的 GitHub HTTPS 仓库地址；脚本不会内置或猜测仓库地址。
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$RepositoryUrl,

    # 必须是不可变的 GitHub Release 标签，例如 v1.0.0。
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ReleaseTag,

    # 当同一个 Release 含有多个 Marketplace ZIP 时，用此参数精确指定其中一个文件名。
    [string]$ReleaseAssetName,

    # 安装目录默认为当前用户的 LocalAppData，避免写入系统目录或依赖临时文件夹。
    [string]$InstallDirectory,

    # 仅供 CI、排障或已明确知晓后果的场景跳过 Microsoft PowerPoint 桌面版检查。
    [switch]$SkipPowerPointCheck,

    # 普通用户不需要此开关；它只供 CI 或维护者仅验证发行包，不尝试 CLI 自动安装或个人 Marketplace 回退。
    # SkipCodexInstall 是旧版兼容别名，保留后仍表示跳过全部安装分支。
    [Alias('SkipCodexInstall')]
    [switch]$SkipPersonalMarketplaceDeployment
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'
$MaximumReleaseBytes = 512MB
# ZIP 中的单个文件、总解压量和压缩比都受限，避免小型恶意 ZIP 耗尽磁盘或内存。
$MaximumArchiveEntries = 256
$MaximumExpandedBytes = 1GB
$MaximumSingleEntryBytes = 256MB
$MaximumCompressionRatio = 200

function Get-ExistingPathAttributes {
    <#
      使用 File.GetAttributes 检查“目录项本身”的属性，而不是先解析链接目标。
      对不存在的路径返回 Exists = false；权限或 I/O 异常则继续抛出，不能把异常误当成安全路径。
    #>
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
      逐级检查绝对路径中已经存在的部分，拒绝符号链接、联接点和其他重解析点。
      不存在的末级或中间目录是正常安装场景，不能因此误报；但已有的中间项必须是目录。
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
            # 后续子目录也不可能安全地“已存在”；保留其字面路径，供创建操作使用。
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
    <#
      读取或移动一个既有目录前，逐个目录项检查其属性。使用队列而非 -Recurse，
      避免 PowerShell 在枚举时跟随目录联接点进入目录外。
    #>
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
            $childAttributes = Get-ExistingPathAttributes -Path $child.FullName
            if (-not $childAttributes.Exists) {
                throw "$Purpose 在检查期间发生并发路径变化：$($child.FullName)"
            }
            if (($childAttributes.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
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
    <# 创建目录前后都检查其现有路径链，防止联接点在创建窗口中重定向写入。 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose
    )

    $fullPath = Assert-NoReparsePointInExistingPath -Path $Path -Purpose $Purpose
    $existingAttributes = Get-ExistingPathAttributes -Path $fullPath
    if ($existingAttributes.Exists -and -not [System.IO.Directory]::Exists($fullPath)) {
        throw "$Purpose 已存在但不是目录：$fullPath"
    }
    [System.IO.Directory]::CreateDirectory($fullPath) | Out-Null
    $fullPath = Assert-NoReparsePointInExistingPath -Path $fullPath -Purpose $Purpose -RequireExisting
    if (-not [System.IO.Directory]::Exists($fullPath)) {
        throw "$Purpose 创建后不是目录：$fullPath"
    }
    return $fullPath
}

function Assert-SafeNewPath {
    <# 仅接受位于安全既有父目录下、且尚未存在的目标，避免覆盖链接或普通用户文件。 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose
    )

    $fullPath = Assert-NoReparsePointInExistingPath -Path $Path -Purpose $Purpose
    $attributes = Get-ExistingPathAttributes -Path $fullPath
    if ($attributes.Exists) {
        throw "$Purpose 已存在，拒绝覆盖：$fullPath"
    }
    $parentPath = Split-Path -Parent $fullPath
    $parentPath = Assert-NoReparsePointInExistingPath -Path $parentPath -Purpose "$Purpose 的父目录" -RequireExisting
    if (-not [System.IO.Directory]::Exists($parentPath)) {
        throw "$Purpose 的父路径不是目录：$parentPath"
    }
    return $fullPath
}

function Assert-SafeNewFilePath {
    <# 下载和 ZIP 解压的文件目标使用通用的新路径检查，并保留明确的函数名。 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose
    )

    return (Assert-SafeNewPath -Path $Path -Purpose $Purpose)
}

function Assert-SafeExistingFilePath {
    <# 对下载缓存、校验文件和 ZIP 读取路径做同样的逐级重解析点检查。 #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose
    )

    $fullPath = Assert-NoReparsePointInExistingPath -Path $Path -Purpose $Purpose -RequireExisting
    if (-not [System.IO.File]::Exists($fullPath)) {
        throw "$Purpose 必须是普通文件：$fullPath"
    }
    return $fullPath
}

function New-UniqueChildPath {
    <# 返回尚未创建的随机子路径；允许路径不存在，但父目录必须是安全的实体目录。 #>
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
        $candidatePath = Join-Path $parentPath ("{0}{1}" -f $Prefix, [Guid]::NewGuid().ToString('N'))
        $candidatePath = Assert-NoReparsePointInExistingPath -Path $candidatePath -Purpose $Purpose
        if (-not (Get-ExistingPathAttributes -Path $candidatePath).Exists) {
            return $candidatePath
        }
    }
    throw "$Purpose 无法生成未占用的随机路径。"
}

function Move-SafeInstallItem {
    <#
      所有候选/current/previous/failed 的移动都经由此处：移动前检查源、目标及双方父链，
      移动后再次检查新路径，避免把目录移动进联接点或从链接目标读取。
    #>
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$Purpose
    )

    $sourcePath = Assert-NoReparsePointInExistingPath -Path $Source -Purpose "$Purpose 的源路径" -RequireExisting
    $sourceAttributes = Get-ExistingPathAttributes -Path $sourcePath
    if (-not $sourceAttributes.Exists) {
        throw "$Purpose 的源路径不存在：$sourcePath"
    }
    $destinationPath = Assert-SafeNewPath -Path $Destination -Purpose "$Purpose 的目标路径"
    Move-Item -LiteralPath $sourcePath -Destination $destinationPath -ErrorAction Stop
    $destinationPath = Assert-NoReparsePointInExistingPath -Path $destinationPath -Purpose "$Purpose 的移动后路径" -RequireExisting
    return $destinationPath
}

function Get-RepositoryInformation {
    <# 将 URL 收敛为 GitHub owner/repository，拒绝查询串、凭据和非 HTTPS 地址。 #>
    param([Parameter(Mandatory)][string]$Url)

    $uri = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) {
        throw 'RepositoryUrl 必须是完整的 GitHub HTTPS 仓库地址。'
    }
    if ($uri.Scheme -ne 'https' -or -not [string]::Equals($uri.Host, 'github.com', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'RepositoryUrl 只允许 https://github.com/<owner>/<repository>[.git] 形式的地址。'
    }
    if (-not [string]::IsNullOrWhiteSpace($uri.UserInfo) -or -not [string]::IsNullOrWhiteSpace($uri.Query) -or -not [string]::IsNullOrWhiteSpace($uri.Fragment)) {
        throw 'RepositoryUrl 不能包含用户名、密码、查询串或片段。'
    }

    $segments = @($uri.AbsolutePath.Trim('/') -split '/')
    if ($segments.Count -ne 2 -or [string]::IsNullOrWhiteSpace($segments[0]) -or [string]::IsNullOrWhiteSpace($segments[1])) {
        throw 'RepositoryUrl 必须精确指向一个 GitHub owner/repository。'
    }

    $owner = $segments[0]
    $repository = $segments[1]
    if ($repository.EndsWith('.git', [System.StringComparison]::OrdinalIgnoreCase)) {
        $repository = $repository.Substring(0, $repository.Length - 4)
    }
    if ($owner -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$' -or $repository -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$') {
        throw 'RepositoryUrl 中的 owner 和 repository 只能包含字母、数字、点、下划线和连字符。'
    }

    return [pscustomobject]@{
        Owner = $owner
        Repository = $repository
        CanonicalUrl = "https://github.com/$owner/$repository"
    }
}

function Assert-ReleaseTag {
    <# 限制标签字符，避免把参数解释为 URL 路径或命令行选项。 #>
    param([Parameter(Mandatory)][string]$Tag)

    if ($Tag -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$') {
        throw 'ReleaseTag 只能包含字母、数字、点、下划线和连字符，且必须以字母或数字开头。'
    }
}

function Get-DefaultInstallDirectory {
    $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        $localAppData = Join-Path $HOME '.codex'
    }
    return Join-Path $localAppData 'Codex\marketplaces\paper-to-journal-club'
}

function Assert-SafeInstallDirectory {
    <# 禁止驱动器根目录及路径链中的重解析点，后续所有写入只发生在专用实体目录。 #>
    param([Parameter(Mandatory)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $trimmedPath = $fullPath.TrimEnd([char]'\', [char]'/')
    $trimmedRoot = [System.IO.Path]::GetPathRoot($fullPath).TrimEnd([char]'\', [char]'/')
    if ([string]::Equals($trimmedPath, $trimmedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'InstallDirectory 不能是驱动器根目录。请选择一个专用的用户目录。'
    }
    return (Assert-NoReparsePointInExistingPath -Path $fullPath -Purpose 'InstallDirectory')
}

function Get-GitHubRelease {
    param(
        [Parameter(Mandatory)]$Repository,
        [Parameter(Mandatory)][string]$Tag
    )

    # Windows PowerShell 5.1 默认的 TLS 配置可能过旧；只补充 TLS 1.2，不降低现有安全级别。
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $escapedTag = [System.Uri]::EscapeDataString($Tag)
    $apiUrl = "https://api.github.com/repos/$($Repository.Owner)/$($Repository.Repository)/releases/tags/$escapedTag"
    $headers = @{
        'Accept' = 'application/vnd.github+json'
        'User-Agent' = 'paper-to-journal-club-installer'
    }

    try {
        $release = Invoke-RestMethod -Uri $apiUrl -Headers $headers -UseBasicParsing
    } catch {
        throw "无法读取 GitHub Release '$Tag'。请确认仓库地址、标签和网络可用：$($_.Exception.Message)"
    }

    if ($null -eq $release -or $release.draft -or $release.prerelease) {
        throw "GitHub Release '$Tag' 不可用、仍是草稿或属于预发布版本。"
    }
    return $release
}

function Assert-TrustedGitHubAssetUrl {
    <# GitHub API 中的资产下载 URL 也必须回到 github.com 的 Release 下载路由，拒绝任意 HTTPS 外链。 #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$ExpectedAssetName
    )

    $uri = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or -not [string]::Equals($uri.Host, 'github.com', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'GitHub Release 返回的资产下载地址不受信任。'
    }
    if ($uri.AbsolutePath -notmatch '^/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+/releases/download/') {
        throw 'GitHub Release 资产下载地址不是预期的 Release 路径。'
    }
    $downloadName = [System.Uri]::UnescapeDataString($uri.Segments[$uri.Segments.Length - 1])
    if ($downloadName -ne $ExpectedAssetName) {
        throw 'GitHub Release 资产下载地址与资产名称不一致。'
    }
}

function Get-ReleaseAsset {
    param(
        [Parameter(Mandatory)]$Release,
        [AllowEmptyString()][string]$RequestedName
    )

    $assets = @($Release.assets)
    if ([string]::IsNullOrWhiteSpace($RequestedName)) {
        $candidates = @($assets | Where-Object { $_.name -match '^paper-to-journal-club-marketplace-[A-Za-z0-9._-]+\.zip$' })
        if ($candidates.Count -ne 1) {
            throw "Release 中必须恰好有一个 Marketplace ZIP；当前找到 $($candidates.Count) 个。请使用 -ReleaseAssetName 明确指定。"
        }
        $asset = $candidates[0]
    } else {
        if ($RequestedName -notmatch '^paper-to-journal-club-marketplace-[A-Za-z0-9._-]+\.zip$') {
            throw 'ReleaseAssetName 不是预期的 Paper to Journal Club Marketplace ZIP 文件名。'
        }
        $candidates = @($assets | Where-Object { $_.name -eq $RequestedName })
        if ($candidates.Count -ne 1) {
            throw "Release 中未找到指定的 ZIP 资产：$RequestedName"
        }
        $asset = $candidates[0]
    }

    Assert-TrustedGitHubAssetUrl -Url ([string]$asset.browser_download_url) -ExpectedAssetName ([string]$asset.name)
    if ([int64]$asset.size -le 0 -or [int64]$asset.size -gt $MaximumReleaseBytes) {
        throw "ZIP 资产大小异常：$($asset.size) 字节。允许的最大值为 $MaximumReleaseBytes 字节。"
    }
    return $asset
}

function Get-ChecksumAsset {
    param(
        [Parameter(Mandatory)]$Release,
        [Parameter(Mandatory)]$ArchiveAsset
    )

    $checksumName = "$($ArchiveAsset.name).sha256"
    $matches = @($Release.assets | Where-Object { $_.name -eq $checksumName })
    if ($matches.Count -ne 1) {
        throw "Release 必须包含 ZIP 对应的 SHA-256 文件：$checksumName"
    }
    Assert-TrustedGitHubAssetUrl -Url ([string]$matches[0].browser_download_url) -ExpectedAssetName ([string]$matches[0].name)
    return $matches[0]
}

function Save-GitHubAsset {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$Destination
    )

    $Destination = Assert-SafeNewFilePath -Path $Destination -Purpose 'GitHub Release 下载目标'
    $headers = @{
        'Accept' = 'application/octet-stream'
        'User-Agent' = 'paper-to-journal-club-installer'
    }
    try {
        Invoke-WebRequest -Uri $Url -Headers $headers -OutFile $Destination -UseBasicParsing
    } catch {
        throw "下载 Release 资产失败：$($_.Exception.Message)"
    }
    $Destination = Assert-SafeExistingFilePath -Path $Destination -Purpose '下载的 GitHub Release 资产'
    if ((Get-Item -LiteralPath $Destination -ErrorAction Stop).Length -eq 0) {
        throw '下载的 Release 资产为空或不存在。'
    }
}

function Get-ExpectedArchiveHash {
    param(
        [Parameter(Mandatory)][string]$ChecksumPath,
        [Parameter(Mandatory)][string]$ArchiveName
    )

    $ChecksumPath = Assert-SafeExistingFilePath -Path $ChecksumPath -Purpose 'SHA-256 校验文件'
    $lines = @(Get-Content -LiteralPath $ChecksumPath -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -ne 1) {
        throw 'SHA-256 文件必须只包含一条非空校验记录。'
    }
    $match = [System.Text.RegularExpressions.Regex]::Match($lines[0], '^\s*([A-Fa-f0-9]{64})(?:\s+\*?(.+?))?\s*$')
    if (-not $match.Success) {
        throw 'SHA-256 文件格式无效。'
    }
    if ($match.Groups[2].Success -and -not [string]::Equals($match.Groups[2].Value.Trim(), $ArchiveName, [System.StringComparison]::Ordinal)) {
        throw 'SHA-256 文件指向的 ZIP 文件名与 Release 资产不一致。'
    }
    return $match.Groups[1].Value.ToUpperInvariant()
}

function Assert-ArchiveHash {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedHash
    )

    $Path = Assert-SafeExistingFilePath -Path $Path -Purpose 'ZIP 校验文件'
    $actualHash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actualHash -ne $ExpectedHash) {
        throw "ZIP 的 SHA-256 校验失败。期望：$ExpectedHash；实际：$actualHash"
    }
}

function Assert-ExactPropertyNames {
    <# 仅允许发行版预期的 JSON 字段，阻断第二个 Marketplace/MCP 入口或额外启动参数。 #>
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

function ConvertTo-CanonicalRelativePath {
    <# 将 ZIP/哈希清单路径标准化为相对正斜杠路径，拒绝路径穿越、驱动器和 ADS 路径。 #>
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or $Path.IndexOf([char]0) -ge 0) {
        throw '路径为空或包含 NUL 字符。'
    }
    $normalized = $Path.Replace('\', '/').Trim()
    if ($normalized.StartsWith('/') -or $normalized.EndsWith('/') -or $normalized.Contains(':')) {
        throw "路径不安全：$Path"
    }
    $segments = @($normalized -split '/')
    if ($segments.Count -eq 0 -or @($segments | Where-Object { $_ -in @('', '.', '..') }).Count -gt 0) {
        throw "路径不安全：$Path"
    }
    return $normalized
}

function Get-RelativePath {
    <# 兼容 Windows PowerShell 5.1：.NET Framework 不存在 Path.GetRelativePath。 #>
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$TargetPath
    )

    $baseUri = [System.Uri]::new((Join-Path $BasePath ''))
    $targetUri = [System.Uri]::new($TargetPath)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString()).Replace('\', '/')
}

function Test-ChecksumManifest {
    <#
      清单不只校验“列出的文件没有变化”，还要求除 SHA256SUMS.txt 外的实际文件集合完全相等。
      因此即使攻击者往解压目录额外写入一个未列出的脚本或配置，也会被拒绝。
    #>
    param([Parameter(Mandatory)][string]$Root)

    $Root = Assert-NoReparsePointsInDirectoryTree -Root $Root -Purpose '发行包目录'
    $checksumPath = Join-Path $Root 'SHA256SUMS.txt'
    $checksumPath = Assert-SafeExistingFilePath -Path $checksumPath -Purpose '发行包 SHA-256 清单'
    if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
        throw '发行包缺少 SHA256SUMS.txt。'
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
    foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File -Force) {
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

    $missingEntries = @($manifestEntries.Keys | Where-Object { -not $actualEntries.ContainsKey($_) })
    $unexpectedEntries = @($actualEntries.Keys | Where-Object { -not $manifestEntries.ContainsKey($_) })
    if ($missingEntries.Count -gt 0 -or $unexpectedEntries.Count -gt 0) {
        $details = @()
        if ($missingEntries.Count -gt 0) { $details += "清单列出但实际缺失：$($missingEntries -join ', ')" }
        if ($unexpectedEntries.Count -gt 0) { $details += "实际存在但未列入清单：$($unexpectedEntries -join ', ')" }
        throw "SHA-256 清单双向集合校验失败：$($details -join '；')"
    }

    foreach ($relativePath in $manifestEntries.Keys) {
        $actualHash = (Get-FileHash -LiteralPath $actualEntries[$relativePath] -Algorithm SHA256).Hash.ToUpperInvariant()
        if ($actualHash -ne $manifestEntries[$relativePath]) {
            throw "SHA-256 不匹配：$relativePath"
        }
    }
}

function Get-SafeZipEntryPath {
    param([Parameter(Mandatory)][string]$EntryName)

    # ZipArchive 目录项以 / 结尾；去除目录标记后仍需通过同一条路径安全规则。
    $withoutDirectoryMarker = $EntryName -replace '[\\/]+$', ''
    return ConvertTo-CanonicalRelativePath -Path $withoutDirectoryMarker
}

function Assert-SafeZipArchive {
    <# 在真正写盘前审阅 ZIP 中央目录，限制条目数、声明解压量和高压缩比。 #>
    param([Parameter(Mandatory)][string]$ArchivePath)

    $ArchivePath = Assert-SafeExistingFilePath -Path $ArchivePath -Purpose '待解压 ZIP 文件'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $entries = @($archive.Entries)
        if ($entries.Count -eq 0 -or $entries.Count -gt $MaximumArchiveEntries) {
            throw "ZIP 条目数异常：$($entries.Count)，允许范围为 1 到 $MaximumArchiveEntries。"
        }

        $seenPaths = @{}
        [int64]$totalLength = 0
        foreach ($entry in $entries) {
            $relativePath = Get-SafeZipEntryPath -EntryName $entry.FullName
            if ($seenPaths.ContainsKey($relativePath)) {
                throw "ZIP 包含重复或大小写冲突的路径：$relativePath"
            }
            $seenPaths[$relativePath] = $true

            # Name 为空代表目录项；目录本身不消耗解压配额。
            if ([string]::IsNullOrEmpty($entry.Name)) {
                continue
            }
            if ($entry.Length -lt 0 -or $entry.Length -gt $MaximumSingleEntryBytes) {
                throw "ZIP 条目解压大小异常：$relativePath"
            }
            $totalLength += [int64]$entry.Length
            if ($totalLength -gt $MaximumExpandedBytes) {
                throw "ZIP 声明的总解压量超过限制：$MaximumExpandedBytes 字节。"
            }
            if ($entry.Length -gt 0) {
                if ($entry.CompressedLength -le 0 -or ([double]$entry.Length / [double]$entry.CompressedLength) -gt $MaximumCompressionRatio) {
                    throw "ZIP 条目压缩比异常：$relativePath"
                }
            }
        }
    } finally {
        $archive.Dispose()
    }
}

function Expand-SafeZipArchive {
    <#
      逐条目流式解压，而非直接使用 Expand-Archive：每次写入都重新限制实际字节数，
      同时以规范化后的绝对路径验证目标始终位于候选目录内。
    #>
    param(
        [Parameter(Mandatory)][string]$ArchivePath,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    $ArchivePath = Assert-SafeExistingFilePath -Path $ArchivePath -Purpose '待解压 ZIP 文件'
    Assert-SafeZipArchive -ArchivePath $ArchivePath
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $destinationRoot = Assert-NoReparsePointsInDirectoryTree -Root $DestinationPath -Purpose 'ZIP 候选解压目录'
    $destinationRoot = $destinationRoot.TrimEnd([char]'\', [char]'/')
    $destinationPrefix = "$destinationRoot$([System.IO.Path]::DirectorySeparatorChar)"
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    [int64]$totalWritten = 0
    try {
        foreach ($entry in $archive.Entries) {
            $relativePath = Get-SafeZipEntryPath -EntryName $entry.FullName
            $targetPath = [System.IO.Path]::GetFullPath((Join-Path $destinationRoot ($relativePath -replace '/', '\\')))
            if (-not $targetPath.StartsWith($destinationPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "ZIP 条目试图写入候选目录外：$relativePath"
            }
            if ([string]::IsNullOrEmpty($entry.Name)) {
                Ensure-SafeDirectory -Path $targetPath -Purpose "ZIP 目录条目 $relativePath" | Out-Null
                continue
            }
            $parentDirectory = Split-Path -Parent $targetPath
            Ensure-SafeDirectory -Path $parentDirectory -Purpose "ZIP 条目 $relativePath 的父目录" | Out-Null
            $targetPath = Assert-SafeNewFilePath -Path $targetPath -Purpose "ZIP 条目 $relativePath"

            $input = $null
            $output = $null
            try {
                $input = $entry.Open()
                $output = [System.IO.File]::Open($targetPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                $buffer = New-Object byte[] 81920
                [int64]$entryWritten = 0
                while (($read = $input.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $entryWritten += $read
                    $totalWritten += $read
                    if ($entryWritten -gt $MaximumSingleEntryBytes -or $totalWritten -gt $MaximumExpandedBytes) {
                        throw "ZIP 实际解压量超过安全限制：$relativePath"
                    }
                    $output.Write($buffer, 0, $read)
                }
                if ($entryWritten -ne $entry.Length) {
                    throw "ZIP 条目长度与中央目录声明不一致：$relativePath"
                }
            } finally {
                if ($null -ne $output) { $output.Dispose() }
                if ($null -ne $input) { $input.Dispose() }
            }
        }
    } finally {
        $archive.Dispose()
    }
}

function Assert-AllowedMarketplaceAndMcpConfiguration {
    <# 在执行包内脚本前以纯 JSON 数据检查唯一的本地 Marketplace 和固定 MCP 入口。 #>
    param([Parameter(Mandatory)][string]$Root)

    $Root = Assert-NoReparsePointsInDirectoryTree -Root $Root -Purpose 'Release Marketplace 目录'
    $marketplacePath = Join-Path $Root '.agents\plugins\marketplace.json'
    $mcpPath = Join-Path $Root "plugins\$PluginName\.mcp.json"
    $pluginManifestPath = Join-Path $Root "plugins\$PluginName\.codex-plugin\plugin.json"
    try {
        $marketplacePath = Assert-SafeExistingFilePath -Path $marketplacePath -Purpose 'Marketplace 配置'
        $mcpPath = Assert-SafeExistingFilePath -Path $mcpPath -Purpose 'MCP 配置'
        $pluginManifestPath = Assert-SafeExistingFilePath -Path $pluginManifestPath -Purpose '插件清单'
        $marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $mcp = Get-Content -LiteralPath $mcpPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $pluginManifest = Get-Content -LiteralPath $pluginManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "无法解析 Release 中的 Marketplace 或 MCP JSON：$($_.Exception.Message)"
    }

    Assert-ExactPropertyNames -Object $marketplace -AllowedNames @('name', 'interface', 'plugins') -Context 'Marketplace 清单'
    Assert-ExactPropertyNames -Object $marketplace.interface -AllowedNames @('displayName') -Context 'Marketplace interface'
    $entries = @($marketplace.plugins)
    if ($marketplace.name -ne $MarketplaceName -or $entries.Count -ne 1) {
        throw 'Release Marketplace 不在允许清单内。'
    }
    $entry = $entries[0]
    Assert-ExactPropertyNames -Object $entry -AllowedNames @('name', 'source', 'policy', 'category') -Context 'Marketplace 插件条目'
    Assert-ExactPropertyNames -Object $entry.source -AllowedNames @('source', 'path') -Context 'Marketplace source'
    Assert-ExactPropertyNames -Object $entry.policy -AllowedNames @('installation', 'authentication') -Context 'Marketplace policy'
    if ($entry.name -ne $PluginName -or $entry.source.source -ne 'local' -or $entry.source.path -ne "./plugins/$PluginName" -or $entry.policy.installation -ne 'AVAILABLE' -or $entry.policy.authentication -ne 'ON_INSTALL' -or [string]::IsNullOrWhiteSpace([string]$entry.category) -or [string]::IsNullOrWhiteSpace([string]$marketplace.interface.displayName)) {
        throw 'Release Marketplace 条目不在允许清单内。'
    }

    Assert-ExactPropertyNames -Object $mcp -AllowedNames @('mcpServers') -Context 'MCP 配置'
    Assert-ExactPropertyNames -Object $mcp.mcpServers -AllowedNames @($PluginName) -Context 'MCP server 清单'
    $server = $mcp.mcpServers.$PluginName
    Assert-ExactPropertyNames -Object $server -AllowedNames @('command', 'args', 'cwd') -Context 'MCP server'
    $expectedArguments = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', './scripts/paper-to-journal-club-server.ps1')
    $actualArguments = @($server.args | ForEach-Object { [string]$_ })
    if ($server.command -ne 'powershell.exe' -or $server.cwd -ne '.' -or $actualArguments.Count -ne $expectedArguments.Count) {
        throw 'Release MCP server 不在允许清单内。'
    }
    for ($index = 0; $index -lt $expectedArguments.Count; $index++) {
        if ($actualArguments[$index] -cne $expectedArguments[$index]) {
            throw 'Release MCP server 启动参数不在允许清单内。'
        }
    }

    # plugin.json 可声明技能与 MCP 配置，故也必须是固定的本地路径和字段集合。
    Assert-ExactPropertyNames -Object $pluginManifest -AllowedNames @('name', 'version', 'description', 'author', 'license', 'keywords', 'skills', 'mcpServers', 'interface') -Context '插件清单'
    Assert-ExactPropertyNames -Object $pluginManifest.author -AllowedNames @('name') -Context '插件作者'
    Assert-ExactPropertyNames -Object $pluginManifest.interface -AllowedNames @('displayName', 'shortDescription', 'longDescription', 'developerName', 'category', 'capabilities', 'defaultPrompt') -Context '插件 interface'
    if ($pluginManifest.name -ne $PluginName -or [string]::IsNullOrWhiteSpace([string]$pluginManifest.version) -or [string]::IsNullOrWhiteSpace([string]$pluginManifest.description) -or [string]::IsNullOrWhiteSpace([string]$pluginManifest.author.name) -or $pluginManifest.license -ne 'MIT' -or $pluginManifest.skills -ne './skills/' -or $pluginManifest.mcpServers -ne './.mcp.json' -or @($pluginManifest.keywords).Count -eq 0 -or @($pluginManifest.interface.capabilities).Count -eq 0 -or @($pluginManifest.interface.defaultPrompt).Count -eq 0) {
        throw 'Release 插件清单不在允许的本地技能/MCP 配置范围内。'
    }
}

function Assert-ReleasePackageRoot {
    <# 在执行包内安装器前，仅接受符合本插件 Marketplace 结构的解压目录。 #>
    param([Parameter(Mandatory)][string]$Root)

    $Root = Assert-NoReparsePointsInDirectoryTree -Root $Root -Purpose 'Release 候选目录'
    $marketplacePath = Join-Path $Root '.agents\plugins\marketplace.json'
    $pluginRoot = Join-Path $Root "plugins\$PluginName"
    $requiredPaths = @(
        $marketplacePath,
        (Join-Path $pluginRoot '.codex-plugin\plugin.json'),
        (Join-Path $pluginRoot '.mcp.json'),
        (Join-Path $pluginRoot 'scripts\paper-to-journal-club-server.ps1'),
        (Join-Path $pluginRoot 'assets\paper-parser.exe'),
        (Join-Path $Root 'install.ps1'),
        (Join-Path $Root 'verify-release.ps1'),
        (Join-Path $Root 'SHA256SUMS.txt')
    )
    foreach ($requiredPath in $requiredPaths) {
        try {
            $requiredPath = Assert-SafeExistingFilePath -Path $requiredPath -Purpose 'Release 必需文件'
        } catch {
            throw "Release ZIP 缺少必需文件：$requiredPath"
        }
    }

    Assert-AllowedMarketplaceAndMcpConfiguration -Root $Root
}

function Assert-ManagedMarketplaceDirectory {
    <# 覆盖旧版本前确认 current 是本安装器管理的目录，绝不替换任意用户文件夹。 #>
    param([Parameter(Mandatory)][string]$Root)

    $Root = Assert-NoReparsePointInExistingPath -Path $Root -Purpose '现有 current 目录'
    if (-not (Get-ExistingPathAttributes -Path $Root).Exists) {
        return
    }
    $Root = Assert-NoReparsePointsInDirectoryTree -Root $Root -Purpose '现有 current 目录'
    $marketplacePath = Join-Path $Root '.agents\plugins\marketplace.json'
    if (-not (Get-ExistingPathAttributes -Path $marketplacePath).Exists) {
        throw "拒绝替换未知目录：$Root"
    }
    Test-ChecksumManifest -Root $Root
    try {
        Assert-AllowedMarketplaceAndMcpConfiguration -Root $Root
    } catch {
        throw "拒绝替换不属于 $PluginName 的现有 Marketplace：$Root；$($_.Exception.Message)"
    }
}

function New-UniqueDirectory {
    param(
        [Parameter(Mandatory)][string]$Parent,
        [Parameter(Mandatory)][string]$Prefix
    )

    $path = New-UniqueChildPath -Parent $Parent -Prefix $Prefix -Purpose '候选目录'
    return (Ensure-SafeDirectory -Path $path -Purpose '候选目录')
}

function ConvertFrom-PackageInstallerResult {
    <#
      包内安装器会把所有过程信息写到 Host，最后只输出一条压缩 JSON。根安装器必须
      严格解析这条结果，而不能读取残留的 $LASTEXITCODE：CLI 自动安装失败后，包内
      安装器可能已安全回退到个人 Marketplace，但原生命令的退出码仍会保留为非零。
    #>
    param([Parameter(Mandatory)][object[]]$Output)

    $entries = @($Output | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($entries.Count -ne 1 -or $entries[0] -isnot [string]) {
        throw '包内安装器必须且只能输出一条机器可读的 JSON 最终结果。'
    }

    try {
        $result = $entries[0] | ConvertFrom-Json -ErrorAction Stop
    } catch {
        throw "包内安装器的最终结果不是有效 JSON：$($_.Exception.Message)"
    }

    foreach ($requiredProperty in @('pass', 'installation_mode', 'plugin_installed', 'codex_cli_called', 'next_step')) {
        if ($null -eq $result.PSObject.Properties[$requiredProperty]) {
            throw "包内安装器的最终结果缺少字段：$requiredProperty"
        }
    }
    if ($result.pass -ne $true -or [string]::IsNullOrWhiteSpace([string]$result.installation_mode) -or [string]::IsNullOrWhiteSpace([string]$result.next_step)) {
        throw '包内安装器的最终结果不表示成功完成安装或安全回退。'
    }
    if ($result.plugin_installed -isnot [bool] -or $result.codex_cli_called -isnot [bool]) {
        throw '包内安装器的最终结果中的安装状态字段必须为布尔值。'
    }
    return $result
}

Assert-ReleaseTag -Tag $ReleaseTag
$repository = Get-RepositoryInformation -Url $RepositoryUrl
if ([string]::IsNullOrWhiteSpace($InstallDirectory)) {
    $InstallDirectory = Get-DefaultInstallDirectory
}
$InstallDirectory = Assert-SafeInstallDirectory -Path $InstallDirectory

if ($WhatIfPreference) {
    # -WhatIf 不联网、不运行 CLI，也不创建目录；它如实说明真实安装会先验证受信 CLI，
    # 再自动安装或在失败时回退到个人 Marketplace。
    [pscustomobject]@{
        pass = $true
        action = 'would-download-verify-and-attempt-cli-install-or-personal-marketplace-fallback'
        repository = $repository.CanonicalUrl
        release_tag = $ReleaseTag
        install_directory = $InstallDirectory
        marketplace_root = (Join-Path $InstallDirectory 'current')
        installation_mode = 'trusted-cli-then-personal-marketplace-fallback'
        next_step = 'after-user-confirmation-verify-trusted-cli-then-auto-install-or-restart-codex-and-install-from-plugins-directory'
        plugin_installed = $false
        codex_cli_called = $false
        cli_temporary_data_possible = $true
        marketplace_registration_rollback_on_failure = $true
        node_required = $false
    } | ConvertTo-Json -Depth 4
    return
}

if ($env:OS -ne 'Windows_NT') {
    throw 'Paper to Journal Club 只能在 Windows 上安装，因为它通过 PowerPoint COM 控制桌面 Microsoft PowerPoint。'
}
if (-not [Environment]::Is64BitOperatingSystem) {
    throw '需要 64 位 Windows，才能运行随包提供的 PDF 解析器。'
}

$InstallDirectory = Ensure-SafeDirectory -Path $InstallDirectory -Purpose 'InstallDirectory'
$release = Get-GitHubRelease -Repository $repository -Tag $ReleaseTag
$archiveAsset = Get-ReleaseAsset -Release $release -RequestedName $ReleaseAssetName
$checksumAsset = Get-ChecksumAsset -Release $release -ArchiveAsset $archiveAsset

$downloadRoot = Join-Path $InstallDirectory (Join-Path 'downloads' $ReleaseTag)
$downloadRoot = Ensure-SafeDirectory -Path $downloadRoot -Purpose 'Release 下载缓存目录'
$checksumDownloadPath = New-UniqueChildPath -Parent $downloadRoot -Prefix 'checksum-' -Purpose 'SHA-256 下载文件'
$checksumDownloadPath = "$checksumDownloadPath.txt"
$checksumDownloadPath = Assert-SafeNewFilePath -Path $checksumDownloadPath -Purpose 'SHA-256 下载文件'
Save-GitHubAsset -Url $checksumAsset.browser_download_url -Destination $checksumDownloadPath
$expectedArchiveHash = Get-ExpectedArchiveHash -ChecksumPath $checksumDownloadPath -ArchiveName $archiveAsset.name

# 缓存名以内容哈希开头；同一标签被重新发布时不会悄悄覆盖已经校验过的旧下载。
$archiveCachePath = Join-Path $downloadRoot ("{0}-{1}" -f $expectedArchiveHash, $archiveAsset.name)
$archiveCachePath = Assert-NoReparsePointInExistingPath -Path $archiveCachePath -Purpose 'ZIP 缓存路径'
if ((Get-ExistingPathAttributes -Path $archiveCachePath).Exists) {
    $archiveCachePath = Assert-SafeExistingFilePath -Path $archiveCachePath -Purpose 'ZIP 缓存文件'
    Assert-ArchiveHash -Path $archiveCachePath -ExpectedHash $expectedArchiveHash
} else {
    $temporaryArchivePath = New-UniqueChildPath -Parent $downloadRoot -Prefix 'download-' -Purpose '临时 ZIP 下载文件'
    $temporaryArchivePath = "$temporaryArchivePath.tmp"
    $temporaryArchivePath = Assert-SafeNewFilePath -Path $temporaryArchivePath -Purpose '临时 ZIP 下载文件'
    Save-GitHubAsset -Url $archiveAsset.browser_download_url -Destination $temporaryArchivePath
    $temporaryArchivePath = Assert-SafeExistingFilePath -Path $temporaryArchivePath -Purpose '临时 ZIP 下载文件'
    if ((Get-Item -LiteralPath $temporaryArchivePath -ErrorAction Stop).Length -gt $MaximumReleaseBytes) {
        throw "下载的 ZIP 超过允许的最大大小 $MaximumReleaseBytes 字节。"
    }
    Assert-ArchiveHash -Path $temporaryArchivePath -ExpectedHash $expectedArchiveHash
    $archiveCachePath = Move-SafeInstallItem -Source $temporaryArchivePath -Destination $archiveCachePath -Purpose '缓存已校验 ZIP'
}
$archiveCachePath = Assert-SafeExistingFilePath -Path $archiveCachePath -Purpose 'ZIP 缓存文件'
if ((Get-Item -LiteralPath $archiveCachePath -ErrorAction Stop).Length -gt $MaximumReleaseBytes) {
    throw "缓存的 ZIP 超过允许的最大大小 $MaximumReleaseBytes 字节。"
}

# 解压到新的候选目录。先流式防护 ZIP 资源炸弹，再做完整哈希集合和固定入口检查，之后才替换 current。
$candidateDirectory = New-UniqueDirectory -Parent $InstallDirectory -Prefix '.incoming-'
try {
    Expand-SafeZipArchive -ArchivePath $archiveCachePath -DestinationPath $candidateDirectory
    $candidateDirectory = Assert-NoReparsePointsInDirectoryTree -Root $candidateDirectory -Purpose 'Release 候选目录'
    Test-ChecksumManifest -Root $candidateDirectory
    Assert-ReleasePackageRoot -Root $candidateDirectory
} catch {
    throw "Release ZIP 解压或结构检查失败：$($_.Exception.Message)"
}

$currentDirectory = Join-Path $InstallDirectory 'current'
$previousDirectory = $null
$activatedCandidate = $false
$packageInstallationResult = $null
try {
    $currentDirectory = Assert-NoReparsePointInExistingPath -Path $currentDirectory -Purpose 'current 安装目录'
    if ((Get-ExistingPathAttributes -Path $currentDirectory).Exists) {
        Assert-ManagedMarketplaceDirectory -Root $currentDirectory
        $previousDirectory = New-UniqueChildPath -Parent $InstallDirectory -Prefix ("previous-{0}-" -f (Get-Date -Format 'yyyyMMdd-HHmmss')) -Purpose 'previous 回滚目录'
        $previousDirectory = Move-SafeInstallItem -Source $currentDirectory -Destination $previousDirectory -Purpose '备份旧 current 目录'
    }

    $candidateDirectory = Assert-NoReparsePointsInDirectoryTree -Root $candidateDirectory -Purpose '待激活候选目录'
    $currentDirectory = Move-SafeInstallItem -Source $candidateDirectory -Destination $currentDirectory -Purpose '激活候选目录'
    $activatedCandidate = $true

    # 包内安装器负责检查包内 SHA256SUMS、PowerPoint 和 MCP 入口；在用户确认后验证受信 CLI，
    # 优先自动安装，失败时回滚本次注册并回退个人 Marketplace，避免两套规则漂移。
    Test-ChecksumManifest -Root $currentDirectory
    Assert-ReleasePackageRoot -Root $currentDirectory
    $packageInstaller = Assert-SafeExistingFilePath -Path (Join-Path $currentDirectory 'install.ps1') -Purpose '包内安装器'
    $installerParameters = @{ MarketplaceRoot = $currentDirectory }
    if ($SkipPowerPointCheck) {
        $installerParameters.SkipPowerPointCheck = $true
    }
    if ($SkipPersonalMarketplaceDeployment) {
        $installerParameters.SkipPersonalMarketplaceDeployment = $true
    }
    # 包内安装器会捕获 CLI 的非零退出码并自行回退；因此绝不能读取此处残留的
    # $LASTEXITCODE 判断成失败。最终状态只以其唯一机器可读结果或终止异常为准。
    $packageInstallerOutput = @(& $packageInstaller @installerParameters)
    $packageInstallationResult = ConvertFrom-PackageInstallerResult -Output $packageInstallerOutput
} catch {
    $installationError = $_
    # 失败时保留候选包供排障，并将已知的旧版本恢复到固定 current 路径。
    if ($activatedCandidate) {
        try {
            $currentDirectory = Assert-NoReparsePointInExistingPath -Path $currentDirectory -Purpose '失败 current 目录' -RequireExisting
            if ([System.IO.Directory]::Exists($currentDirectory)) {
                $failedDirectory = New-UniqueChildPath -Parent $InstallDirectory -Prefix ("failed-{0}-" -f (Get-Date -Format 'yyyyMMdd-HHmmss')) -Purpose 'failed 排障目录'
                Move-SafeInstallItem -Source $currentDirectory -Destination $failedDirectory -Purpose '隔离失败 current 目录' | Out-Null
            }
        } catch {
            Write-Warning "未移动失败候选目录，避免在不安全路径上操作：$($_.Exception.Message)"
        }
    }
    if ($null -ne $previousDirectory) {
        try {
            $previousDirectory = Assert-NoReparsePointsInDirectoryTree -Root $previousDirectory -Purpose 'previous 回滚目录'
            $currentDirectory = Assert-NoReparsePointInExistingPath -Path $currentDirectory -Purpose '回滚 current 目录'
            if ([System.IO.Directory]::Exists($previousDirectory) -and -not (Get-ExistingPathAttributes -Path $currentDirectory).Exists) {
                $currentDirectory = Move-SafeInstallItem -Source $previousDirectory -Destination $currentDirectory -Purpose '恢复 previous 目录'
            }
        } catch {
            Write-Warning "未执行 previous 回滚，避免在不安全路径上操作：$($_.Exception.Message)"
        }
    }
    throw $installationError
}

if ($null -eq $packageInstallationResult) {
    throw '包内安装器未返回可用的最终安装结果。'
}

[pscustomobject]@{
    pass = $true
    plugin = "$PluginName@$MarketplaceName"
    repository = $repository.CanonicalUrl
    release_tag = $ReleaseTag
    archive = $archiveAsset.name
    archive_sha256 = $expectedArchiveHash
    marketplace_root = $currentDirectory
    previous_marketplace_root = $previousDirectory
    installation_mode = $packageInstallationResult.installation_mode
    next_step = $packageInstallationResult.next_step
    plugin_installed = $packageInstallationResult.plugin_installed
    codex_cli_called = $packageInstallationResult.codex_cli_called
    package_marketplace_path = if ($null -ne $packageInstallationResult.PSObject.Properties['marketplace_path']) { $packageInstallationResult.marketplace_path } else { $null }
    package_detail = if ($null -ne $packageInstallationResult.PSObject.Properties['detail']) { $packageInstallationResult.detail } else { $null }
    node_required = $false
    personal_marketplace_deployment_skipped = [bool]$SkipPersonalMarketplaceDeployment
} | ConvertTo-Json -Depth 4
