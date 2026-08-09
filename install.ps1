<#
  Paper to Journal Club 的 GitHub Release 自助安装器。

  此脚本不克隆开发仓库，也不会要求终端用户安装 Git、Node.js、npm、Python 或 .NET SDK。
  它从指定 GitHub 仓库的一个固定 Release 下载已经构建好的本地 Marketplace ZIP，先校验
  同一 Release 中的 SHA-256 文件，再把通过校验的包安装到稳定的本地目录并调用包内安装器。

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

    # 只下载、校验并部署 Marketplace 文件，不执行 Codex Marketplace 注册与插件安装。
    [switch]$SkipCodexInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'
$MaximumReleaseBytes = 512MB

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
    <# 禁止把安装根目录设置为驱动器根目录，后续所有写入只发生在该目录的专用子目录。 #>
    param([Parameter(Mandatory)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $trimmedPath = $fullPath.TrimEnd([char]'\', [char]'/')
    $trimmedRoot = [System.IO.Path]::GetPathRoot($fullPath).TrimEnd([char]'\', [char]'/')
    if ([string]::Equals($trimmedPath, $trimmedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'InstallDirectory 不能是驱动器根目录。请选择一个专用的用户目录。'
    }
    return $fullPath
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

    if ($null -eq $release -or $release.draft) {
        throw "GitHub Release '$Tag' 不可用或仍是草稿。"
    }
    return $release
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

    if ([string]::IsNullOrWhiteSpace([string]$asset.browser_download_url) -or -not ([string]$asset.browser_download_url).StartsWith('https://', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'GitHub Release 返回的 ZIP 下载地址无效。'
    }
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
    if ([string]::IsNullOrWhiteSpace([string]$matches[0].browser_download_url) -or -not ([string]$matches[0].browser_download_url).StartsWith('https://', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'GitHub Release 返回的 SHA-256 下载地址无效。'
    }
    return $matches[0]
}

function Save-GitHubAsset {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$Destination
    )

    $headers = @{
        'Accept' = 'application/octet-stream'
        'User-Agent' = 'paper-to-journal-club-installer'
    }
    try {
        Invoke-WebRequest -Uri $Url -Headers $headers -OutFile $Destination -UseBasicParsing
    } catch {
        throw "下载 Release 资产失败：$($_.Exception.Message)"
    }
    if (-not (Test-Path -LiteralPath $Destination -PathType Leaf) -or (Get-Item -LiteralPath $Destination).Length -eq 0) {
        throw '下载的 Release 资产为空或不存在。'
    }
}

function Get-ExpectedArchiveHash {
    param(
        [Parameter(Mandatory)][string]$ChecksumPath,
        [Parameter(Mandatory)][string]$ArchiveName
    )

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

    $actualHash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actualHash -ne $ExpectedHash) {
        throw "ZIP 的 SHA-256 校验失败。期望：$ExpectedHash；实际：$actualHash"
    }
}

function Assert-ReleasePackageRoot {
    <# 在执行包内安装器前，仅接受符合本插件 Marketplace 结构的解压目录。 #>
    param([Parameter(Mandatory)][string]$Root)

    $marketplacePath = Join-Path $Root '.agents\plugins\marketplace.json'
    $pluginRoot = Join-Path $Root "plugins\$PluginName"
    $requiredPaths = @(
        $marketplacePath,
        (Join-Path $pluginRoot '.codex-plugin\plugin.json'),
        (Join-Path $pluginRoot '.mcp.json'),
        (Join-Path $pluginRoot 'assets\paper-parser.exe'),
        (Join-Path $Root 'install.ps1'),
        (Join-Path $Root 'verify-release.ps1'),
        (Join-Path $Root 'SHA256SUMS.txt')
    )
    foreach ($requiredPath in $requiredPaths) {
        if (-not (Test-Path -LiteralPath $requiredPath)) {
            throw "Release ZIP 缺少必需文件：$requiredPath"
        }
    }

    try {
        $marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "无法解析 Release 中的 Marketplace 清单：$($_.Exception.Message)"
    }
    if ($marketplace.name -ne $MarketplaceName) {
        throw "Release Marketplace 名称不正确：$($marketplace.name)"
    }
    $entry = @($marketplace.plugins | Where-Object { $_.name -eq $PluginName })
    if ($entry.Count -ne 1 -or $entry[0].source.source -ne 'local' -or $entry[0].source.path -ne "./plugins/$PluginName") {
        throw 'Release Marketplace 清单没有指向预期的本地插件路径。'
    }
}

function Assert-ManagedMarketplaceDirectory {
    <# 覆盖旧版本前确认 current 是本安装器管理的目录，绝不替换任意用户文件夹。 #>
    param([Parameter(Mandatory)][string]$Root)

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        return
    }
    $marketplacePath = Join-Path $Root '.agents\plugins\marketplace.json'
    if (-not (Test-Path -LiteralPath $marketplacePath -PathType Leaf)) {
        throw "拒绝替换未知目录：$Root"
    }
    try {
        $marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "拒绝替换无法解析的现有 Marketplace：$Root"
    }
    $entry = @($marketplace.plugins | Where-Object { $_.name -eq $PluginName })
    if ($marketplace.name -ne $MarketplaceName -or $entry.Count -ne 1 -or $entry[0].source.path -ne "./plugins/$PluginName") {
        throw "拒绝替换不属于 $PluginName 的现有 Marketplace：$Root"
    }
}

function New-UniqueDirectory {
    param(
        [Parameter(Mandatory)][string]$Parent,
        [Parameter(Mandatory)][string]$Prefix
    )

    $path = Join-Path $Parent ("{0}{1}" -f $Prefix, [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $path -ErrorAction Stop | Out-Null
    return $path
}

Assert-ReleaseTag -Tag $ReleaseTag
$repository = Get-RepositoryInformation -Url $RepositoryUrl
if ([string]::IsNullOrWhiteSpace($InstallDirectory)) {
    $InstallDirectory = Get-DefaultInstallDirectory
}
$InstallDirectory = Assert-SafeInstallDirectory -Path $InstallDirectory

if ($WhatIfPreference) {
    # -WhatIf 只展示将要发生的网络和文件操作，便于 CI 在不联网的条件下检查参数与布局。
    [pscustomobject]@{
        pass = $true
        action = 'would-download-and-install-release'
        repository = $repository.CanonicalUrl
        release_tag = $ReleaseTag
        install_directory = $InstallDirectory
        marketplace_root = (Join-Path $InstallDirectory 'current')
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

New-Item -ItemType Directory -Force -Path $InstallDirectory | Out-Null
$release = Get-GitHubRelease -Repository $repository -Tag $ReleaseTag
$archiveAsset = Get-ReleaseAsset -Release $release -RequestedName $ReleaseAssetName
$checksumAsset = Get-ChecksumAsset -Release $release -ArchiveAsset $archiveAsset

$downloadRoot = Join-Path $InstallDirectory (Join-Path 'downloads' $ReleaseTag)
New-Item -ItemType Directory -Force -Path $downloadRoot | Out-Null
$checksumDownloadPath = Join-Path $downloadRoot ("checksum-{0}.txt" -f [Guid]::NewGuid().ToString('N'))
Save-GitHubAsset -Url $checksumAsset.browser_download_url -Destination $checksumDownloadPath
$expectedArchiveHash = Get-ExpectedArchiveHash -ChecksumPath $checksumDownloadPath -ArchiveName $archiveAsset.name

# 缓存名以内容哈希开头；同一标签被重新发布时不会悄悄覆盖已经校验过的旧下载。
$archiveCachePath = Join-Path $downloadRoot ("{0}-{1}" -f $expectedArchiveHash, $archiveAsset.name)
if (Test-Path -LiteralPath $archiveCachePath -PathType Leaf) {
    Assert-ArchiveHash -Path $archiveCachePath -ExpectedHash $expectedArchiveHash
} else {
    $temporaryArchivePath = Join-Path $downloadRoot ("download-{0}.tmp" -f [Guid]::NewGuid().ToString('N'))
    Save-GitHubAsset -Url $archiveAsset.browser_download_url -Destination $temporaryArchivePath
    if ((Get-Item -LiteralPath $temporaryArchivePath).Length -gt $MaximumReleaseBytes) {
        throw "下载的 ZIP 超过允许的最大大小 $MaximumReleaseBytes 字节。"
    }
    Assert-ArchiveHash -Path $temporaryArchivePath -ExpectedHash $expectedArchiveHash
    Move-Item -LiteralPath $temporaryArchivePath -Destination $archiveCachePath -ErrorAction Stop
}

# 解压到新的候选目录，只有结构和哈希都通过包内校验后才会替换 current。
$candidateDirectory = New-UniqueDirectory -Parent $InstallDirectory -Prefix '.incoming-'
try {
    Expand-Archive -LiteralPath $archiveCachePath -DestinationPath $candidateDirectory -ErrorAction Stop
    Assert-ReleasePackageRoot -Root $candidateDirectory
} catch {
    throw "Release ZIP 解压或结构检查失败：$($_.Exception.Message)"
}

$currentDirectory = Join-Path $InstallDirectory 'current'
$previousDirectory = $null
$activatedCandidate = $false
try {
    if (Test-Path -LiteralPath $currentDirectory -PathType Container) {
        Assert-ManagedMarketplaceDirectory -Root $currentDirectory
        $previousDirectory = Join-Path $InstallDirectory ("previous-{0}-{1}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [Guid]::NewGuid().ToString('N'))
        Move-Item -LiteralPath $currentDirectory -Destination $previousDirectory -ErrorAction Stop
    }

    Move-Item -LiteralPath $candidateDirectory -Destination $currentDirectory -ErrorAction Stop
    $activatedCandidate = $true

    # 包内安装器负责检查包内 SHA256SUMS、PowerPoint、Codex CLI 和 MCP 入口，避免两套规则漂移。
    $packageInstaller = Join-Path $currentDirectory 'install.ps1'
    $installerParameters = @{ MarketplaceRoot = $currentDirectory }
    if ($SkipPowerPointCheck) {
        $installerParameters.SkipPowerPointCheck = $true
    }
    if ($SkipCodexInstall) {
        $installerParameters.SkipCodexInstall = $true
    }
    & $packageInstaller @installerParameters
    if ($null -ne $LASTEXITCODE -and $LASTEXITCODE -ne 0) {
        throw "包内安装器以退出码 $LASTEXITCODE 结束。"
    }
} catch {
    $installationError = $_
    # 失败时保留候选包供排障，并将已知的旧版本恢复到固定 current 路径。
    if ($activatedCandidate -and (Test-Path -LiteralPath $currentDirectory -PathType Container)) {
        $failedDirectory = Join-Path $InstallDirectory ("failed-{0}-{1}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [Guid]::NewGuid().ToString('N'))
        Move-Item -LiteralPath $currentDirectory -Destination $failedDirectory -ErrorAction SilentlyContinue
    }
    if ($null -ne $previousDirectory -and (Test-Path -LiteralPath $previousDirectory -PathType Container) -and -not (Test-Path -LiteralPath $currentDirectory)) {
        Move-Item -LiteralPath $previousDirectory -Destination $currentDirectory -ErrorAction SilentlyContinue
    }
    throw $installationError
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
    node_required = $false
    codex_install_skipped = [bool]$SkipCodexInstall
} | ConvertTo-Json -Depth 4
