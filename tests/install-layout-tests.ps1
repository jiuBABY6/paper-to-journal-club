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
$marketplacePath = Join-Path $RepositoryRoot '.agents\plugins\marketplace.json'
$pluginRoot = Join-Path $RepositoryRoot "plugins\$pluginName"
$releaseToolPaths = @(
    (Join-Path $RepositoryRoot 'tools\Build-ReleasePackage.ps1'),
    (Join-Path $RepositoryRoot 'tools\Install-PaperToJournalClub.ps1'),
    (Join-Path $RepositoryRoot 'tools\Test-ReleasePackage.ps1'),
    (Join-Path $RepositoryRoot 'tools\INSTALL.md')
)

foreach ($path in @($installPowerShellPath, $publishingPath, $workflowPath, $marketplacePath, (Join-Path $pluginRoot '.codex-plugin\plugin.json'), (Join-Path $pluginRoot '.mcp.json')) + $releaseToolPaths) {
    Assert-FileExists -Path $path
}

# Windows PowerShell 5.1 需要 UTF-8 BOM 才能可靠读取中文字符串和注释。
$installBytes = [System.IO.File]::ReadAllBytes($installPowerShellPath)
Assert-True -Condition ($installBytes.Length -ge 3 -and $installBytes[0] -eq 0xEF -and $installBytes[1] -eq 0xBB -and $installBytes[2] -eq 0xBF) -Message 'install.ps1 必须采用 UTF-8 BOM 编码。'

$tokens = $null
$parseErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($installPowerShellPath, [ref]$tokens, [ref]$parseErrors)
$parseErrorMessages = @($parseErrors | ForEach-Object { $_.Message })
Assert-True -Condition ($parseErrorMessages.Count -eq 0) -Message ("install.ps1 存在 PowerShell 语法错误：{0}" -f ($parseErrorMessages -join '; '))

$marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-True -Condition ($marketplace.name -eq $marketplaceName) -Message 'Marketplace 名称不正确。'
$entries = @($marketplace.plugins | Where-Object { $_.name -eq $pluginName })
Assert-True -Condition ($entries.Count -eq 1) -Message 'Marketplace 必须恰好包含一个目标插件条目。'
Assert-True -Condition ($entries[0].source.source -eq 'local' -and $entries[0].source.path -eq "./plugins/$pluginName") -Message 'Marketplace 必须指向插件的相对本地路径。'

$installSource = [System.IO.File]::ReadAllText($installPowerShellPath, [System.Text.UTF8Encoding]::new($true))
Assert-True -Condition ($installSource -match 'RepositoryUrl' -and $installSource -match 'ReleaseTag' -and $installSource -match 'SupportsShouldProcess') -Message 'install.ps1 必须要求仓库地址和固定 Release 标签，并支持 -WhatIf。'
Assert-True -Condition ($installSource -notmatch '(?i)ExecutionPolicy\s+Bypass') -Message 'install.ps1 不得引入 ExecutionPolicy Bypass。'

$publishingSource = Get-Content -LiteralPath $publishingPath -Raw -Encoding UTF8
Assert-True -Condition ($publishingSource -match '\.zip\.sha256' -and $publishingSource -match 'bootstrap\.ps1' -and $publishingSource -match 'RemoteSigned') -Message 'PUBLISHING.md 必须说明外部 SHA-256、引导安装器与 RemoteSigned。'

$workflowSource = Get-Content -LiteralPath $workflowPath -Raw -Encoding UTF8
foreach ($requiredFragment in @('actions/checkout@v4', 'actions/setup-dotnet@v4', 'Get-FileHash', 'bootstrap.ps1', 'gh release')) {
    Assert-True -Condition ($workflowSource.Contains($requiredFragment)) -Message "release.yml 缺少预期发布步骤：$requiredFragment"
}
Assert-True -Condition ($workflowSource.Contains('tools/Build-ReleasePackage.ps1')) -Message 'release.yml 必须调用 tools/Build-ReleasePackage.ps1。'
Assert-True -Condition (-not $workflowSource.Contains('release/Build-ReleasePackage.ps1')) -Message 'release.yml 不得调用迁移前的 release/Build-ReleasePackage.ps1。'

# -WhatIf 必须在请求网络或创建目录前退出，因而可用于 CI 中的安全预检。
$unusedDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("paper-to-journal-club-install-test-{0}" -f [Guid]::NewGuid().ToString('N'))
$whatIfOutput = & $installPowerShellPath -RepositoryUrl 'https://github.com/test-owner/test-repository.git' -ReleaseTag 'v1.0.0' -InstallDirectory $unusedDirectory -WhatIf
$whatIf = $whatIfOutput | ConvertFrom-Json
Assert-True -Condition ($whatIf.pass -eq $true -and $whatIf.action -eq 'would-download-and-install-release' -and $whatIf.node_required -eq $false) -Message 'install.ps1 的 -WhatIf 输出不符合预期。'
Assert-True -Condition (-not (Test-Path -LiteralPath $unusedDirectory)) -Message '-WhatIf 不得创建安装目录。'

Write-Host 'PASS: install-layout-tests.ps1'
