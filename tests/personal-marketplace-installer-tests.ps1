<#
  个人 Marketplace 安装器回归测试。

  不读取或写入真实用户的 ~/.codex、~/.agents，也不启动 Codex 或 PowerPoint。
  测试会从包内安装器提取辅助函数，在临时用户目录中验证复制、配置合并、备份和
  Windows PowerShell 5.1 兼容性。
#>
[CmdletBinding()]
param(
    [string]$PackageInstallerPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )

    if (-not $Condition) {
        throw "ASSERTION FAILED: $Message"
    }
}

if ([string]::IsNullOrWhiteSpace($PackageInstallerPath)) {
    $PackageInstallerPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'tools\Install-PaperToJournalClub.ps1'
}
$PackageInstallerPath = (Resolve-Path -LiteralPath $PackageInstallerPath -ErrorAction Stop).Path

# 只提取 function 定义，避免执行安装器末尾的真实发行包校验和用户目录部署。
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($PackageInstallerPath, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) {
    throw "包内安装器存在语法错误：$(@($parseErrors | ForEach-Object Message) -join '; ')"
}
$functionDefinitions = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))
if ($functionDefinitions.Count -eq 0) {
    throw '未能从包内安装器提取任何辅助函数。'
}

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'
$helperSource = ($functionDefinitions | ForEach-Object { $_.Extent.Text }) -join [Environment]::NewLine
. ([scriptblock]::Create($helperSource))

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("paper-to-journal-club-personal-marketplace-test-{0}" -f [Guid]::NewGuid().ToString('N'))
$temporaryParent = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd([char]'\', [char]'/')
$testRoot = [System.IO.Path]::GetFullPath($testRoot)

try {
    [System.IO.Directory]::CreateDirectory($testRoot) | Out-Null

    # 验证完整性函数会返回哈希表，供后续复制阶段继续约束源文件，而不是只做一次性校验。
    $checksumTestRoot = Join-Path $testRoot 'checksum-release'
    [System.IO.Directory]::CreateDirectory($checksumTestRoot) | Out-Null
    $payloadPath = Join-Path $checksumTestRoot 'payload.txt'
    [System.IO.File]::WriteAllText($payloadPath, 'verified payload', [System.Text.UTF8Encoding]::new($false))
    $payloadHash = (Get-FileHash -LiteralPath $payloadPath -Algorithm SHA256).Hash.ToUpperInvariant()
    [System.IO.File]::WriteAllText(
        (Join-Path $checksumTestRoot 'SHA256SUMS.txt'),
        "$payloadHash  *payload.txt",
        [System.Text.UTF8Encoding]::new($false)
    )
    $verifiedChecksums = Test-Checksums -Root $checksumTestRoot
    Assert-True -Condition ($verifiedChecksums -is [hashtable] -and $verifiedChecksums['payload.txt'] -eq $payloadHash) -Message '完整性检查必须返回经验证的 SHA-256 映射。'

    # 构造最小可复制插件树，并由安装器自身计算期望哈希表。
    $sourcePluginRoot = Join-Path $testRoot 'verified-release-plugin'
    [System.IO.Directory]::CreateDirectory((Join-Path $sourcePluginRoot '.codex-plugin')) | Out-Null
    [System.IO.Directory]::CreateDirectory((Join-Path $sourcePluginRoot 'skills\demo')) | Out-Null
    [System.IO.File]::WriteAllText(
        (Join-Path $sourcePluginRoot '.codex-plugin\plugin.json'),
        '{"name":"paper-to-journal-club","version":"test"}',
        [System.Text.UTF8Encoding]::new($false)
    )
    [System.IO.File]::WriteAllText(
        (Join-Path $sourcePluginRoot 'skills\demo\SKILL.md'),
        '# test',
        [System.Text.UTF8Encoding]::new($false)
    )
    $expectedPluginChecksums = Get-DirectoryDigestMap -Root $sourcePluginRoot -Purpose '测试发行插件目录'

    # 清单确认后若源文件被替换，复制函数必须拒绝它，不能只比较“源和副本彼此相同”。
    $skillPath = Join-Path $sourcePluginRoot 'skills\demo\SKILL.md'
    [System.IO.File]::WriteAllText($skillPath, '# tampered', [System.Text.UTF8Encoding]::new($false))
    $tamperError = $null
    try {
        Deploy-VerifiedPluginCopy -SourcePluginRoot $sourcePluginRoot -UserProfilePath (Join-Path $testRoot 'tamper-profile') -ExpectedPluginChecksums $expectedPluginChecksums | Out-Null
    } catch {
        $tamperError = $_
    }
    Assert-True -Condition ($null -ne $tamperError -and $tamperError.Exception.Message -match 'SHA-256') -Message '部署必须拒绝校验后被篡改的插件源文件。'
    [System.IO.File]::WriteAllText($skillPath, '# test', [System.Text.UTF8Encoding]::new($false))

    $testProfile = Join-Path $testRoot 'test-user-profile'
    [System.IO.Directory]::CreateDirectory($testProfile) | Out-Null
    $deployedPluginPath = Deploy-VerifiedPluginCopy -SourcePluginRoot $sourcePluginRoot -UserProfilePath $testProfile -ExpectedPluginChecksums $expectedPluginChecksums
    Assert-True -Condition ([System.IO.Directory]::Exists($deployedPluginPath)) -Message '没有部署个人插件目录。'
    Assert-DirectoryCopyMatches -Source $sourcePluginRoot -Destination $deployedPluginPath -ExpectedSourceHashes $expectedPluginChecksums

    # 模拟进程刚完成 old → previous 移动就中断的场景；下一次部署必须先恢复旧副本，
    # 而不能静默丢失插件目录或留下多个不可见版本。
    $pluginsDirectory = Split-Path -Parent $deployedPluginPath
    $interruptedPrevious = Join-Path $pluginsDirectory (".$PluginName.previous-{0}" -f [Guid]::NewGuid().ToString('N'))
    [System.IO.Directory]::Move($deployedPluginPath, $interruptedPrevious)
    $deployedPluginPath = Deploy-VerifiedPluginCopy -SourcePluginRoot $sourcePluginRoot -UserProfilePath $testProfile -ExpectedPluginChecksums $expectedPluginChecksums
    Assert-True -Condition ([System.IO.Directory]::Exists($deployedPluginPath) -and -not [System.IO.Directory]::Exists($interruptedPrevious)) -Message '中断后的个人插件旧副本没有被安全恢复。'

    # 预置另一个插件，验证更新时不会覆盖用户已有的 Marketplace 条目。
    $agentsPluginsDirectory = Join-Path $testProfile '.agents\plugins'
    [System.IO.Directory]::CreateDirectory($agentsPluginsDirectory) | Out-Null
    $marketplacePath = Join-Path $agentsPluginsDirectory 'marketplace.json'
    $oldMarketplace = [pscustomobject]@{
        name = 'personal'
        interface = [pscustomobject]@{ displayName = 'Personal' }
        plugins = @(
            [pscustomobject]@{
                name = 'existing-plugin'
                source = [pscustomobject]@{ source = 'local'; path = './plugins/existing-plugin' }
                policy = [pscustomobject]@{ installation = 'AVAILABLE'; authentication = 'ON_INSTALL' }
                category = 'Productivity'
            }
        )
    } | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($marketplacePath, $oldMarketplace, [System.Text.UTF8Encoding]::new($false))

    # 现有配置的 plugins 必须是 JSON 数组，不能把对象误当数组重写。
    $malformedProfile = Join-Path $testRoot 'malformed-profile'
    $malformedMarketplaceDirectory = Join-Path $malformedProfile '.agents\plugins'
    [System.IO.Directory]::CreateDirectory($malformedMarketplaceDirectory) | Out-Null
    [System.IO.File]::WriteAllText(
        (Join-Path $malformedMarketplaceDirectory 'marketplace.json'),
        '{"name":"personal","interface":{"displayName":"Personal"},"plugins":{"name":"not-an-array"}}',
        [System.Text.UTF8Encoding]::new($false)
    )
    $malformedError = $null
    try {
        Update-PersonalMarketplaceFile -UserProfilePath $malformedProfile | Out-Null
    } catch {
        $malformedError = $_
    }
    Assert-True -Condition ($null -ne $malformedError -and $malformedError.Exception.Message -match 'JSON 数组') -Message '个人 Marketplace 的 plugins 不是数组时必须停止，不能覆盖用户配置。'

    $firstUpdate = Update-PersonalMarketplaceFile -UserProfilePath $testProfile
    Assert-True -Condition ($null -ne $firstUpdate.BackupPath -and [System.IO.File]::Exists($firstUpdate.BackupPath)) -Message '更新已有 Marketplace 时应保留原配置备份。'
    $updatedMarketplace = Get-Content -LiteralPath $firstUpdate.Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $targetEntries = @($updatedMarketplace.plugins | Where-Object { $_.name -eq $PluginName })
    $existingEntries = @($updatedMarketplace.plugins | Where-Object { $_.name -eq 'existing-plugin' })
    Assert-True -Condition ($targetEntries.Count -eq 1 -and $existingEntries.Count -eq 1) -Message '个人 Marketplace 合并后必须同时保留已有条目和唯一目标条目。'
    Assert-PersonalMarketplaceEntry -Entry $targetEntries[0]

    # 第二次更新验证 Generic List 已被明确转成 object[]，不会在 Windows PowerShell 5.1 失败或重复条目。
    $secondUpdate = Update-PersonalMarketplaceFile -UserProfilePath $testProfile
    $twiceUpdatedMarketplace = Get-Content -LiteralPath $secondUpdate.Path -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True -Condition (@($twiceUpdatedMarketplace.plugins | Where-Object { $_.name -eq $PluginName }).Count -eq 1) -Message '重复部署不能产生多个 Paper to Journal Club 条目。'
    Assert-True -Condition (@($twiceUpdatedMarketplace.plugins | Where-Object { $_.name -eq 'existing-plugin' }).Count -eq 1) -Message '重复部署不能删除已有 Marketplace 条目。'

    # CLI 事务 fixture：不启动真实 codex.exe。add 一旦被尝试，安装器只能在明确
    # 没有残留，或精确 root 的本次条目成功 remove 且最终列表为空时才允许回退。
    function New-MockMarketplaceInventoryJson {
        param([Parameter(Mandatory)][string]$State)

        switch ($State) {
            'none' {
                return ([pscustomobject]@{ marketplaces = [System.Array]::CreateInstance([object], 0) } | ConvertTo-Json -Compress)
            }
            'expected-root' {
                return ([pscustomobject]@{
                        marketplaces = [object[]]@([pscustomobject]@{ name = $MarketplaceName; root = 'D:\verified-release' })
                    } | ConvertTo-Json -Compress)
            }
            'mismatched-root' {
                return ([pscustomobject]@{
                        marketplaces = [object[]]@([pscustomobject]@{ name = $MarketplaceName; root = 'D:\another-user-marketplace' })
                    } | ConvertTo-Json -Compress)
            }
            'invalid' {
                # JSON 可解析但 marketplaces 不是数组，必须被安装器拒绝。
                return '{"marketplaces":{"name":"not-an-array"}}'
            }
            default {
                throw "未知 CLI inventory fixture：$State"
            }
        }
    }

    $script:mockCliCalls = New-Object 'System.Collections.Generic.List[string]'
    $script:mockInventorySteps = New-Object 'System.Collections.Generic.Queue[string]'
    $script:mockAddSucceeded = $true
    $script:mockPluginAddSucceeded = $true
    $script:mockRemoveSucceeded = $true
    function Invoke-CodexCliCommand {
        param(
            [Parameter(Mandatory)][string]$ExecutablePath,
            [Parameter(Mandatory)][string[]]$Arguments
        )

        $commandText = $Arguments -join ' '
        [void]$script:mockCliCalls.Add($commandText)
        if ($commandText -eq 'plugin marketplace list --json') {
            if ($script:mockInventorySteps.Count -eq 0) {
                throw 'CLI inventory fixture 已耗尽。'
            }
            $state = $script:mockInventorySteps.Dequeue()
            return [pscustomobject]@{
                Succeeded = $true
                ExitCode = 0
                Summary = ''
                Output = New-MockMarketplaceInventoryJson -State $state
                ErrorOutput = ''
            }
        }
        if ($commandText -eq 'plugin marketplace add D:\verified-release') {
            return [pscustomobject]@{
                Succeeded = [bool]$script:mockAddSucceeded
                ExitCode = if ($script:mockAddSucceeded) { 0 } else { 1 }
                Summary = 'mock marketplace add result'
                Output = ''
                ErrorOutput = ''
            }
        }
        if ($commandText -eq "plugin add $PluginName@$MarketplaceName") {
            return [pscustomobject]@{
                Succeeded = [bool]$script:mockPluginAddSucceeded
                ExitCode = if ($script:mockPluginAddSucceeded) { 0 } else { 1 }
                Summary = 'mock plugin add result'
                Output = ''
                ErrorOutput = ''
            }
        }
        if ($commandText -eq "plugin marketplace remove $MarketplaceName") {
            return [pscustomobject]@{
                Succeeded = [bool]$script:mockRemoveSucceeded
                ExitCode = if ($script:mockRemoveSucceeded) { 0 } else { 1 }
                Summary = 'mock marketplace remove result'
                Output = ''
                ErrorOutput = ''
            }
        }
        throw "意外的 CLI fixture 调用：$commandText"
    }

    # add 成功但第二次 list 无法验证：不得个人回退，也不得猜测 remove。
    $script:mockCliCalls.Clear(); $script:mockInventorySteps.Clear()
    foreach ($state in @('none', 'invalid')) { $script:mockInventorySteps.Enqueue($state) }
    $script:mockAddSucceeded = $true; $script:mockPluginAddSucceeded = $true; $script:mockRemoveSucceeded = $true
    $uncertainAfterAdd = $null
    try {
        Try-InstallWithCodexCli -ExecutablePath 'C:\trusted\codex.exe' -VerifiedMarketplaceRoot 'D:\verified-release' | Out-Null
    } catch {
        $uncertainAfterAdd = $_
    }
    Assert-True -Condition ($null -ne $uncertainAfterAdd -and $script:mockCliCalls -notcontains "plugin marketplace remove $MarketplaceName") -Message 'add 后 inventory 无法验证时必须阻止回退且不得 remove 未确认条目。'

    # add 返回非零也可能已经写入；若最终目标仍存在且 remove 后仍存在，必须 fail-closed。
    $script:mockCliCalls.Clear(); $script:mockInventorySteps.Clear()
    foreach ($state in @('none', 'expected-root', 'expected-root', 'expected-root')) { $script:mockInventorySteps.Enqueue($state) }
    $script:mockAddSucceeded = $false; $script:mockPluginAddSucceeded = $true; $script:mockRemoveSucceeded = $true
    $persistedAfterFailedAdd = $null
    try {
        Try-InstallWithCodexCli -ExecutablePath 'C:\trusted\codex.exe' -VerifiedMarketplaceRoot 'D:\verified-release' | Out-Null
    } catch {
        $persistedAfterFailedAdd = $_
    }
    Assert-True -Condition ($null -ne $persistedAfterFailedAdd -and @($script:mockCliCalls | Where-Object { $_ -eq "plugin marketplace remove $MarketplaceName" }).Count -eq 1) -Message 'add 非零且条目残留时必须尝试精确回滚；最终仍残留则必须阻止个人回退。'

    # 同名条目 root 不匹配时绝不能调用 remove，否则可能删除用户原有 Marketplace。
    $script:mockCliCalls.Clear(); $script:mockInventorySteps.Clear()
    foreach ($state in @('none', 'mismatched-root')) { $script:mockInventorySteps.Enqueue($state) }
    $script:mockAddSucceeded = $true; $script:mockPluginAddSucceeded = $true; $script:mockRemoveSucceeded = $true
    $mismatchedRoot = $null
    try {
        Try-InstallWithCodexCli -ExecutablePath 'C:\trusted\codex.exe' -VerifiedMarketplaceRoot 'D:\verified-release' | Out-Null
    } catch {
        $mismatchedRoot = $_
    }
    Assert-True -Condition ($null -ne $mismatchedRoot -and $script:mockCliCalls -notcontains "plugin marketplace remove $MarketplaceName") -Message 'root 不匹配时必须拒绝 remove 并阻止个人回退。'

    Write-Host 'PASS: personal-marketplace-installer-tests.ps1'
} finally {
    # 只删除本测试刚在系统临时目录内创建的 GUID 目录；先验证边界，避免误删用户文件。
    $temporaryPrefix = "$temporaryParent$([System.IO.Path]::DirectorySeparatorChar)"
    if ($testRoot.StartsWith($temporaryPrefix, [System.StringComparison]::OrdinalIgnoreCase) -and [System.IO.Directory]::Exists($testRoot)) {
        [System.IO.Directory]::Delete($testRoot, $true)
    }
}
