<#
  发布前端到端回归测试。

  默认只验证 MCP 协议、论文证据链和内容门禁，因此可在没有 PowerPoint 的 CI 环境中运行。
  传入 -RunPowerPoint 后，会在用户临时目录新建 PPTX、导出原生预览并重新打开做只读审计。
#>
[CmdletBinding()]
param(
    [switch]$RunPowerPoint,
    [switch]$KeepArtifacts,
    [string]$PluginRoot = ''
)

$ErrorActionPreference = 'Stop'
$PluginRoot = if ([string]::IsNullOrWhiteSpace($PluginRoot)) { Join-Path $PSScriptRoot '..' } else { $PluginRoot }
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$pluginFullPath = [IO.Path]::GetFullPath($PluginRoot)
$serverPath = Join-Path $pluginFullPath 'scripts\paper-to-journal-club-server.ps1'
$fixturePath = Join-Path $pluginFullPath 'examples\sample-paper.md'
$pluginManifestPath = Join-Path $pluginFullPath '.codex-plugin\plugin.json'

function Assert-Condition {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Invoke-McpRaw {
    param($Request)
    # 显式声明标准输入/输出为 UTF-8，防止宿主 PowerShell 的 OEM 代码页损坏中文 JSON。
    $jsonLine = $Request | ConvertTo-Json -Depth 100 -Compress
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'powershell.exe'
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy RemoteSigned -File `"$serverPath`""
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = $utf8

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        [void]$process.Start()
        $process.StandardInput.WriteLine($jsonLine)
        $process.StandardInput.Close()
        $standardOutput = $process.StandardOutput.ReadToEnd()
        $standardError = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "MCP server exited with code $($process.ExitCode): $standardError" }
    } finally {
        $process.Dispose()
    }

    $responseLine = @($standardOutput -split "`r?`n" | Where-Object { $_ -and $_.TrimStart().StartsWith('{') } | Select-Object -Last 1)
    if (-not $responseLine) { throw 'MCP server returned no JSON response.' }
    return $responseLine[0] | ConvertFrom-Json
}

function Invoke-McpTool {
    param([int]$Id, [string]$Name, $Arguments)
    $response = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'
        id = $Id
        method = 'tools/call'
        params = [ordered]@{ name = $Name; arguments = $Arguments }
    })
    if ($response.error) { throw "MCP $Name failed: $($response.error.message)" }
    return $response.result.content[0].text | ConvertFrom-Json
}

if (-not (Test-Path -LiteralPath $serverPath -PathType Leaf)) { throw "Server script not found: $serverPath" }
if (-not (Test-Path -LiteralPath $fixturePath -PathType Leaf)) { throw "Fixture paper not found: $fixturePath" }
$expectedPluginVersion = [string]((Get-Content -LiteralPath $pluginManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json).version)
if ([string]::IsNullOrWhiteSpace($expectedPluginVersion)) { throw 'Plugin manifest version is required for the MCP handshake test.' }

# 1) MCP 协商必须支持当前协议，并向 Agent 暴露真实的 PowerPoint 能力边界。
$initialize = Invoke-McpRaw ([ordered]@{
    jsonrpc = '2.0'
    id = 1
    method = 'initialize'
    params = [ordered]@{ protocolVersion = '2025-06-18'; capabilities = [ordered]@{}; clientInfo = [ordered]@{ name = 'release-e2e-test'; version = $expectedPluginVersion } }
})
Assert-Condition (-not $initialize.error) 'initialize must succeed.'
Assert-Condition ($initialize.result.protocolVersion -eq '2025-06-18') 'Server must negotiate MCP 2025-06-18.'
Assert-Condition ($initialize.result.serverInfo.version -eq $expectedPluginVersion) 'Server version must match the release manifest.'

$status = Invoke-McpTool 2 'powerpoint_status' @{}
Assert-Condition ($status.target_application -eq 'Microsoft PowerPoint') 'Status must target Microsoft PowerPoint.'
Assert-Condition ($status.generation_scope -eq 'new-background-presentation') 'Generation scope must be explicit and truthful.'

# 2) 论文到 deck-spec 的主链路必须默认包含六个组会模块，并在生成前通过来源审核。
$evidence = Invoke-McpTool 3 'analyse_paper' @{ file_path = [IO.Path]::GetFullPath($fixturePath) }
$deck = Invoke-McpTool 4 'design_journal_club_deck' @{ evidence_pack = $evidence; duration_minutes = 15; language = 'zh-CN'; audience = 'lab' }
$expectedSections = @('background', 'innovation', 'methods', 'experimental_data', 'limitations', 'future_directions')
Assert-Condition ((@($deck.required_sections) -join '|') -eq ($expectedSections -join '|')) 'Deck must include the six default journal-club modules.'
$contentAudit = Invoke-McpTool 5 'audit_journal_club_deck' @{ deck_spec = $deck }
Assert-Condition ($contentAudit.pass) 'Complete fixture must pass content audit.'
Assert-Condition ($contentAudit.quality.safe_to_generate) 'Content audit must mark the deck safe to generate.'

# 3) 安全回归：字符串 "false" 绝不能被 PowerShell 当作 $true，未知参数和越界路径必须在
# 调用任何 Office COM 或文件写入前被拒绝。测试目录位于插件专用临时根，清理目标固定且可验证。
$securityRoot = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\security-tests\$([Guid]::NewGuid().ToString('N'))"
$outsideRoot = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club-security-outside-$([Guid]::NewGuid().ToString('N'))"
try {
    New-Item -ItemType Directory -Force -Path $securityRoot, $outsideRoot | Out-Null
    $outsidePaper = Join-Path $outsideRoot 'outside-boundary.txt'
    Set-Content -LiteralPath $outsidePaper -Value 'This file must not be readable through the MCP boundary.' -Encoding UTF8

    $stringBoolean = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 20; method = 'tools/call'
        params = [ordered]@{ name = 'inspect_powerpoint'; arguments = @{ export_previews = 'false' } }
    })
    Assert-Condition ($null -ne $stringBoolean.error -and $stringBoolean.error.message -match 'JSON boolean') 'String false must be rejected instead of being coerced to true.'

    # 该请求超过服务端 1 MiB 限制；服务端应逐字符丢弃余量并返回 JSON-RPC 错误，
    # 而不是先由 ReadLine 分配完整无限输入。
    $oversizedPath = 'x' * (1MB + 4096)
    $oversizedRequest = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 20.5; method = 'tools/call'
        params = [ordered]@{ name = 'analyse_paper'; arguments = @{ file_path = $oversizedPath } }
    })
    Assert-Condition ($null -ne $oversizedRequest.error -and $oversizedRequest.error.message -match 'MCP request exceeds') 'Oversized MCP lines must be rejected before JSON parsing.'

    $unknownArgument = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 21; method = 'tools/call'
        params = [ordered]@{ name = 'analyse_paper'; arguments = @{ file_path = [IO.Path]::GetFullPath($fixturePath); unexpected = $true } }
    })
    Assert-Condition ($null -ne $unknownArgument.error -and $unknownArgument.error.message -match 'unsupported argument') 'Runtime validation must reject undeclared MCP arguments.'

    $outsideRead = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 22; method = 'tools/call'
        params = [ordered]@{ name = 'analyse_paper'; arguments = @{ file_path = $outsidePaper } }
    })
    Assert-Condition ($null -ne $outsideRead.error -and $outsideRead.error.message -match 'approved user data directory') 'MCP must reject paper paths outside approved roots.'

    $outsideOutput = Join-Path $outsideRoot 'outside-boundary.deck-spec.json'
    $outsideWrite = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 23; method = 'tools/call'
        params = [ordered]@{ name = 'design_journal_club_deck'; arguments = @{ evidence_pack = $evidence; output_path = $outsideOutput } }
    })
    Assert-Condition ($null -ne $outsideWrite.error -and $outsideWrite.error.message -match 'approved user data directory') 'MCP must reject output paths outside approved roots.'
    Assert-Condition (-not (Test-Path -LiteralPath $outsideOutput)) 'Rejected out-of-bound output request must not create a file.'

    $existingSpec = Join-Path $securityRoot 'existing.deck-spec.json'
    Set-Content -LiteralPath $existingSpec -Value 'do-not-overwrite' -Encoding UTF8
    $stringOverwrite = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 24; method = 'tools/call'
        params = [ordered]@{ name = 'design_journal_club_deck'; arguments = @{ evidence_pack = $evidence; output_path = $existingSpec; overwrite = 'false' } }
    })
    Assert-Condition ($null -ne $stringOverwrite.error -and $stringOverwrite.error.message -match 'JSON boolean') 'String overwrite=false must not authorize replacing an existing file.'
    Assert-Condition ((Get-Content -LiteralPath $existingSpec -Raw -Encoding UTF8).Trim() -eq 'do-not-overwrite') 'Rejected overwrite request must leave the existing file unchanged.'

    $assetDirectory = Join-Path $securityRoot 'assets'
    New-Item -ItemType Directory -Force -Path $assetDirectory | Out-Null
    $assetSentinel = Join-Path $assetDirectory 'sentinel.txt'
    Set-Content -LiteralPath $assetSentinel -Value 'keep' -Encoding UTF8
    $stringConfirm = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 25; method = 'tools/call'
        params = [ordered]@{ name = 'cleanup_paper_assets'; arguments = @{ asset_output_dir = $assetDirectory; confirm = 'false' } }
    })
    Assert-Condition ($null -ne $stringConfirm.error -and $stringConfirm.error.message -match 'JSON boolean') 'String confirm=false must not authorize deletion.'
    Assert-Condition (Test-Path -LiteralPath $assetSentinel -PathType Leaf) 'Rejected cleanup request must preserve temporary assets.'

    # SVG/EMF/WMF 等可执行或复杂矢量格式在进入 PowerPoint COM 前必须被拒绝。
    . (Join-Path $pluginFullPath 'scripts\powerpoint-quality.ps1')
    $svgPath = Join-Path $assetDirectory 'untrusted.svg'
    Set-Content -LiteralPath $svgPath -Value '<svg xmlns="http://www.w3.org/2000/svg"></svg>' -Encoding UTF8
    $svgRejected = $false
    try { Get-PaperToJournalClubApprovedRasterImage -ImagePath $svgPath -AllowedRoots @($assetDirectory) | Out-Null } catch { $svgRejected = $_.Exception.Message -match 'PNG or JPEG' }
    Assert-Condition $svgRejected 'SVG must be rejected before PowerPoint image insertion.'

    # 仅写入 PNG 头部且宣称 10001 像素宽；安全检查必须在 GDI+ 解码前根据 IHDR 拒绝它。
    $pngBombPath = Join-Path $assetDirectory 'dimension-bomb.png'
    [byte[]]$pngHeader = 137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 39, 17, 0, 0, 0, 1
    [IO.File]::WriteAllBytes($pngBombPath, $pngHeader)
    $pngBombRejected = $false
    try { Get-PaperToJournalClubApprovedRasterImage -ImagePath $pngBombPath -AllowedRoots @($assetDirectory) | Out-Null } catch { $pngBombRejected = $_.Exception.Message -match 'safety limit' }
    Assert-Condition $pngBombRejected 'Oversized PNG dimensions must be rejected before GDI+ decoding.'
} finally {
    if (Test-Path -LiteralPath $securityRoot) { Remove-Item -LiteralPath $securityRoot -Recurse -Force }
    if (Test-Path -LiteralPath $outsideRoot) { Remove-Item -LiteralPath $outsideRoot -Recurse -Force }
}

if (-not $RunPowerPoint) {
    Write-Output 'PASS: end-to-end-tests.ps1 (MCP and content path)'
    exit 0
}

if (-not $status.powerpoint_com_registered) { throw 'PowerPoint COM is required when -RunPowerPoint is specified.' }
$workDirectory = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\e2e-tests\$([Guid]::NewGuid().ToString('N'))"
try {
    $pptxPath = Join-Path $workDirectory 'journal-club.pptx'
    $previewDirectory = Join-Path $workDirectory 'previews'
    $generation = Invoke-McpTool 6 'generate_editable_pptx' @{
        deck_spec = $deck
        output_path = $pptxPath
        preview_directory = $previewDirectory
        export_previews = $true
        keep_powerpoint_open = $false
    }
    Assert-Condition (Test-Path -LiteralPath $pptxPath -PathType Leaf) 'Generation must write a PPTX.'
    Assert-Condition ($generation.quality_audit.pass) 'Native PowerPoint quality audit must pass.'
    Assert-Condition (@($generation.preview_paths).Count -eq @($deck.slides).Count) 'Every generated slide must have a native preview.'

    # 重新打开文件验证，确保不是仅在生成会话中看似可用。
    $fileAudit = Invoke-McpTool 7 'audit_editable_pptx' @{ file_path = $pptxPath; export_previews = $false }
    Assert-Condition ($fileAudit.connection_scope -eq 'file-read-only') 'Saved-deck audit must truthfully report file-read-only scope.'
    Assert-Condition ($fileAudit.quality_audit.pass) 'Reopened PPTX must pass native PowerPoint audit.'
    Write-Output 'PASS: end-to-end-tests.ps1 (including PowerPoint generation and reopen audit)'
} finally {
    if ((Test-Path -LiteralPath $workDirectory) -and -not $KeepArtifacts) {
        # 仅删除本测试刚创建的、带 GUID 的临时目录。
        Remove-Item -LiteralPath $workDirectory -Recurse -Force
    }
}
