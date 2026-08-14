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

function Get-PptxShapeNames {
    param([string]$PresentationPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($PresentationPath)
        $names = @()
        foreach ($entry in @($archive.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' })) {
            $reader = $null
            try {
                $reader = New-Object IO.StreamReader($entry.Open())
                $xml = $reader.ReadToEnd()
                foreach ($match in [regex]::Matches($xml, '<p:cNvPr[^>]*\bname="([^"]+)"')) {
                    $names += $match.Groups[1].Value
                }
            } finally {
                if ($reader) { $reader.Dispose() }
            }
        }
        return @($names)
    } finally {
        if ($archive) { $archive.Dispose() }
    }
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

    # 图片长期保留是严格布尔开关；字符串值和任意导出路径参数都不能在启动 Office 前
    # 改变文件系统行为。
    $stringFigureOutput = Join-Path $securityRoot 'string-figure-assets.pptx'
    $stringFigureAssets = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 24.5; method = 'tools/call'
        params = [ordered]@{ name = 'generate_editable_pptx'; arguments = @{ deck_spec = $deck; output_path = $stringFigureOutput; export_figure_assets = 'false' } }
    })
    Assert-Condition ($null -ne $stringFigureAssets.error -and $stringFigureAssets.error.message -match 'JSON boolean') 'String export_figure_assets=false must be rejected before PowerPoint generation.'
    Assert-Condition (-not (Test-Path -LiteralPath $stringFigureOutput)) 'Rejected figure-asset export flag must not create a PPTX.'

    $unknownFigureExportArgument = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 24.6; method = 'tools/call'
        params = [ordered]@{ name = 'generate_editable_pptx'; arguments = @{ deck_spec = $deck; output_path = (Join-Path $securityRoot 'unknown-figure-assets.pptx'); asset_export_directory = $securityRoot } }
    })
    Assert-Condition ($null -ne $unknownFigureExportArgument.error -and $unknownFigureExportArgument.error.message -match 'unsupported argument') 'Figure-asset export must not accept a caller-controlled output directory.'

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

    # 仅位于插件临时根并不足以成为可删对象：长期保存的 <PPT名>_assets 若恰好位于
    # 临时根，也必须因缺少 analyse_paper 所有权标记而受保护。
    $unownedCleanup = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 25.5; method = 'tools/call'
        params = [ordered]@{ name = 'cleanup_paper_assets'; arguments = @{ asset_output_dir = $assetDirectory; confirm = $true } }
    })
    Assert-Condition ($null -ne $unownedCleanup.error -and $unownedCleanup.error.message -match 'ownership marker') 'Cleanup must refuse an unowned directory even when confirm=true.'
    Assert-Condition (Test-Path -LiteralPath $assetSentinel -PathType Leaf) 'Ownership-marker cleanup rejection must preserve files.'

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
    # 测试直接构造的图片资产必须位于服务端认可的插件临时根，而不是 e2e-tests
    # 的父级。否则 JSON 回传到独立 MCP 进程后会被路径边界正确拒绝。
    $assetWorkDirectory = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\e2e-figure-assets\$([Guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Force -Path $assetWorkDirectory | Out-Null
    # 默认交付物应保持干净：仅输出 PPTX；内部 deck spec 在生成后删除，PNG 仅按需导出。
    $defaultPptxPath = Join-Path $workDirectory 'journal-club-default.pptx'
    $defaultGeneration = Invoke-McpTool 6 'generate_editable_pptx' @{
        deck_spec = $deck
        output_path = $defaultPptxPath
        keep_powerpoint_open = $false
    }
    Assert-Condition (Test-Path -LiteralPath $defaultPptxPath -PathType Leaf) 'Default generation must write a PPTX.'
    Assert-Condition ($defaultGeneration.quality_audit.pass) 'Default generation must still run the native quality audit.'
    Assert-Condition (-not $defaultGeneration.deck_spec_saved -and $null -eq $defaultGeneration.deck_spec_path) 'Default generation must not persist a deck-spec sidecar.'
    $defaultPreviewPaths = @($defaultGeneration.preview_paths | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) })
    Assert-Condition ([string]::IsNullOrWhiteSpace([string]$defaultGeneration.preview_directory)) 'Default generation must not return a preview directory.'
    Assert-Condition ($defaultPreviewPaths.Count -eq 0) 'Default generation must not export PNG previews.'
    Assert-Condition (-not $defaultGeneration.figure_assets_exported -and $null -eq $defaultGeneration.figure_assets_directory -and @($defaultGeneration.figure_asset_paths).Count -eq 0) 'Default generation must not export paper figure assets.'
    Assert-Condition (-not (Test-Path -LiteralPath "$defaultPptxPath.deck-spec.json")) 'Default generation must not create a deck-spec file next to the PPTX.'
    Assert-Condition (-not (Test-Path -LiteralPath "$defaultPptxPath.previews")) 'Default generation must not create a preview directory next to the PPTX.'
    Assert-Condition (-not (Test-Path -LiteralPath (Join-Path $workDirectory 'journal-club-default_assets'))) 'Default generation must not create a figure-asset directory next to the PPTX.'
    $defaultShapeNames = @(Get-PptxShapeNames -PresentationPath $defaultPptxPath)
    Assert-Condition (@($defaultShapeNames | Where-Object { $_ -match '^slide-\d+-evidence$' }).Count -eq 0) 'Text-only slides must not contain an Evidence/figure placeholder card.'

    # 一张经批准、可追溯的论文图片必须成为真正的 PowerPoint 图片对象；这同时覆盖
    # “有图则插入”与生成器的安全资产目录边界，避免仅在 deck spec 层面看起来正确。
    $automaticFigurePath = Join-Path $assetWorkDirectory 'automatic-figure.png'
    Add-Type -AssemblyName System.Drawing
    $testBitmap = New-Object System.Drawing.Bitmap 640, 360
    $testGraphics = [System.Drawing.Graphics]::FromImage($testBitmap)
    try {
        $testGraphics.Clear([System.Drawing.Color]::White)
        $testGraphics.FillRectangle([System.Drawing.Brushes]::SteelBlue, 40, 40, 560, 280)
        $testGraphics.DrawLine([System.Drawing.Pens]::White, 80, 260, 560, 100)
        $testBitmap.Save($automaticFigurePath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        if ($testGraphics) { $testGraphics.Dispose() }
        if ($testBitmap) { $testBitmap.Dispose() }
    }
    $figureDeck = $deck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    # JSON round-trip 后 asset_directory 是已有但值为 null 的属性；Add-Member -Force 在
    # Windows PowerShell 5.1 不会可靠替换此类属性，因此直接赋值即可。
    $figureDeck.evidence_pack.extraction.asset_directory = $assetWorkDirectory
    $figureDeck.evidence_pack.figures += [pscustomobject]@{
        id = 'fig-automatic-e2e'
        label = 'Fig. automatic e2e'
        context = 'A single extracted experimental figure for end-to-end validation.'
        source_page = 1
        source_pages = @(1)
        figure_asset_candidates = @($automaticFigurePath)
        asset_match = 'same-page-single-raster'
        automatic_image_path = $automaticFigurePath
    }
    $figureSlide = @($figureDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    Assert-Condition ($figureSlide.Count -eq 1) 'Fixture must contain an experimental-data slide for automatic figure insertion.'
    $figureSlide[0].source_figure_ids = @('fig-automatic-e2e')
    $figureSlide[0] | Add-Member -NotePropertyName 'suggested_figure_id' -NotePropertyValue 'fig-automatic-e2e' -Force
    $figureSlide[0] | Add-Member -NotePropertyName 'suggested_image_path' -NotePropertyValue $automaticFigurePath -Force
    $figureSlide[0] | Add-Member -NotePropertyName 'figure_asset_candidates' -NotePropertyValue @($automaticFigurePath) -Force
    $figurePptxPath = Join-Path $workDirectory 'journal-club-with-figure.pptx'
    $figureGeneration = Invoke-McpTool 6.5 'generate_editable_pptx' @{
        deck_spec = $figureDeck
        output_path = $figurePptxPath
        export_figure_assets = $true
        keep_powerpoint_open = $false
    }
    Assert-Condition ($figureGeneration.quality_audit.pass) 'A deck with one approved figure must pass PowerPoint quality audit.'
    $expectedFigureAssetDirectory = Join-Path $workDirectory 'journal-club-with-figure_assets\images'
    Assert-Condition ($figureGeneration.figure_assets_exported -and $figureGeneration.figure_assets_directory -eq $expectedFigureAssetDirectory) 'Explicit figure-asset export must return the derived sibling images directory.'
    Assert-Condition (@($figureGeneration.figure_asset_paths).Count -eq 1 -and (Test-Path -LiteralPath $figureGeneration.figure_asset_paths[0] -PathType Leaf)) 'Explicit figure-asset export must copy exactly the inserted figure.'
    $exportedFigureBytes = [IO.File]::ReadAllBytes($figureGeneration.figure_asset_paths[0])
    $sourceFigureBytes = [IO.File]::ReadAllBytes($automaticFigurePath)
    Assert-Condition ([Convert]::ToBase64String($exportedFigureBytes) -eq [Convert]::ToBase64String($sourceFigureBytes)) 'Exported figure asset must preserve the original approved raster bytes.'
    Assert-Condition ($figureGeneration.figure_asset_export.exported_count -eq 1 -and $figureGeneration.figure_asset_export.assets[0].source_sha256 -eq $figureGeneration.figure_asset_export.assets[0].sha256) 'Figure-asset export manifest must report the verified source hash.'
    $figureShapeNames = @(Get-PptxShapeNames -PresentationPath $figurePptxPath)
    Assert-Condition (@($figureShapeNames | Where-Object { $_ -match '^slide-\d+-figure-image$' }).Count -eq 1) 'A reliable extracted figure must be embedded as one editable PowerPoint picture object.'

    # 导出资产即使位于插件临时根，也不是 analyse_paper 的临时目录；清理工具不得删除。
    $exportedAssetCleanup = Invoke-McpRaw ([ordered]@{
        jsonrpc = '2.0'; id = 6.6; method = 'tools/call'
        params = [ordered]@{ name = 'cleanup_paper_assets'; arguments = @{ asset_output_dir = $expectedFigureAssetDirectory; confirm = $true } }
    })
    Assert-Condition ($null -ne $exportedAssetCleanup.error -and $exportedAssetCleanup.error.message -match 'ownership marker') 'Cleanup must refuse the persistent figure-asset directory.'
    Assert-Condition (Test-Path -LiteralPath $figureGeneration.figure_asset_paths[0] -PathType Leaf) 'Rejected cleanup must preserve the exported source image.'

    $pptxPath = Join-Path $workDirectory 'journal-club.pptx'
    $previewDirectory = Join-Path $workDirectory 'previews'
    $deckSpecPath = Join-Path $workDirectory 'journal-club.deck-spec.json'
    $generation = Invoke-McpTool 7 'generate_editable_pptx' @{
        deck_spec = $deck
        output_path = $pptxPath
        deck_spec_output_path = $deckSpecPath
        preview_directory = $previewDirectory
        export_previews = $true
        keep_powerpoint_open = $false
    }
    Assert-Condition (Test-Path -LiteralPath $pptxPath -PathType Leaf) 'Generation must write a PPTX.'
    Assert-Condition ($generation.quality_audit.pass) 'Native PowerPoint quality audit must pass.'
    Assert-Condition ($generation.deck_spec_saved -and (Test-Path -LiteralPath $deckSpecPath -PathType Leaf)) 'An explicit deck-spec output path must be preserved.'
    Assert-Condition (@($generation.preview_paths).Count -eq @($deck.slides).Count) 'Every generated slide must have a native preview.'

    # 重新打开文件验证，确保不是仅在生成会话中看似可用。
    $fileAudit = Invoke-McpTool 8 'audit_editable_pptx' @{ file_path = $pptxPath; export_previews = $false }
    Assert-Condition ($fileAudit.connection_scope -eq 'file-read-only') 'Saved-deck audit must truthfully report file-read-only scope.'
    Assert-Condition ($fileAudit.quality_audit.pass) 'Reopened PPTX must pass native PowerPoint audit.'
    Write-Output 'PASS: end-to-end-tests.ps1 (including PowerPoint generation and reopen audit)'
} finally {
    if (Test-Path -LiteralPath $assetWorkDirectory) {
        # 仅删除本测试刚创建的、带 GUID 的临时图像目录。
        Remove-Item -LiteralPath $assetWorkDirectory -Recurse -Force
    }
    if ((Test-Path -LiteralPath $workDirectory) -and -not $KeepArtifacts) {
        # 仅删除本测试刚创建的、带 GUID 的临时目录。
        Remove-Item -LiteralPath $workDirectory -Recurse -Force
    }
}
