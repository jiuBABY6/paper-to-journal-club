<#
  组会内容模板回归测试。
  测试只通过公开的 PowerShell MCP 入口交互，确保默认模块、证据回链和硬性审核门禁
  在真实插件运行路径上都能工作。
#>
param(
    [string]$ServerPath = (Join-Path $PSScriptRoot '..\scripts\paper-to-journal-club-server.ps1')
)

$ErrorActionPreference = 'Stop'
# 通过 Windows PowerShell 调用子进程时显式使用 UTF-8，避免中文 evidence_pack 在管道中变成问号。
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$TestFixtureRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'fixtures'))
if (-not (Test-Path -LiteralPath $TestFixtureRoot -PathType Container)) {
    throw "Test fixture directory not found: $TestFixtureRoot"
}
$serverFullPath = [IO.Path]::GetFullPath($ServerPath)
if (-not (Test-Path -LiteralPath $serverFullPath)) { throw "Server script not found: $serverFullPath" }

function Assert-Condition {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Invoke-McpResponse {
    param($Request)
    # 不使用 Windows PowerShell 的原生管道：其输出编码会受调用脚本作用域影响，中文 JSON
    # 可能被替换为问号。显式指定子进程双向 UTF-8，才能真实覆盖中文论文的 MCP 路径。
    $jsonLine = $Request | ConvertTo-Json -Depth 100 -Compress
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'powershell.exe'
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy RemoteSigned -File `"$serverFullPath`""
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = $utf8

    # GitHub runner 的工作目录不属于普通用户 Desktop/Documents/Downloads。仅给本测试
    # 启动的 MCP 子进程追加固定夹具目录，既覆盖中文样本，又不改变插件在真实用户环境
    # 中的文件访问白名单，也不污染执行测试的父 PowerShell 会话。
    $existingAllowedRoots = [Environment]::GetEnvironmentVariable('PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS', 'Process')
    $childAllowedRoots = if ([string]::IsNullOrWhiteSpace($existingAllowedRoots)) {
        $TestFixtureRoot
    } else {
        "$existingAllowedRoots;$TestFixtureRoot"
    }
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        # Windows PowerShell 5.1 对 ProcessStartInfo.EnvironmentVariables 存在空集合
        # 兼容性问题。只在启动子进程的瞬间设置当前进程变量，子进程会继承它；随后立即
        # 恢复，因而不会影响本测试之外的任何 MCP 调用。
        try {
            [Environment]::SetEnvironmentVariable('PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS', [string]$childAllowedRoots, 'Process')
            [void]$process.Start()
        } finally {
            [Environment]::SetEnvironmentVariable('PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS', $existingAllowedRoots, 'Process')
        }
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

function Invoke-McpRequest {
    param($Request)
    $response = Invoke-McpResponse $Request
    if ($response.error) { throw "MCP error: $($response.error.message)" }
    return $response.result.content[0].text | ConvertFrom-Json
}

function New-McpToolRequest {
    param([int]$Id, [string]$ToolName, $Arguments)
    return [ordered]@{
        jsonrpc = '2.0'
        id = $Id
        method = 'tools/call'
        params = [ordered]@{
            name = $ToolName
            arguments = $Arguments
        }
    }
}

function Invoke-DesignFromPaper {
    param([string]$PaperPath, [int]$RequestId = 1, [string[]]$RequiredSections = $null)
    $evidence = Invoke-McpRequest (New-McpToolRequest -Id $RequestId -ToolName 'analyse_paper' -Arguments @{ file_path = [IO.Path]::GetFullPath($PaperPath) })
    $designArguments = [ordered]@{ evidence_pack = $evidence; duration_minutes = 15; language = 'zh-CN'; audience = 'lab' }
    if ($null -ne $RequiredSections) { $designArguments.required_sections = $RequiredSections }
    return Invoke-McpRequest (New-McpToolRequest -Id ($RequestId + 1) -ToolName 'design_journal_club_deck' -Arguments $designArguments)
}

$expectedSections = @('background', 'innovation', 'methods', 'experimental_data', 'limitations', 'future_directions')
$englishFixture = Join-Path $PSScriptRoot '..\examples\sample-paper.md'
$chineseFixture = Join-Path $PSScriptRoot 'fixtures\sample-paper-zh.md'

# 0) MCP schema 必须向调用方暴露必备模块和安全的图片资产输出目录配置。
$toolsResponse = Invoke-McpResponse ([ordered]@{ jsonrpc = '2.0'; id = 1; method = 'tools/list'; params = @{} })
Assert-Condition (-not $toolsResponse.error) 'tools/list must succeed.'
$designTool = @($toolsResponse.result.tools | Where-Object { $_.name -eq 'design_journal_club_deck' } | Select-Object -First 1)
$analyseTool = @($toolsResponse.result.tools | Where-Object { $_.name -eq 'analyse_paper' } | Select-Object -First 1)
$generateTool = @($toolsResponse.result.tools | Where-Object { $_.name -eq 'generate_editable_pptx' } | Select-Object -First 1)
$auditPptxTool = @($toolsResponse.result.tools | Where-Object { $_.name -eq 'audit_editable_pptx' } | Select-Object -First 1)
$renderVisualTool = @($toolsResponse.result.tools | Where-Object { $_.name -eq 'render_paper_visual' } | Select-Object -First 1)
Assert-Condition ($designTool.Count -eq 1 -and $null -ne $designTool[0].inputSchema.properties.required_sections) 'Design schema must expose required_sections.'
Assert-Condition ($analyseTool.Count -eq 1 -and $null -ne $analyseTool[0].inputSchema.properties.asset_output_dir) 'Analyse schema must expose optional asset_output_dir.'
Assert-Condition ($generateTool.Count -eq 1 -and $null -ne $generateTool[0].inputSchema.properties.deck_spec_output_path) 'Generation schema must expose an explicit optional deck-spec export path.'
Assert-Condition ($generateTool[0].inputSchema.properties.export_previews.default -eq $false) 'Generation must not export PNG previews by default.'
Assert-Condition ($generateTool[0].inputSchema.properties.export_figure_assets.default -eq $false) 'Generation must not export original figure assets by default.'
Assert-Condition ($generateTool[0].inputSchema.properties.export_figure_assets.description -match 'PPT-name.*_assets') 'Figure-asset export schema must describe the fixed sibling output directory.'
Assert-Condition ($auditPptxTool.Count -eq 1 -and $auditPptxTool[0].inputSchema.properties.export_previews.default -eq $false) 'Saved-PPTX audit must not export PNG previews by default.'
Assert-Condition ($designTool[0].inputSchema.properties.figure_asset_selection.description -match 'automatic insertion') 'Design schema must explain the conservative automatic figure-insertion rule.'
Assert-Condition ($renderVisualTool.Count -eq 1 -and $renderVisualTool[0].inputSchema.properties.crop) 'Schema must expose an explicit reviewed crop for PDF table or vector visual candidates.'
Assert-Condition ($renderVisualTool[0].description -match 'never inserted automatically') 'Rendered PDF visuals must be described as opt-in rather than automatic evidence.'

# 1) 默认英文论文必须产生六个必备模块，且每个模块都有来源与状态标记。
$deck = Invoke-DesignFromPaper $englishFixture 100
Assert-Condition (($deck.required_sections -join '|') -eq ($expectedSections -join '|')) 'Default required_sections must contain all six journal-club modules in order.'
foreach ($section in $expectedSections) {
    $sectionSlides = @($deck.slides | Where-Object { $_.section -eq $section })
    Assert-Condition ($sectionSlides.Count -gt 0) "Missing generated section: $section"
    foreach ($slide in $sectionSlides) {
        Assert-Condition ($slide.evidence_status -eq 'source-backed') "Required section $section must be source-backed in the complete fixture."
        Assert-Condition ((@($slide.source_claim_ids).Count + @($slide.source_section_ids).Count + @($slide.source_figure_ids).Count) -gt 0) "Required section $section must contain a traceable source id."
    }
}
$dataSlides = @($deck.slides | Where-Object { $_.section -eq 'experimental_data' })
Assert-Condition (@($dataSlides | Where-Object { @($_.source_claim_ids).Count -eq 0 }).Count -eq 0) 'Every experimental-data slide must cite at least one claim.'
foreach ($slide in @($deck.slides | Where-Object { $_.kind -ne 'title' })) {
    Assert-Condition ($slide.title -notmatch '[:：?？]') "Non-title slide heading must be declarative: $($slide.title)"
    $titleLimit = if ($slide.title -match '[\u4e00-\u9fff]') { 28 } else { 64 }
    Assert-Condition ($slide.title.Length -le $titleLimit) "Non-title heading must remain within the readability limit: $($slide.id)"
    Assert-Condition (@($slide.bullets).Count -le 3) "Slide must have at most three supporting points: $($slide.id)"
    $takeawayLimit = if ($slide.takeaway -match '[\u4e00-\u9fff]') { 90 } else { 180 }
    Assert-Condition ($slide.takeaway.Length -le $takeawayLimit) "Slide takeaway must remain within the readability limit: $($slide.id)"
    foreach ($bullet in @($slide.bullets)) {
        $bulletLimit = if ([string]$bullet -match '[\u4e00-\u9fff]') { 60 } else { 120 }
        Assert-Condition (([string]$bullet).Length -le $bulletLimit) "Supporting point must remain within the readability limit: $($slide.id)"
    }
}
foreach ($dataSlide in $dataSlides) {
    Assert-Condition ($null -ne $dataSlide.result_analysis) "Experimental-data slide must include structured result analysis: $($dataSlide.id)"
    Assert-Condition (-not [string]::IsNullOrWhiteSpace($dataSlide.result_analysis.comparison)) "Experimental-data slide must state a source-grounded comparison: $($dataSlide.id)"
    Assert-Condition (-not [string]::IsNullOrWhiteSpace($dataSlide.result_analysis.interpretation)) "Experimental-data slide must include an interpretation: $($dataSlide.id)"
    Assert-Condition (-not [string]::IsNullOrWhiteSpace($dataSlide.result_analysis.caveat)) "Experimental-data slide must include a caveat: $($dataSlide.id)"
}
$sectionOrder = @('summary', 'background', 'innovation', 'methods', 'experimental_data', 'limitations', 'future_directions', 'takeaway')
$lastIndex = -1
foreach ($sectionName in $sectionOrder) {
    $sectionIndex = [array]::FindIndex(@($deck.slides), [Predicate[object]]{ param($item) $item.section -eq $sectionName })
    if ($sectionIndex -ge 0) {
        Assert-Condition ($sectionIndex -gt $lastIndex) "Narrative section order must remain coherent around $sectionName."
        $lastIndex = $sectionIndex
    }
}
Assert-Condition ($deck.slides[-1].section -eq 'takeaway') 'The take-home conclusion must be the final slide.'
$futureSlide = @($deck.slides | Where-Object { $_.section -eq 'future_directions' } | Select-Object -First 1)
Assert-Condition ($futureSlide[0].content_mode -eq 'presenter-discussion') 'Fixture without author future wording should use an explicitly labelled presenter discussion.'

# 生成器必须在启动 PowerPoint COM 前接受服务端实际输出的结构化 result_analysis。
# 这条回归不依赖 Office：只执行 generator 的输入预检，预期随后因故意传入的不存在
# output 目录而停止。若 result_analysis 被错误当作标量，该命令会在到达路径检查前失败。
$generatorPath = Join-Path $PSScriptRoot '..\scripts\generate-editable-pptx.ps1'
$generatorDirectory = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\generator-preflight-$([Guid]::NewGuid().ToString('N'))"
$generatorDeckPath = Join-Path $generatorDirectory 'deck.json'
try {
    New-Item -ItemType Directory -Path $generatorDirectory -Force | Out-Null
    # 输出文件预先存在，确保在 deck 预检完成后、启动任何 PowerPoint COM 前因安全覆盖
    # 保护而停止。不能依赖不存在的父目录，因为生成器允许创建用户指定的输出目录。
    $generatorOutputPath = Join-Path $generatorDirectory 'existing-output.pptx'
    [IO.File]::WriteAllText($generatorOutputPath, 'do-not-overwrite', $utf8)
    # 服务端按每条文本是否含中文选择 60/120 的上限；生成器按 deck.language 统一选择。
    # 为只测 result_analysis 对象兼容性，使用英文副本使两端限额一致。
    $generatorDeck = $deck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $generatorDeck.language = 'en'
    $generatorDeck | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $generatorDeckPath -Encoding UTF8
    $previousErrorActionPreference = $ErrorActionPreference
    try {
        # Native stderr 在 $ErrorActionPreference=Stop 下会变成 PowerShell 异常；预期的
        # 非零预检退出码要作为断言输入，而不是让测试在此处中断。
        $ErrorActionPreference = 'Continue'
        $generatorOutput = & powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $generatorPath -DeckSpecPath $generatorDeckPath -OutputPath $generatorOutputPath 2>&1
        $generatorExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
    Assert-Condition ($generatorExitCode -ne 0) 'Generator preflight must stop for the deliberately existing output file.'
    $generatorText = @($generatorOutput | ForEach-Object { [string]$_ }) -join "`n"
    Assert-Condition ($generatorText -notmatch 'result_analysis must be a string|result_analysis must be a string or JSON array') 'Generator must accept the service-generated result_analysis object before validating output paths.'
    # 输入预检和路径验证均在启动 COM 前；不同 PowerPoint/安全路径环境会给出不同的后续
    # 错误文本，因此这里只验证没有回退到旧的 result_analysis 类型错误。
} finally {
    if (Test-Path -LiteralPath $generatorDirectory) { Remove-Item -LiteralPath $generatorDirectory -Recurse -Force }
}

# 2) 删除一个必备页必须导致审核硬失败，不能把不完整的组会稿送往 PowerPoint。
$brokenDeck = $deck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
$brokenDeck.slides = @($brokenDeck.slides | Where-Object { $_.section -ne 'innovation' })
$brokenAudit = Invoke-McpRequest (New-McpToolRequest -Id 200 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $brokenDeck })
Assert-Condition (-not $brokenAudit.pass) 'Audit must fail when a required section is removed.'
Assert-Condition (@($brokenAudit.findings | Where-Object { $_.severity -eq 'hard' -and $_.category -eq 'required-section' -and $_.issue -match 'innovation' }).Count -gt 0) 'Audit must report the missing innovation section as a hard finding.'
Assert-Condition (-not $brokenAudit.quality.safe_to_generate -and $brokenAudit.quality.hard_finding_count -gt 0) 'Audit quality JSON must mark a blocked deck as unsafe to generate.'

# 生成工具必须重复执行审核，不能因为调用方跳过 audit 就启动 PowerPoint。
$blockedOutputPath = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club-blocked-$([Guid]::NewGuid().ToString('N')).pptx"
$blockedGeneration = Invoke-McpResponse (New-McpToolRequest -Id 250 -ToolName 'generate_editable_pptx' -Arguments @{ deck_spec = $brokenDeck; output_path = $blockedOutputPath })
Assert-Condition ($null -ne $blockedGeneration.error -and $blockedGeneration.error.message -match 'mandatory content audit') 'Generation must block a deck with missing required content before PowerPoint is invoked.'
Assert-Condition (-not (Test-Path -LiteralPath $blockedOutputPath)) 'Blocked generation must not create a PPTX output file.'

# claim id 存在但其内部证据被篡改时，同样必须触发可追溯性硬失败。
$traceabilityBrokenDeck = $deck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
$traceabilityBrokenDeck.evidence_pack.claims[0].evidence[0].section_id = 'nonexistent-section'
$traceabilityAudit = Invoke-McpRequest (New-McpToolRequest -Id 275 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $traceabilityBrokenDeck })
Assert-Condition (-not $traceabilityAudit.pass) 'Audit must fail when a claim evidence section id is invalid.'
Assert-Condition (@($traceabilityAudit.findings | Where-Object { $_.severity -eq 'hard' -and $_.category -eq 'traceability' -and $_.issue -match 'unknown evidence section' }).Count -gt 0) 'Audit must report the nested claim-evidence traceability failure.'

# 3) required_sections 可配置；在缩短汇报的场景下审核应仅要求调用方声明的模块。
$customDeck = Invoke-DesignFromPaper $englishFixture 300 @('background', 'methods')
Assert-Condition (($customDeck.required_sections -join '|') -eq 'background|methods') 'Custom required_sections must be preserved in the deck spec.'
$customAudit = Invoke-McpRequest (New-McpToolRequest -Id 400 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $customDeck })
Assert-Condition ($customAudit.pass) 'Complete fixture must pass under a smaller valid required_sections set.'
Assert-Condition ($customAudit.quality.safe_to_generate -and @($customAudit.quality.required_section_coverage).Count -eq 2) 'Audit quality JSON must report custom required-section coverage.'

# 4) 中文标题、句号、图号和结论关键词也要能走完整证据链。
$chineseDeck = Invoke-DesignFromPaper $chineseFixture 500
$chineseDataSlide = @($chineseDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
Assert-Condition ($chineseDataSlide.Count -eq 1) 'Chinese fixture must generate an experimental-data slide.'
Assert-Condition (@($chineseDataSlide[0].source_figure_ids).Count -gt 0) 'Chinese figure references must be retained on experimental-data slides.'
$chineseAudit = Invoke-McpRequest (New-McpToolRequest -Id 600 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $chineseDeck })
if (-not $chineseAudit.pass) {
    # 仅在失败时输出完整审核结果，便于定位 Windows 编码或证据回链回归。
    Write-Output ($chineseDeck | ConvertTo-Json -Depth 100)
    Write-Output ($chineseAudit | ConvertTo-Json -Depth 100)
}
Assert-Condition ($chineseAudit.pass) 'Chinese fixture must pass mandatory-section audit.'

# 4a) 表格引用被单独识别，但在没有用户确认的裁剪 PNG 时不能自动插入。这个测试不依赖
# WinRT 渲染器，覆盖“有 Table 不等于可以把整页 PDF 当作数据表”的设计边界。
$tableEvidence = $deck.evidence_pack | ConvertTo-Json -Depth 100 | ConvertFrom-Json
$tableEvidence.tables = @([pscustomobject]@{
    id = 'table-1'
    label = 'Table 1'
    context = 'Table 1 summarizes the primary experimental comparison.'
    source_page = 1
    source_pages = @(1)
    selection_mode = 'user-confirmed-page-crop-required'
})
$tableEvidence.claims[0].text = 'Table 1 reports that the intervention improves the primary experimental outcome.'
$tableEvidence.claims[0].evidence[0].page_number = 1
$tableDeck = Invoke-McpRequest (New-McpToolRequest -Id 625 -ToolName 'design_journal_club_deck' -Arguments @{
    evidence_pack = $tableEvidence
    duration_minutes = 5
    language = 'en'
    audience = 'lab'
    required_sections = @('experimental_data', 'methods')
})
$tableResultSlide = @($tableDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
Assert-Condition ($tableResultSlide.Count -eq 1 -and $tableResultSlide[0].suggested_figure_id -eq 'table-1') 'A result claim may trace to an identified source table.'
Assert-Condition (-not $tableResultSlide[0].PSObject.Properties['suggested_image_path']) 'An unreviewed table page must not be inserted automatically.'

# 4b) 文本阅读性与叙事顺序都是硬门禁：不能让“标题像提问、实验在方法前、结论夹在中间”的
# deck spec 进入 PowerPoint 生成阶段。
$punctuationDeck = $deck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
$punctuationDeck.slides[1].title = '研究问题：这篇论文解决了什么问题？'
$punctuationAuditResponse = Invoke-McpResponse (New-McpToolRequest -Id 650 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $punctuationDeck })
Assert-Condition ($null -ne $punctuationAuditResponse.error -and $punctuationAuditResponse.error.message -match 'colon or question mark') 'Heading punctuation must be rejected before audit generation.'

$missingTakeawayDeck = $deck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
$missingTakeawaySlide = @($missingTakeawayDeck.slides | Where-Object { $_.kind -ne 'title' } | Select-Object -First 1)[0]
[void]$missingTakeawaySlide.PSObject.Properties.Remove('takeaway')
$missingTakeawayAudit = Invoke-McpRequest (New-McpToolRequest -Id 6501 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $missingTakeawayDeck })
Assert-Condition (-not $missingTakeawayAudit.pass -and @($missingTakeawayAudit.findings | Where-Object { $_.category -eq 'narrative' -and $_.severity -eq 'hard' -and $_.issue -eq 'Missing takeaway' }).Count -eq 1) 'Audit must hard-block a non-title slide without a takeaway, just like the generator.'

$orderDeck = $deck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
$methodSlide = @($orderDeck.slides | Where-Object { $_.section -eq 'methods' } | Select-Object -First 1)[0]
$resultSlide = @($orderDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)[0]
$methodIndex = [array]::IndexOf(@($orderDeck.slides), $methodSlide)
$resultIndex = [array]::IndexOf(@($orderDeck.slides), $resultSlide)
$orderDeck.slides[$methodIndex] = $resultSlide
$orderDeck.slides[$resultIndex] = $methodSlide
$orderAudit = Invoke-McpRequest (New-McpToolRequest -Id 651 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $orderDeck })
Assert-Condition (-not $orderAudit.pass) 'Audit must block a deck that places experimental evidence before methods.'
Assert-Condition (@($orderAudit.findings | Where-Object { $_.severity -eq 'hard' -and $_.category -eq 'narrative-order' }).Count -gt 0) 'Audit must report a narrative-order finding.'

# 5) 一图号 + 一张合规图片时，方法页和实验数据页应自动插图；多个候选或没有候选时
# 绝不能任取第一张图片。图片夹具位于插件专用临时根，不改变真实用户资料目录边界。
$figureFixtureRoot = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\figure-mapping-test-$([Guid]::NewGuid().ToString('N'))"
$firstFigurePath = Join-Path $figureFixtureRoot 'figure-1.png'
$secondFigurePath = Join-Path $figureFixtureRoot 'figure-2.png'
try {
    New-Item -ItemType Directory -Path $figureFixtureRoot -Force | Out-Null
    # 经过验证的 1×1 PNG；这里只验证安全的图像关联逻辑，不依赖 PowerPoint 或外部 PDF。
    $pngBytes = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL9ZwAAAABJRU5ErkJggg==')
    [IO.File]::WriteAllBytes($firstFigurePath, $pngBytes)
    [IO.File]::WriteAllBytes($secondFigurePath, $pngBytes)

    $automaticEvidence = $deck.evidence_pack | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $automaticEvidence.extraction.asset_directory = $figureFixtureRoot
    $automaticEvidence.extraction.pages = @([pscustomobject]@{
        page_number = 1
        excerpt = 'Methods and results. Fig. 1 shows the system and experimental result.'
        assets = @([pscustomobject]@{ id = 'page-01-image-01'; page_number = 1; path = $firstFigurePath; bytes = $pngBytes.Length })
    })
    $automaticEvidence.figures = @([pscustomobject]@{
        id = 'fig-1'
        label = 'Fig. 1'
        # 合成夹具显式提供系统图图注，覆盖新的保守角色判定：只有明确的架构/流程信号
        # 才能进入方法页自动插图。
        context = 'Fig. 1: Overview of the proposed system architecture and workflow.'
        source_page = 1
        source_pages = @(1)
        figure_asset_candidates = @($firstFigurePath)
        asset_match = 'same-page-single-raster'
        automatic_image_path = $firstFigurePath
        # 图号自动匹配不是只有一个裸路径；必须保留图号、提取资产编号、PDF 页码和匹配模式，
        # 使后续 PPT 结果可以追溯到 evidence pack 中的具体原始图片。
        automatic_binding = [pscustomobject]@{
            figure_id = 'fig-1'
            asset_id = 'page-01-image-01'
            source_page = 1
            path = $firstFigurePath
            mode = 'automatic-figure-number-single-raster'
        }
    })
    $automaticEvidence.figures | Add-Member -NotePropertyName 'journal_club_role' -NotePropertyValue 'methods-system' -Force
    $automaticEvidence.figures | Add-Member -NotePropertyName 'eligible_for_journal_club_visual' -NotePropertyValue $true -Force
    $automaticEvidence.figures | Add-Member -NotePropertyName 'selection_guidance' -NotePropertyValue 'Synthetic system visual for test.' -Force
    $automaticEvidence.figures | Add-Member -NotePropertyName 'classification_evidence' -NotePropertyValue ([pscustomobject]@{ source_page = 1; excerpt = 'Fig. 1: Overview of the proposed system architecture and workflow.'; basis = 'caption-or-nearby-text-keywords' }) -Force
    $methodsSection = @($automaticEvidence.sections | Where-Object { $_.title -match '(?i)method' } | Select-Object -First 1)
    Assert-Condition ($methodsSection.Count -eq 1) 'Fixture must contain a methods section for system-figure mapping.'
    $methodsSection[0].excerpt = 'The system architecture and workflow are shown in Fig. 1. Input samples enter the processing module before validation. The validation stage reports the final output.'
    $methodsSection[0].source_page = 1
    $automaticEvidence.claims[0].text = 'Fig. 1 reports the primary experimental performance of the proposed approach.'
    $firstClaimEvidence = @($automaticEvidence.claims[0].evidence | Select-Object -First 1)
    Assert-Condition ($firstClaimEvidence.Count -eq 1) 'Fixture must provide source evidence for the first result claim.'
    $firstClaimEvidence[0].page_number = 1

    $automaticDeck = Invoke-McpRequest (New-McpToolRequest -Id 700 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $automaticEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
    $automaticMethodsSlide = @($automaticDeck.slides | Where-Object { $_.section -eq 'methods' } | Select-Object -First 1)
    $automaticResultSlide = @($automaticDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    $automaticFigure = @($automaticDeck.evidence_pack.figures | Where-Object { $_.id -eq 'fig-1' } | Select-Object -First 1)
    $automaticBinding = $automaticFigure[0].automatic_binding
    Assert-Condition ($automaticFigure.Count -eq 1 -and $null -ne $automaticBinding) 'A unique Fig. 1 must retain a structured automatic_binding in the returned evidence pack.'
    Assert-Condition ($automaticBinding.figure_id -eq 'fig-1' -and $automaticBinding.asset_id -eq 'page-01-image-01' -and $automaticBinding.source_page -eq 1 -and $automaticBinding.path -eq $firstFigurePath -and $automaticBinding.mode -eq 'automatic-figure-number-single-raster') 'Automatic Fig. 1 binding must preserve its figure id, asset id, source page, path, and high-confidence mode.'
    Assert-Condition ($automaticMethodsSlide.Count -eq 1 -and $automaticMethodsSlide[0].suggested_image_path -eq $firstFigurePath) 'A unique system/method figure must be automatically inserted into the methods slide spec.'
    Assert-Condition ($automaticResultSlide.Count -eq 1 -and -not $automaticResultSlide[0].PSObject.Properties['suggested_image_path']) 'A methods-system figure must not be reused as an experimental-result visual.'
    Assert-Condition ($automaticMethodsSlide[0].source_asset_id -eq 'page-01-image-01') 'An automatically bound methods figure must expose its extracted source_asset_id.'
    Assert-Condition ([string]$automaticResultSlide[0].source_asset_id -eq '') 'A result slide without a result/ablation visual must not claim a source asset id.'
    Assert-Condition (@($automaticMethodsSlide[0].figure_asset_candidates).Count -eq 1) 'Automatic image paths must remain traceable through figure_asset_candidates.'
    Assert-Condition ($automaticMethodsSlide[0].visual_role -eq 'system-architecture') 'A method figure must use the dedicated system-architecture layout.'
    Assert-Condition (@($automaticMethodsSlide[0].explanation_points).Count -ge 2) 'A system-architecture slide must include at least two paper-backed explanation points.'
    Assert-Condition (@($automaticMethodsSlide[0].explanation_points | Where-Object { [string]$_.source_section_id -eq [string]$methodsSection[0].id }).Count -eq @($automaticMethodsSlide[0].explanation_points).Count) 'Every system-architecture explanation point must retain a traceable Methods section id.'
    Assert-Condition (@($automaticMethodsSlide[0].explanation_points | Where-Object { [string]$_.source_figure_id -eq 'fig-1' }).Count -ge 1) 'System-architecture explanation must retain its Figure id when a system figure is selected.'
    Assert-Condition (@($automaticMethodsSlide[0].explanation_points | Where-Object { $_.stage -in @('input', 'module', 'output-validation') }).Count -ge 1) 'System-architecture explanation must organize the paper text around input, modules, or output/validation.'
    Assert-Condition ($automaticResultSlide[0].visual_role -eq 'result-analysis') 'A result slide without an eligible result/ablation visual must remain text-first.'
    Assert-Condition ($automaticMethodsSlide[0].visual_role_source -eq 'methods-system' -and $automaticMethodsSlide[0].visual_source_page -eq 1 -and $automaticMethodsSlide[0].visual_evidence.excerpt -match 'architecture') 'Methods visual selection must preserve role, source page, and caption evidence.'

    # 即使手工编辑 deck spec 把系统页改成 generic，只要仍插入方法图，审核器也必须要求
    # 可回链的讲解点，不能让“只放一张图”绕过结构审查。
    $missingExplanationDeck = $automaticDeck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $missingExplanationSlide = @($missingExplanationDeck.slides | Where-Object { $_.section -eq 'methods' } | Select-Object -First 1)[0]
    $missingExplanationSlide.visual_role = 'generic'
    [void]$missingExplanationSlide.PSObject.Properties.Remove('explanation_points')
    $missingExplanationAudit = Invoke-McpRequest (New-McpToolRequest -Id 705 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $missingExplanationDeck })
    Assert-Condition (-not $missingExplanationAudit.pass -and @($missingExplanationAudit.findings | Where-Object { $_.category -eq 'methods-explanation' -and $_.severity -eq 'hard' }).Count -eq 1) 'A methods image without an explanation must remain a hard audit failure.'

    $unlinkedExplanationDeck = $automaticDeck | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $unlinkedExplanationSlide = @($unlinkedExplanationDeck.slides | Where-Object { $_.section -eq 'methods' } | Select-Object -First 1)[0]
    $unlinkedExplanationSlide.explanation_points[0].source_section_id = 'introduction-2'
    $unlinkedExplanationAudit = Invoke-McpRequest (New-McpToolRequest -Id 706 -ToolName 'audit_journal_club_deck' -Arguments @{ deck_spec = $unlinkedExplanationDeck })
    Assert-Condition (-not $unlinkedExplanationAudit.pass -and @($unlinkedExplanationAudit.findings | Where-Object { $_.category -eq 'methods-explanation' -and $_.issue -match 'not linked to a methods section' }).Count -eq 1) 'A system explanation must link to a methods section, not merely any paper text.'

    $ambiguousEvidence = $automaticEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $ambiguousFigure = $ambiguousEvidence.figures[0]
    $ambiguousFigure.figure_asset_candidates = @($firstFigurePath, $secondFigurePath)
    $ambiguousFigure.asset_match = 'ambiguous'
    # 不可信/手工编辑的 evidence pack 即使伪造 automatic_image_path，也不能绕过
    # “一图号 + 一张候选图”的自动选择条件。
    $ambiguousFigure | Add-Member -NotePropertyName 'automatic_image_path' -NotePropertyValue $firstFigurePath -Force
    [void]$ambiguousFigure.PSObject.Properties.Remove('automatic_binding')
    $ambiguousDeck = Invoke-McpRequest (New-McpToolRequest -Id 710 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $ambiguousEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
    $ambiguousMethodsSlide = @($ambiguousDeck.slides | Where-Object { $_.section -eq 'methods' } | Select-Object -First 1)
    $ambiguousResultSlide = @($ambiguousDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    $ambiguousReturnedFigure = @($ambiguousDeck.evidence_pack.figures | Where-Object { $_.id -eq 'fig-1' } | Select-Object -First 1)
    Assert-Condition ($ambiguousReturnedFigure.Count -eq 1 -and $null -eq $ambiguousReturnedFigure[0].PSObject.Properties['automatic_binding']) 'An ambiguous page must not retain or synthesize an automatic_binding.'
    Assert-Condition (-not $ambiguousMethodsSlide[0].PSObject.Properties['suggested_image_path']) 'Ambiguous method figures must not be inserted automatically.'
    Assert-Condition (-not $ambiguousResultSlide[0].PSObject.Properties['suggested_image_path']) 'Ambiguous experimental figures must not be inserted automatically.'
    Assert-Condition ([string]$ambiguousMethodsSlide[0].source_asset_id -eq '' -and [string]$ambiguousResultSlide[0].source_asset_id -eq '') 'Ambiguous figures must not assign a source_asset_id to a slide.'

    # 图号匹配必须按完整编号处理：Fig. 1 不得误命中 Fig. 10；一句话同时出现两个
    # 图号时也不能擅自选择第一个。
    $numberBoundaryEvidence = $automaticEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $numberBoundaryEvidence.figures = @($numberBoundaryEvidence.figures) + @([pscustomobject]@{
        id = 'fig-10'
        label = 'Fig. 10'
        context = 'Fig. 10: Primary experimental performance comparison.'
        source_page = 2
        source_pages = @(2)
        figure_asset_candidates = @($secondFigurePath)
        asset_match = 'same-page-single-raster'
        automatic_image_path = $secondFigurePath
        automatic_binding = [pscustomobject]@{
            figure_id = 'fig-10'
            asset_id = 'page-02-image-01'
            source_page = 2
            path = $secondFigurePath
            mode = 'automatic-figure-number-single-raster'
        }
    })
    $numberBoundaryEvidence.figures[-1] | Add-Member -NotePropertyName 'journal_club_role' -NotePropertyValue 'main-result' -Force
    $numberBoundaryEvidence.figures[-1] | Add-Member -NotePropertyName 'eligible_for_journal_club_visual' -NotePropertyValue $true -Force
    $numberBoundaryEvidence.figures[-1] | Add-Member -NotePropertyName 'selection_guidance' -NotePropertyValue 'Synthetic main-result visual for test.' -Force
    $numberBoundaryEvidence.figures[-1] | Add-Member -NotePropertyName 'classification_evidence' -NotePropertyValue ([pscustomobject]@{ source_page = 2; excerpt = 'Fig. 10: Primary experimental performance comparison.'; basis = 'caption-or-nearby-text-keywords' }) -Force
    $numberBoundaryEvidence.extraction.pages = @($numberBoundaryEvidence.extraction.pages) + @([pscustomobject]@{
        page_number = 2
        excerpt = 'Fig. 10 contains the independent result.'
        assets = @([pscustomobject]@{ id = 'page-02-image-01'; page_number = 2; path = $secondFigurePath; bytes = $pngBytes.Length })
    })
    $numberBoundaryEvidence.claims[0].text = 'Fig. 10 shows the independent primary experimental outcome.'
    $numberBoundaryDeck = Invoke-McpRequest (New-McpToolRequest -Id 715 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $numberBoundaryEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
    $numberBoundaryResultSlide = @($numberBoundaryDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    $numberBoundaryFigure = @($numberBoundaryDeck.evidence_pack.figures | Where-Object { $_.id -eq 'fig-10' } | Select-Object -First 1)
    $numberBoundaryBinding = $numberBoundaryFigure[0].automatic_binding
    Assert-Condition ($numberBoundaryResultSlide[0].suggested_figure_id -eq 'fig-10' -and $numberBoundaryResultSlide[0].suggested_image_path -eq $secondFigurePath) 'Fig. 10 must not be confused with Fig. 1 during automatic mapping.'
    Assert-Condition ($numberBoundaryFigure.Count -eq 1 -and $null -ne $numberBoundaryBinding -and $numberBoundaryBinding.figure_id -eq 'fig-10' -and $numberBoundaryBinding.asset_id -eq 'page-02-image-01' -and $numberBoundaryBinding.source_page -eq 2 -and $numberBoundaryBinding.path -eq $secondFigurePath -and $numberBoundaryBinding.mode -eq 'automatic-figure-number-single-raster') 'Fig. 10 must retain its own automatic binding rather than inherit Fig. 1 metadata.'
    Assert-Condition ($numberBoundaryResultSlide[0].source_asset_id -eq 'page-02-image-01') 'Fig. 10 result slide must retain the exact extracted asset id.'
    Assert-Condition ($numberBoundaryResultSlide[0].source_visual_id -eq 'fig-10' -and $numberBoundaryResultSlide[0].visual_role_source -eq 'main-result' -and $numberBoundaryResultSlide[0].visual_source_page -eq 2 -and $numberBoundaryResultSlide[0].visual_evidence.excerpt -match 'performance') 'A result visual must preserve source id, role, page, and caption evidence.'

    $multipleMentionEvidence = $numberBoundaryEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $multipleMentionEvidence.claims[0].text = 'Fig. 1 and Fig. 10 jointly support the primary experimental outcome.'
    $multipleMentionDeck = Invoke-McpRequest (New-McpToolRequest -Id 716 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $multipleMentionEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
$multipleMentionResultSlide = @($multipleMentionDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
Assert-Condition (-not $multipleMentionResultSlide[0].PSObject.Properties['suggested_image_path']) 'A claim mentioning multiple figures must not choose an image automatically.'

    # 五分钟汇报也不能只挑“改善”结果。若论文同时报告无改善/较弱证据，唯一的结果页
    # 必须回链到两条 claim，并在限定条件中呈现反例。
    $mixedOutcomeEvidence = $automaticEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $mixedOutcomeEvidence.claims = @(
        [pscustomobject]@{
            id = 'claim-positive'
            text = 'Fig. 1 reports that the intervention improved the primary response relative to control.'
            evidence = @([pscustomobject]@{ section_id = 'results-4'; section_title = 'Results'; page_number = 1; excerpt = 'The intervention improved the primary response relative to control.' })
            confidence = 'needs-review'
            comparison_kind = 'reported-better'
        },
        [pscustomobject]@{
            id = 'claim-counter'
            text = 'A secondary endpoint showed no improvement relative to control.'
            evidence = @([pscustomobject]@{ section_id = 'results-4'; section_title = 'Results'; page_number = 1; excerpt = 'A secondary endpoint showed no improvement relative to control.' })
            confidence = 'needs-review'
            comparison_kind = 'reported-worse-or-no-benefit'
        }
    )
    $mixedOutcomeDeck = Invoke-McpRequest (New-McpToolRequest -Id 718 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $mixedOutcomeEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
    $mixedOutcomeSlide = @($mixedOutcomeDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    Assert-Condition (($mixedOutcomeSlide[0].source_claim_ids -join '|') -match 'claim-positive' -and ($mixedOutcomeSlide[0].source_claim_ids -join '|') -match 'claim-counter') 'A five-minute result slide must retain source ids for both positive and counter evidence.'
    Assert-Condition ($mixedOutcomeSlide[0].result_analysis.caveat -match 'no improvement') 'A five-minute result slide must mention the reported counter evidence rather than silently dropping it.'
    Assert-Condition ($mixedOutcomeSlide[0].result_analysis.comparison_source_claim_id -eq 'claim-positive' -and $mixedOutcomeSlide[0].result_analysis.comparison_source_excerpt -match 'improved') 'Result analysis must expose the source claim and original comparison text rather than infer it from the visual.'
    Assert-Condition ($mixedOutcomeSlide[0].result_analysis.caveat_source_claim_id -eq 'claim-counter' -and $mixedOutcomeSlide[0].result_analysis.caveat_source_excerpt -match 'no improvement') 'Result analysis must preserve a cited counterexample with its original-paper excerpt.'

    $noImageEvidence = $automaticEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $noImageFigure = $noImageEvidence.figures[0]
    $noImageFigure.figure_asset_candidates = @()
    $noImageFigure.asset_match = 'none'
    [void]$noImageFigure.PSObject.Properties.Remove('automatic_image_path')
    [void]$noImageFigure.PSObject.Properties.Remove('automatic_binding')
    $noImageDeck = Invoke-McpRequest (New-McpToolRequest -Id 720 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $noImageEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
    $noImageResultSlide = @($noImageDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    Assert-Condition (-not $noImageResultSlide[0].PSObject.Properties['suggested_image_path']) 'Papers without a usable figure must produce text-only result slides.'

    # 6) 图表角色是主线选择门禁。案例图、错误案例、相似样例与数据集统计都不能因存在
    # 合规 PNG 或手工选择路径而进入默认 PPT；主实验、消融和系统图则保留角色、页码和
    # 图注证据，供生成器与汇报文本共同追溯。
    $roleFixtureVisuals = @(
        [pscustomobject]@{ id = 'fig-methods'; context = 'Figure 2: An overview of our framework architecture and workflow.'; expected_role = 'methods-system'; expected_eligible = $true },
        [pscustomobject]@{ id = 'fig-result'; context = 'Figure 4: Effect of TopK on model performance.'; expected_role = 'main-result'; expected_eligible = $true },
        [pscustomobject]@{ id = 'fig-ablation'; context = 'Figure 5: Ablation studies for each component.'; expected_role = 'ablation'; expected_eligible = $true },
        [pscustomobject]@{ id = 'fig-case'; context = 'Figure 6: Examples of wrongly predicted memes and similar memes.'; expected_role = 'case-analysis'; expected_eligible = $false },
        [pscustomobject]@{ id = 'fig-unknown'; context = 'Figure 10.'; expected_role = 'excluded-or-unknown'; expected_eligible = $false },
        [pscustomobject]@{ id = 'table-result'; context = 'Table 1: Benchmark evaluation results on three datasets.'; expected_role = 'main-result'; expected_eligible = $true },
        [pscustomobject]@{ id = 'table-ablation'; context = 'Table 3: Ablation studies on the proposed framework.'; expected_role = 'ablation'; expected_eligible = $true },
        [pscustomobject]@{ id = 'table-statistics'; context = 'Table 4: Statistics of test sets and class distribution.'; expected_role = 'excluded-or-unknown'; expected_eligible = $false }
    )
    $roleEvidence = $automaticEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $roleEvidence.figures = @($roleFixtureVisuals | Where-Object { $_.id -like 'fig-*' } | ForEach-Object {
        [pscustomobject]@{ id = $_.id; context = $_.context; source_page = 1; source_pages = @(1) }
    })
    $roleEvidence.tables = @($roleFixtureVisuals | Where-Object { $_.id -like 'table-*' } | ForEach-Object {
        [pscustomobject]@{ id = $_.id; context = $_.context; source_page = 1; source_pages = @(1); selection_mode = 'user-confirmed-page-crop-required' }
    })
    $rolePack = Invoke-McpRequest (New-McpToolRequest -Id 722 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $roleEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
    foreach ($visual in $roleFixtureVisuals) {
        $collection = if ($visual.id -like 'table-*') { $rolePack.evidence_pack.tables } else { $rolePack.evidence_pack.figures }
        $classified = @($collection | Where-Object { $_.id -eq $visual.id } | Select-Object -First 1)
        Assert-Condition ($classified.Count -eq 1 -and $classified[0].journal_club_role -eq $visual.expected_role) "Visual role must be classified conservatively: $($visual.id)"
        Assert-Condition ([bool]$classified[0].eligible_for_journal_club_visual -eq [bool]$visual.expected_eligible) "Visual eligibility must match role policy: $($visual.id)"
    }

    # 当结果 claim 未明确提及图表且同时存在多张合格主实验/消融视觉时，不能把论文顺序
    # 最靠前的候选任意配到该 claim。只有唯一候选才允许保守回退；多候选必须等待图号
    # 或用户确认的选择，确保图表说明与论文原文真正对应。
    $ambiguousResultEvidence = $roleEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $ambiguousResultEvidence.claims[0].text = 'The intervention improved the primary outcome in the held-out evaluation.'
    $ambiguousResultEvidence.claims[0].evidence[0].excerpt = $ambiguousResultEvidence.claims[0].text
    $ambiguousResultEvidence.claims[0].evidence[0].page_number = 99
    $ambiguousResultDeck = Invoke-McpRequest (New-McpToolRequest -Id 724 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $ambiguousResultEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
    })
    $ambiguousResultSlide = @($ambiguousResultDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    Assert-Condition ($ambiguousResultSlide.Count -eq 1 -and @($ambiguousResultSlide[0].source_figure_ids).Count -eq 0) 'Multiple eligible visuals without an explicit Figure/Table mention must not be arbitrarily bound to a result claim.'

    $caseEvidence = $automaticEvidence | ConvertTo-Json -Depth 100 | ConvertFrom-Json
    $caseEvidence.figures[0].context = 'Figure 6: Examples of wrongly predicted memes and similar memes.'
    $caseEvidence.claims[0].text = 'Fig. 1 reports primary experimental performance.'
    $caseDeck = Invoke-McpRequest (New-McpToolRequest -Id 725 -ToolName 'design_journal_club_deck' -Arguments @{
        evidence_pack = $caseEvidence
        duration_minutes = 5
        language = 'en'
        audience = 'lab'
        required_sections = @('methods', 'experimental_data')
        figure_asset_selection = @{ 'fig-1' = $firstFigurePath }
    })
    $caseMethodsSlide = @($caseDeck.slides | Where-Object { $_.section -eq 'methods' } | Select-Object -First 1)
    $caseResultSlide = @($caseDeck.slides | Where-Object { $_.section -eq 'experimental_data' } | Select-Object -First 1)
    Assert-Condition (-not $caseMethodsSlide[0].PSObject.Properties['suggested_image_path'] -and -not $caseResultSlide[0].PSObject.Properties['suggested_image_path']) 'Case-analysis visuals must remain excluded even with a user-selected extracted image.'
    Assert-Condition (-not (@($caseDeck.slides | Where-Object { [string]$_.visual_role_source -eq 'case-analysis' }).Count -gt 0)) 'Case-analysis visuals must never be turned into a default journal-club methods or results narrative.'
    $caseRenderResponse = Invoke-McpResponse (New-McpToolRequest -Id 726 -ToolName 'render_paper_visual' -Arguments @{
        evidence_pack = $caseEvidence
        visual_id = 'fig-1'
        page_number = 1
    })
    Assert-Condition ($null -ne $caseRenderResponse.error -and $caseRenderResponse.error.message -match 'not eligible') 'Case-analysis visuals must be rejected before PDF rendering creates a temporary image.'
} finally {
    if (Test-Path -LiteralPath $figureFixtureRoot) {
        Remove-Item -LiteralPath $figureFixtureRoot -Recurse -Force
    }
}

Write-Output 'PASS: content-template-tests.ps1'
