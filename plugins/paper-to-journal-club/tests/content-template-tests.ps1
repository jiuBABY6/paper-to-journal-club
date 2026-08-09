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
Assert-Condition ($designTool.Count -eq 1 -and $null -ne $designTool[0].inputSchema.properties.required_sections) 'Design schema must expose required_sections.'
Assert-Condition ($analyseTool.Count -eq 1 -and $null -ne $analyseTool[0].inputSchema.properties.asset_output_dir) 'Analyse schema must expose optional asset_output_dir.'

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
$futureSlide = @($deck.slides | Where-Object { $_.section -eq 'future_directions' } | Select-Object -First 1)
Assert-Condition ($futureSlide[0].content_mode -eq 'presenter-discussion') 'Fixture without author future wording should use an explicitly labelled presenter discussion.'

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

Write-Output 'PASS: content-template-tests.ps1'
