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

# 1) MCP 协商必须支持当前协议，并向 Agent 暴露真实的 PowerPoint 能力边界。
$initialize = Invoke-McpRaw ([ordered]@{
    jsonrpc = '2.0'
    id = 1
    method = 'initialize'
    params = [ordered]@{ protocolVersion = '2025-06-18'; capabilities = [ordered]@{}; clientInfo = [ordered]@{ name = 'release-e2e-test'; version = '1.0.0' } }
})
Assert-Condition (-not $initialize.error) 'initialize must succeed.'
Assert-Condition ($initialize.result.protocolVersion -eq '2025-06-18') 'Server must negotiate MCP 2025-06-18.'
Assert-Condition ($initialize.result.serverInfo.version -eq '1.0.0') 'Server version must match the release manifest.'

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

if (-not $RunPowerPoint) {
    Write-Output 'PASS: end-to-end-tests.ps1 (MCP and content path)'
    exit 0
}

if (-not $status.powerpoint_com_registered) { throw 'PowerPoint COM is required when -RunPowerPoint is specified.' }
$workDirectory = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club-e2e-test-$([Guid]::NewGuid().ToString('N'))"
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
