<#
  Standalone native-renderer audit for a generated PPTX.
  It never mutates the deck; preview PNGs are written only when requested.
#>
param(
    [Parameter(Mandatory = $true)][string]$PresentationPath,
    [string]$PreviewDirectory = '',
    # 只有人工视觉复核时才导出逐页 PNG。
    [switch]$ExportPreviews
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'powerpoint-quality.ps1')
if (-not (Test-PowerPointComRegistration)) { throw 'Microsoft PowerPoint desktop is not registered on this computer.' }
if (-not (Test-Path -LiteralPath $PresentationPath)) { throw "Presentation was not found: $PresentationPath" }
$shouldExportPreviews = $ExportPreviews -or -not [string]::IsNullOrWhiteSpace($PreviewDirectory)
$resolvedPreviewDirectory = if ($shouldExportPreviews) {
    if ([string]::IsNullOrWhiteSpace($PreviewDirectory)) {
        Join-Path (Get-PaperToJournalClubTemporaryRoot) "audit-previews-$([Guid]::NewGuid().ToString('N'))"
    } else {
        [IO.Path]::GetFullPath($PreviewDirectory)
    }
} else {
    $null
}
if ($resolvedPreviewDirectory -and (Test-Path -LiteralPath $resolvedPreviewDirectory) -and $null -ne (Get-ChildItem -LiteralPath $resolvedPreviewDirectory -Force | Select-Object -First 1)) {
    throw "PreviewDirectory must be a new or empty directory: $resolvedPreviewDirectory"
}

$session = $null
try {
    $session = Open-PowerPointPresentationReadOnly -Path $PresentationPath -WithWindow:$shouldExportPreviews
    $result = Invoke-PowerPointQualityAudit -Presentation $session.presentation -ExportPreviews:$shouldExportPreviews -PreviewDirectory $resolvedPreviewDirectory
    # 作为生成器的独立审计工作进程时，结果必须是一行 JSON，避免调用方把内部 findings
    # 对象的 "{" 误当成完整响应；直接命令行使用时同样便于机器读取。
    $result | ConvertTo-Json -Depth 40 -Compress
    if (-not $result.pass) { exit 2 }
} finally {
    Close-PowerPointReadOnlySession $session
}
