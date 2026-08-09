<#
  Standalone native-renderer audit for a generated PPTX.
  It never mutates the deck; preview PNGs are written only when requested.
#>
param(
    [Parameter(Mandatory = $true)][string]$PresentationPath,
    [string]$PreviewDirectory = ''
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'powerpoint-quality.ps1')
if (-not (Test-PowerPointComRegistration)) { throw 'Microsoft PowerPoint desktop is not registered on this computer.' }
if (-not (Test-Path -LiteralPath $PresentationPath)) { throw "Presentation was not found: $PresentationPath" }
$resolvedPreviewDirectory = if ([string]::IsNullOrWhiteSpace($PreviewDirectory)) {
    Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\audit-previews-$([Guid]::NewGuid().ToString('N'))"
} else {
    [IO.Path]::GetFullPath($PreviewDirectory)
}
if ((Test-Path -LiteralPath $resolvedPreviewDirectory) -and $null -ne (Get-ChildItem -LiteralPath $resolvedPreviewDirectory -Force | Select-Object -First 1)) {
    throw "PreviewDirectory must be a new or empty directory: $resolvedPreviewDirectory"
}

$session = $null
try {
    $session = Open-PowerPointPresentationReadOnly -Path $PresentationPath
    $result = Invoke-PowerPointQualityAudit -Presentation $session.presentation -ExportPreviews -PreviewDirectory $resolvedPreviewDirectory
    $result | ConvertTo-Json -Depth 40
    if (-not $result.pass) { exit 2 }
} finally {
    Close-PowerPointReadOnlySession $session
}
