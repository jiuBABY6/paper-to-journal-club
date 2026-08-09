<# 运行无需 Node.js 的本地演示。 #>
param(
    [string]$PaperPath = '',
    [string]$DeckSpecPath = ''
)

$examplesDirectory = Join-Path $PSScriptRoot '..\examples'
if ([string]::IsNullOrWhiteSpace($PaperPath)) { $PaperPath = Join-Path $examplesDirectory 'sample-paper.md' }
if ([string]::IsNullOrWhiteSpace($DeckSpecPath)) {
    $DeckSpecPath = Join-Path $examplesDirectory ("sample-deck-spec-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
& (Join-Path $PSScriptRoot "paper-to-journal-club-server.ps1") -Demo -DemoInputPath $PaperPath -DemoOutputPath $DeckSpecPath
