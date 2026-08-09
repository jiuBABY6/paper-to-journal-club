<#
  PowerPoint COM generator for the editable journal-club deck.
  The PowerShell MCP service owns the journal-club contract. This file avoids
  non-ASCII string literals because Windows PowerShell 5.1 otherwise treats
  UTF-8 files without a BOM as the local ANSI code page.
#>
param(
    [Parameter(Mandatory = $true)][string]$DeckSpecPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$KeepOpen = "False",
    [string]$PreviewDirectory = "",
    [switch]$SkipPreviewExport,
    [switch]$Foreground,
    [switch]$DoNotFailOnAudit,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
. (Join-Path $PSScriptRoot 'powerpoint-quality.ps1')

function Convert-HexToOfficeRgb {
    param([string]$Hex)
    $clean = $Hex.Trim().TrimStart('#')
    if ($clean -notmatch '^[0-9a-fA-F]{6}$') { $clean = '114B5F' }
    $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Add-EditableText {
    param($Slide, [string]$Name, [string]$Text, [single]$Left, [single]$Top, [single]$Width, [single]$Height, [single]$FontSize, [int]$Color, [bool]$Bold = $false)
    # msoTextOrientationHorizontal = 1. Each block has a stable name for audit and repair.
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.TextFrame.TextRange.Text = $Text
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.AutoSize = 0
    $shape.TextFrame.MarginLeft = 2
    $shape.TextFrame.MarginRight = 2
    $shape.TextFrame.MarginTop = 1
    $shape.TextFrame.MarginBottom = 1
    $shape.TextFrame.TextRange.Font.Name = "Aptos"
    $shape.TextFrame.TextRange.Font.Size = $FontSize
    $shape.TextFrame.TextRange.Font.Color.RGB = $Color
    $shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    # TextFrame2 若保留默认 AutoSize，会在保存时把文本框高度缩到一行，引发文字重叠。
    try {
        $shape.TextFrame2.AutoSize = 0
        $shape.TextFrame2.WordWrap = -1
    } catch {
        # 旧版 PowerPoint 未公开 TextFrame2 时，仍可由 TextFrame 生成可编辑文本。
    }
    # PowerPoint 有时会在设定文本后把新建 TextBox 的高度自动缩成一行；先恢复调用方的画布几何。
    $shape.Left = $Left
    $shape.Top = $Top
    $shape.Width = $Width
    $shape.Height = $Height
    # 保持面向观众的最小字号；内容过长时交给最终审计报告，而不是把文字缩到不可读。
    $minimumFontSize = if ($FontSize -ge 35) { 35 } elseif ($FontSize -ge 24) { 24 } elseif ($FontSize -ge 16) { 16 } else { 8 }
    try {
        while ($shape.TextFrame2.TextRange.BoundHeight -gt ($shape.Height - 2) -and $shape.TextFrame.TextRange.Font.Size -gt $minimumFontSize) {
            $shape.TextFrame.TextRange.Font.Size = $shape.TextFrame.TextRange.Font.Size - 1
        }
    } catch {
        # The final native PowerPoint audit reports text-fit defects on legacy shape types.
    }
    # 字号调整也可能触发 AutoSize，因此在返回前再次固定边界。
    $shape.Left = $Left
    $shape.Top = $Top
    $shape.Width = $Width
    $shape.Height = $Height
    return $shape
}

function Add-InfoCard {
    param($Slide, [string]$Name, [string]$Label, [string]$Value, [single]$Top, [int]$Primary, [int]$Accent)
    # msoShapeRoundedRectangle = 5. The card and its text stay independently editable.
    $card = $Slide.Shapes.AddShape(5, 510, $Top, 405, 118)
    $card.Name = "$Name-card"
    $card.Fill.ForeColor.RGB = 16777215
    $card.Line.ForeColor.RGB = $Primary
    $card.Line.Weight = 1.25
    Add-EditableText -Slide $Slide -Name "$Name-label" -Text $Label -Left 533 -Top ($Top + 14) -Width 350 -Height 24 -FontSize 14 -Color $Accent -Bold $true | Out-Null
    Add-EditableText -Slide $Slide -Name "$Name-value" -Text $Value -Left 533 -Top ($Top + 42) -Width 350 -Height 58 -FontSize 16 -Color $Primary | Out-Null
}

function Add-AtomicFigureImage {
    param($Slide, [string]$Name, [string]$ImagePath, [int]$Primary)
    if (-not (Test-Path -LiteralPath $ImagePath -PathType Leaf)) { throw "Suggested figure image was not found: $ImagePath" }
    $image = $Slide.Shapes.AddPicture([IO.Path]::GetFullPath($ImagePath), $false, $true, 510, 135, -1, -1)
    $image.Name = $Name
    $image.LockAspectRatio = -1
    $maxWidth = 405
    $maxHeight = 230
    $scale = [Math]::Min($maxWidth / [double]$image.Width, $maxHeight / [double]$image.Height)
    $image.Width = [single]($image.Width * $scale)
    $image.Left = [single](510 + (($maxWidth - $image.Width) / 2))
    $image.Top = [single](135 + (($maxHeight - $image.Height) / 2))
    $image.Line.ForeColor.RGB = $Primary
    $image.Line.Weight = 1
    # Explicitly preserve the atomic-raster contract for later structural review.
    $image.AlternativeText = 'atomic_raster_unit=true; contains_reconstructable_content=false; source=tightly-cropped-paper-figure'
    return $image
}

if (-not (Test-Path -LiteralPath $DeckSpecPath)) { throw "Deck spec was not found: $DeckSpecPath" }
# PowerShell 5.1 must be told the JSON encoding explicitly.
$deck = Get-Content -Raw -Encoding UTF8 -LiteralPath $DeckSpecPath | ConvertFrom-Json
$slideSpecs = @($deck.slides)
$slideTotal = $slideSpecs.Count
if ($slideTotal -lt 1) { throw "Deck spec must contain at least one slide." }
$isChinese = ([string]$deck.language) -match '^(?i:zh)(-|$)'
$labels = if ($isChinese) {
    [pscustomobject]@{
        title_subtitle = '组会汇报 | 可编辑演示文稿草稿'
        evidence_label = '证据 / 图示'
        figure_prefix = '原子化论文图'
        figure_placeholder = '请在此添加流程图、方法示意图或讨论图。'
        evidence_prefix = '证据'
        evidence_placeholder = '汇报前请补充来源页码或图号。'
    }
} else {
    [pscustomobject]@{
        title_subtitle = 'Journal Club | Editable presentation draft'
        evidence_label = 'Evidence / figure'
        figure_prefix = 'Atomic source figure'
        figure_placeholder = 'Add a workflow, method diagram, or discussion graphic here.'
        evidence_prefix = 'Evidence'
        evidence_placeholder = 'add source page or figure number before presenting.'
    }
}
$outputFull = [IO.Path]::GetFullPath($OutputPath)
if ([IO.Path]::GetExtension($outputFull).ToLowerInvariant() -ne '.pptx') { throw 'OutputPath must end with .pptx.' }
$outputDirectory = Split-Path -Parent $outputFull
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
if ((Test-Path -LiteralPath $outputFull -PathType Leaf) -and -not $Overwrite) {
    throw "Output PPTX already exists. Pass -Overwrite only after explicitly choosing to replace it: $outputFull"
}

if (-not (Test-PowerPointComRegistration)) { throw 'Microsoft PowerPoint desktop is not registered on this computer.' }
if (-not $deck.paper -or -not $deck.paper.title) { throw 'Deck spec must include paper.title.' }
$requiredSections = @(if ($deck.required_sections) { $deck.required_sections } else { @('background', 'innovation', 'methods', 'experimental_data', 'limitations', 'future_directions') })
$presentSections = @($slideSpecs | ForEach-Object { [string]$_.section })
$missingSections = @($requiredSections | Where-Object { $_ -notin $presentSections })
if ($missingSections.Count -gt 0) { throw "Deck spec is missing mandatory journal-club sections: $($missingSections -join ', ')" }
if (@($slideSpecs | Where-Object { $_.kind -eq 'title' }).Count -eq 0) { throw 'Deck spec must include a title slide.' }

$primary = Convert-HexToOfficeRgb $deck.theme.primary
$accent = Convert-HexToOfficeRgb $deck.theme.accent
$background = Convert-HexToOfficeRgb $deck.theme.background
$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    if (-not $Foreground) {
        try { $ppt.WindowState = 2 } catch { }
    }
    $presentation = $ppt.Presentations.Add()
    # 16:9 slide dimensions, in points.
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540

    $index = 0
    foreach ($slideSpec in $slideSpecs) {
        $index++
        # ppLayoutBlank = 12. Blank slides make the object layout deterministic.
        $slide = $presentation.Slides.Add($index, 12)
        $slide.FollowMasterBackground = 0
        $slide.Background.Fill.ForeColor.RGB = $background

        $isTitle = $slideSpec.kind -eq "title"
        if ($isTitle) {
            $band = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
            $band.Name = "slide-$index-background"
            $band.Fill.ForeColor.RGB = $primary
            $band.Line.Visible = 0
            Add-EditableText -Slide $slide -Name "slide-$index-title" -Text $slideSpec.title -Left 78 -Top 118 -Width 800 -Height 160 -FontSize 54 -Color 16777215 -Bold $true | Out-Null
            $titleSubtitle = if ($slideSpec.subtitle) { [string]$slideSpec.subtitle } else { $labels.title_subtitle }
            Add-EditableText -Slide $slide -Name "slide-$index-subtitle" -Text $titleSubtitle -Left 80 -Top 286 -Width 700 -Height 38 -FontSize 24 -Color 16777215 | Out-Null
        } else {
            Add-EditableText -Slide $slide -Name "slide-$index-title" -Text $slideSpec.title -Left 48 -Top 28 -Width 855 -Height 64 -FontSize 36 -Color $primary -Bold $true | Out-Null
            $line = $slide.Shapes.AddShape(1, 48, 92, 88, 5)
            $line.Name = "slide-$index-title-accent"
            $line.Fill.ForeColor.RGB = $accent
            $line.Line.Visible = 0

            Add-EditableText -Slide $slide -Name "slide-$index-takeaway" -Text $slideSpec.takeaway -Left 52 -Top 123 -Width 405 -Height 160 -FontSize 24 -Color $primary -Bold $true | Out-Null
            $bullets = @($slideSpec.bullets)
            $bulletText = if ($bullets.Count) { ($bullets | ForEach-Object { "- $_" }) -join "`r" } else { "- Add evidence or discussion points for this slide." }
            Add-EditableText -Slide $slide -Name "slide-$index-bullets" -Text $bulletText -Left 55 -Top 305 -Width 395 -Height 170 -FontSize 16 -Color $primary | Out-Null

            $imagePath = [string]$slideSpec.suggested_image_path
            if (-not [string]::IsNullOrWhiteSpace($imagePath) -and (Test-Path -LiteralPath $imagePath -PathType Leaf)) {
                Add-AtomicFigureImage -Slide $slide -Name "slide-$index-figure-image" -ImagePath $imagePath -Primary $primary | Out-Null
                $figureLabel = if ($slideSpec.figure_label) { [string]$slideSpec.figure_label } else { "$($labels.figure_prefix): $($slideSpec.suggested_figure_id)" }
                Add-EditableText -Slide $slide -Name "slide-$index-figure-label" -Text $figureLabel -Left 512 -Top 372 -Width 390 -Height 20 -FontSize 11 -Color $primary | Out-Null
            } else {
                $figureLabel = if ($slideSpec.suggested_figure_id) { "$($labels.figure_prefix): $($slideSpec.suggested_figure_id)" } else { $labels.figure_placeholder }
                if ($slideSpec.figure_label) { $figureLabel = [string]$slideSpec.figure_label }
                $evidenceLabel = if ($slideSpec.evidence_label) { [string]$slideSpec.evidence_label } else { $labels.evidence_label }
                Add-InfoCard -Slide $slide -Name "slide-$index-evidence" -Label $evidenceLabel -Value $figureLabel -Top 135 -Primary $primary -Accent $accent
            }
            $sourceText = if (@($slideSpec.source_claim_ids).Count) { "$($labels.evidence_prefix): $($slideSpec.source_claim_ids -join ', ')" } else { "$($labels.evidence_prefix): $($labels.evidence_placeholder)" }
            if ($slideSpec.source_text) { $sourceText = [string]$slideSpec.source_text }
            $sourceTop = if (-not [string]::IsNullOrWhiteSpace($imagePath) -and (Test-Path -LiteralPath $imagePath -PathType Leaf)) { 400 } else { 285 }
            Add-EditableText -Slide $slide -Name "slide-$index-source" -Text $sourceText -Left 512 -Top $sourceTop -Width 390 -Height 50 -FontSize 12 -Color $primary | Out-Null
        }
        Add-EditableText -Slide $slide -Name "slide-$index-footer" -Text "$($deck.paper.title) | $index / $slideTotal" -Left 48 -Top 503 -Width 860 -Height 18 -FontSize 9 -Color $primary | Out-Null
    }
    # ppSaveAsOpenXMLPresentation = 24; output remains editable in PowerPoint.
    $presentation.SaveAs($outputFull, 24)
    if (-not (Test-Path -LiteralPath $outputFull)) { throw "PowerPoint did not write the expected PPTX: $outputFull" }
    $resolvedPreviewDirectory = if ([string]::IsNullOrWhiteSpace($PreviewDirectory)) { "$outputFull.previews" } else { [IO.Path]::GetFullPath($PreviewDirectory) }
    $quality = Invoke-PowerPointQualityAudit -Presentation $presentation -PreviewDirectory $resolvedPreviewDirectory -ExportPreviews:(-not $SkipPreviewExport)
    if (-not $quality.pass -and -not $DoNotFailOnAudit) {
        throw "PowerPoint quality audit failed: $($quality.findings | ConvertTo-Json -Compress)"
    }
    Write-Output (ConvertTo-Json @{ output_path = $outputFull; slide_count = $slideTotal; editable_contract = "native-text-shapes"; quality_audit = $quality; preview_directory = if ($SkipPreviewExport) { $null } else { $resolvedPreviewDirectory } } -Depth 40 -Compress)
} finally {
    if ($presentation -and $KeepOpen -ne "True") { $presentation.Close() }
    if ($ppt -and $KeepOpen -ne "True") { $ppt.Quit() }
    if ($presentation) { [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null }
    if ($ppt) { [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
