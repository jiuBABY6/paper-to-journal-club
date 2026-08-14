<#
  PowerPoint COM generator for the editable journal-club deck.
  The PowerShell MCP service owns the journal-club contract. This file avoids
  non-ASCII string literals because Windows PowerShell 5.1 otherwise treats
  UTF-8 files without a BOM as the local ANSI code page.
#>
param(
    [Parameter(Mandatory = $true)][string]$DeckSpecPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [switch]$KeepOpen,
    [string]$PreviewDirectory = "",
    # 预览为显式选择；避免命令行直调时自动在 PPTX 旁生成 PNG 目录。
    [switch]$ExportPreviews,
    # 保留旧参数，确保旧自动化脚本仍可明确关闭预览。它优先于 ExportPreviews。
    [switch]$SkipPreviewExport,
    [switch]$Foreground,
    [switch]$DoNotFailOnAudit,
    # 仅供 MCP 父进程使用：在生成器进程退出后由未接触 Office COM 的宿主重新打开成品审核。
    [switch]$DeferQualityAudit,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
. (Join-Path $PSScriptRoot 'powerpoint-quality.ps1')
$MaximumDeckSpecBytes = 2MB
$MaximumDeckSlides = 30
$MaximumTextCharactersPerSlide = 8000
$MaximumBulletsPerSlide = 3
$MaximumNonTitleSlideTitleCharacters = 64
$MaximumNonTitleSlideTitleCharactersChinese = 28
$MaximumTitleSlideTitleCharacters = 180
$MaximumTitleSlideTitleCharactersChinese = 90
$MaximumTakeawayCharacters = 180
$MaximumTakeawayCharactersChinese = 90
$MaximumBulletCharacters = 120
$MaximumBulletCharactersChinese = 60
$MaximumExplanationPointsPerSlide = 4
$MaximumExplanationPointCharacters = 120
$MaximumExplanationPointCharactersChinese = 60
$MaximumResultAnalysisPointsPerSlide = 4
$MaximumResultAnalysisPointCharacters = 140
$MaximumResultAnalysisPointCharactersChinese = 70

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
    # 组会正文和图表分析不得靠缩到 16pt 以下来容纳全文。前面的服务端已限制文字量；
    # 若仍放不下，让质量审计报告版式问题，而不是交付难以阅读的小字。
    $minimumFontSize = if ($FontSize -ge 35) { 35 } elseif ($FontSize -ge 24) { 24 } elseif ($FontSize -ge 16) { 18 } else { 8 }
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

function Add-AtomicFigureImage {
    param(
        $Slide,
        [string]$Name,
        [string]$ImagePath,
        [string]$AssetDirectory,
        [int]$Primary,
        [single]$FrameLeft = 510,
        [single]$FrameTop = 135,
        [single]$MaximumWidth = 405,
        [single]$MaximumHeight = 230,
        [string]$VisualRole = 'generic'
    )
    $validatedImage = Get-PaperToJournalClubApprovedRasterImage -ImagePath $ImagePath -AllowedRoots @($AssetDirectory) -ParameterName 'suggested figure image'
    # 不把 -1/-1 的“由 PowerPoint 自行探测原始尺寸”交给 COM：某些 Office 版本会在
    # 后台自动化会话中为 PNG 的尺寸探测停顿。安全校验已经读取了真实尺寸，因此先算出
    # 画布内的目标尺寸，再一次性创建图片对象，比例仍由原始像素宽高保证。
    $scale = [Math]::Min($MaximumWidth / [double]$validatedImage.width, $MaximumHeight / [double]$validatedImage.height)
    $displayWidth = [single]($validatedImage.width * $scale)
    $displayHeight = [single]($validatedImage.height * $scale)
    $displayLeft = [single]($FrameLeft + (($MaximumWidth - $displayWidth) / 2))
    $displayTop = [single]($FrameTop + (($MaximumHeight - $displayHeight) / 2))
    $image = $Slide.Shapes.AddPicture($validatedImage.path, $false, $true, $displayLeft, $displayTop, $displayWidth, $displayHeight)
    $image.Name = $Name
    $image.LockAspectRatio = -1
    $image.Line.ForeColor.RGB = $Primary
    $image.Line.Weight = 1
    # Explicitly preserve the atomic-raster contract for later structural review.
    $image.AlternativeText = "atomic_raster_unit=true; contains_reconstructable_content=false; source=tightly-cropped-paper-figure; visual_role=$VisualRole"
    return $image
}

function Test-GeneratorTextContainsChinese {
    param([string]$Text)
    return -not [string]::IsNullOrWhiteSpace($Text) -and $Text -match '[\u4e00-\u9fff]'
}

function Get-GeneratorTextLimit {
    param([string]$Text, [int]$ChineseLimit, [int]$OtherLimit)

    # 中文汇报常直接引用英文论文证据。每一条按实际文字内容选取限制，而不是按 deck.language
    # 一刀切，否则 120 字符以内的英文原句会被错误地套用中文的 60 字符限制。
    if (Test-GeneratorTextContainsChinese $Text) { return $ChineseLimit }
    return $OtherLimit
}

function Get-GeneratorStringArray {
    <#
      设计端输出的 explanation_points / result_analysis 允许是 JSON 字符串数组；为兼容
      Windows PowerShell 对单元素数组的还原，也接受单个字符串。此处统一收紧条数和长度，
      以免 COM 自动排版把科学结论缩小到不可读。
    #>
    param(
        $Slide,
        [string]$PropertyName,
        [int]$MaximumItems,
        [int]$MaximumItemCharacters,
        [bool]$Required = $false
    )

    $property = $Slide.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) {
        if ($Required) { throw "Deck spec slide $PropertyName is required." }
        return @()
    }

    $rawValue = $property.Value
    $items = if ($rawValue -is [string]) {
        @($rawValue)
    } elseif ($rawValue -is [System.Management.Automation.PSCustomObject] -or $rawValue -is [Collections.IDictionary]) {
        # result_analysis 由服务端作为单个 { comparison, interpretation, caveat } 对象写入。
        # 先包成一项再走下面的结构化展开分支，不能把 PSCustomObject 当成非法标量拒绝。
        @($rawValue)
    } elseif ($rawValue -is [Collections.IEnumerable]) {
        @($rawValue | ForEach-Object { $_ })
    } else {
        throw "Deck spec slide $PropertyName must be a string, a supported object, or a JSON array."
    }

    if ($items.Count -gt $MaximumItems) {
        throw "Deck spec slide $PropertyName may contain at most $MaximumItems items."
    }

    $normalized = @()
    foreach ($item in $items) {
        # 服务端的说明点保存 source_section_id，结果分析保存 comparison / interpretation /
        # caveat 等结构化字段。生成器只把面向听众的文字拼入 PPT，同时保留字符串数组
        # 的兼容性，避免旧 deck spec 无法再生成。
        if ($item -is [string]) {
            $text = $item.Trim()
        } elseif ($item -is [System.Management.Automation.PSCustomObject] -or $item -is [Collections.IDictionary]) {
            if ($PropertyName -eq 'explanation_points') {
                $textProperty = $item.PSObject.Properties['text']
                if ($null -eq $textProperty) { throw 'Deck spec explanation_points objects must contain text.' }
                $text = ([string]$textProperty.Value).Trim()
            } elseif ($PropertyName -eq 'result_analysis') {
                # 每个分析对象拆为独立行。比较、解释和限制各自占一行，既避免把三层
                # 科学表述挤成一句，也能让单条中文长度限制与实际版式保持一致。
                $parts = @()
                foreach ($fieldName in @('comparison', 'interpretation', 'caveat')) {
                    $field = $item.PSObject.Properties[$fieldName]
                    if ($null -ne $field -and $null -ne $field.Value) {
                        $value = ([string]$field.Value).Trim()
                        if ($value) { $parts += $value }
                    }
                }
                if ($parts.Count -eq 0) {
                    throw 'Deck spec result_analysis objects must contain comparison, interpretation, or caveat.'
                }
                foreach ($part in $parts) {
                    $partLimit = Get-GeneratorTextLimit -Text $part -ChineseLimit $MaximumResultAnalysisPointCharactersChinese -OtherLimit $MaximumItemCharacters
                    if ($part.Length -gt $partLimit) {
                        throw "Deck spec slide $PropertyName items must be no longer than $partLimit characters."
                    }
                    $normalized += $part
                }
                continue
            } else {
                throw "Deck spec slide $PropertyName items must be strings."
            }
        } else {
            throw "Deck spec slide $PropertyName items must be strings or supported objects."
        }
        if (-not $text) { throw "Deck spec slide $PropertyName items must not be empty." }
        $chineseLimit = if ($PropertyName -eq 'result_analysis') { $MaximumResultAnalysisPointCharactersChinese } elseif ($PropertyName -eq 'explanation_points') { $MaximumExplanationPointCharactersChinese } else { $MaximumBulletCharactersChinese }
        $textLimit = Get-GeneratorTextLimit -Text $text -ChineseLimit $chineseLimit -OtherLimit $MaximumItemCharacters
        if ($text.Length -gt $textLimit) {
            throw "Deck spec slide $PropertyName items must be no longer than $textLimit characters."
        }
        $normalized += $text
    }
    # result_analysis 的一个结构化对象可展开为“比较、解释、限制”三行；限制最终行数
    # 而不只是原始对象数，避免多组比较把左侧分析区挤满。
    if ($normalized.Count -gt $MaximumItems) {
        throw "Deck spec slide $PropertyName expands to more than $MaximumItems presentation points."
    }
    return @($normalized)
}

function Get-GeneratorVisualRole {
    <#
      视觉角色只影响已确认图片的排版，不参与图像来源判断。没有图片时仍回退为全宽文字，
      绝不画“待插图”区域，避免用户把缺失素材误读为论文证据。
    #>
    param($Slide)

    $property = $Slide.PSObject.Properties['visual_role']
    if ($null -eq $property -or $null -eq $property.Value) { return 'generic' }
    if ($property.Value -isnot [string]) { throw 'Deck spec slide visual_role must be a string.' }
    $role = (($property.Value).Trim().ToLowerInvariant() -replace '[ _]+', '-')
    if (-not $role) { return 'generic' }
    # methods-explanation 与 result-analysis 可由设计器在没有可靠图片时标记出来；此时采用
    # 通用全宽文本版式，而不是把它们误判为未受支持的视觉类型。
    if ($role -in @('methods-explanation', 'result-analysis')) { return 'generic' }
    if ($role -notin @('system-architecture', 'result-figure', 'result-table')) {
        throw "Deck spec slide visual_role is not supported: $role"
    }
    return $role
}

function Join-GeneratorPointText {
    param([string[]]$Points)
    if (@($Points).Count -eq 0) { return '' }
    # 使用编号而不是“标签：内容”的仪表盘式写法，便于汇报者按流程解释系统或结果。
    $number = 0
    return (@($Points | ForEach-Object {
        $number++
        "$number. $_"
    }) -join "`r")
}

function Get-GeneratorDeckAssetDirectory {
    param($Deck)
    $assetDirectory = $null
    try { $assetDirectory = [string]$Deck.evidence_pack.extraction.asset_directory } catch { $assetDirectory = $null }
    if ([string]::IsNullOrWhiteSpace($assetDirectory)) { return $null }
    return Assert-PaperToJournalClubAllowedPath -Path $assetDirectory -AllowedRoots @((Get-PaperToJournalClubTemporaryRoot)) -ParameterName 'deck_spec evidence asset directory'
}

function Assert-GeneratorDeckLimits {
    param($Deck, $Slides)
    if ($null -eq $Deck -or $null -eq $Slides) { throw 'Deck spec must contain slides.' }
    $slideItems = @($Slides | ForEach-Object { $_ })
    if ($slideItems.Count -lt 1 -or $slideItems.Count -gt $MaximumDeckSlides) { throw "Deck spec must contain between 1 and $MaximumDeckSlides slides." }
    foreach ($slide in $slideItems) {
        if ($null -eq $slide -or $slide -isnot [System.Management.Automation.PSCustomObject]) {
            throw 'Deck spec slides must be JSON objects.'
        }
        $isTitleSlide = ([string]$slide.kind).Trim().ToLowerInvariant() -eq 'title'
        $title = $slide.PSObject.Properties['title']
        if ($null -eq $title -or $title.Value -isnot [string] -or [string]::IsNullOrWhiteSpace($title.Value)) {
            throw 'Deck spec slide title must be a non-empty string.'
        }
        $titleMaximum = if ($isTitleSlide) {
            Get-GeneratorTextLimit -Text ([string]$title.Value) -ChineseLimit $MaximumTitleSlideTitleCharactersChinese -OtherLimit $MaximumTitleSlideTitleCharacters
        } else {
            Get-GeneratorTextLimit -Text ([string]$title.Value) -ChineseLimit $MaximumNonTitleSlideTitleCharactersChinese -OtherLimit $MaximumNonTitleSlideTitleCharacters
        }
        if ($title.Value.Trim().Length -gt $titleMaximum) {
            throw "Deck spec slide title exceeds the $titleMaximum-character readability limit."
        }
        # 观众看到的非封面标题应是结论性短语，不使用疑问句或“概念：解释”的目录式写法。
        if (-not $isTitleSlide -and $title.Value -match '[:：?？]') {
            throw 'Deck spec non-title slide titles must not contain a colon or question mark.'
        }

        $takeaway = $slide.PSObject.Properties['takeaway']
        if ($null -ne $takeaway) {
            $takeawayMaximum = Get-GeneratorTextLimit -Text ([string]$takeaway.Value) -ChineseLimit $MaximumTakeawayCharactersChinese -OtherLimit $MaximumTakeawayCharacters
            if ($takeaway.Value -isnot [string] -or [string]::IsNullOrWhiteSpace($takeaway.Value) -or $takeaway.Value.Trim().Length -gt $takeawayMaximum) {
                throw "Deck spec slide takeaway must be a non-empty string no longer than $takeawayMaximum characters."
            }
        } elseif (-not $isTitleSlide) {
            throw 'Deck spec non-title slides must include a takeaway.'
        }

        foreach ($propertyName in @('subtitle', 'source_text', 'suggested_image_path', 'source_asset_id', 'figure_label')) {
            $value = $slide.PSObject.Properties[$propertyName]
            if ($null -ne $value -and ($value.Value -isnot [string] -or $value.Value.Length -gt $MaximumTextCharactersPerSlide)) {
                throw "Deck spec slide $propertyName must be a string no longer than $MaximumTextCharactersPerSlide characters."
            }
        }

        [void](Get-GeneratorStringArray -Slide $slide -PropertyName 'bullets' -MaximumItems $MaximumBulletsPerSlide -MaximumItemCharacters $MaximumBulletCharacters)
        [void](Get-GeneratorStringArray -Slide $slide -PropertyName 'explanation_points' -MaximumItems $MaximumExplanationPointsPerSlide -MaximumItemCharacters $MaximumExplanationPointCharacters)
        [void](Get-GeneratorStringArray -Slide $slide -PropertyName 'result_analysis' -MaximumItems $MaximumResultAnalysisPointsPerSlide -MaximumItemCharacters $MaximumResultAnalysisPointCharacters)
        [void](Get-GeneratorVisualRole -Slide $slide)
    }
}

function Resolve-GeneratorPreviewDirectory {
    param([string]$RequestedDirectory)
    # 即使用户显式请求预览，也默认写入插件专用临时根，而不是污染 PPTX 输出目录。
    $directory = if ([string]::IsNullOrWhiteSpace($RequestedDirectory)) {
        Join-Path (Get-PaperToJournalClubTemporaryRoot) "generated-previews-$([Guid]::NewGuid().ToString('N'))"
    } else {
        $RequestedDirectory
    }
    $safeDirectory = Assert-PaperToJournalClubAllowedPath -Path $directory -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'PreviewDirectory'
    if (Test-Path -LiteralPath $safeDirectory) {
        if ($null -ne (Get-ChildItem -LiteralPath $safeDirectory -Force | Select-Object -First 1)) {
            throw "PreviewDirectory must be a new or empty directory: $safeDirectory"
        }
    }
    return $safeDirectory
}

 $pluginRoot = Split-Path -Parent $PSScriptRoot
 $deckReadRoots = @((Get-PaperToJournalClubApprovedWriteRoots) + (Join-Path $pluginRoot 'examples'))
 $deckSpecFullPath = Assert-PaperToJournalClubAllowedPath -Path $DeckSpecPath -AllowedRoots $deckReadRoots -ParameterName 'DeckSpecPath'
if ([IO.Path]::GetExtension($deckSpecFullPath).ToLowerInvariant() -ne '.json') { throw 'DeckSpecPath must end with .json.' }
if (-not (Test-Path -LiteralPath $deckSpecFullPath -PathType Leaf)) { throw "Deck spec was not found: $deckSpecFullPath" }
if ((Get-Item -LiteralPath $deckSpecFullPath -Force).Length -gt $MaximumDeckSpecBytes) { throw "Deck spec exceeds the $MaximumDeckSpecBytes-byte safety limit." }
# PowerShell 5.1 must be told the JSON encoding explicitly.
$deck = Get-Content -Raw -Encoding UTF8 -LiteralPath $deckSpecFullPath | ConvertFrom-Json
$slideSpecs = @($deck.slides | ForEach-Object { $_ })
$slideTotal = $slideSpecs.Count
Assert-GeneratorDeckLimits -Deck $deck -Slides $slideSpecs
$assetDirectory = Get-GeneratorDeckAssetDirectory $deck
$isChinese = ([string]$deck.language) -match '^(?i:zh)(-|$)'
$labels = if ($isChinese) {
    [pscustomobject]@{
        title_subtitle = '组会汇报 | 可编辑演示文稿草稿'
        source_prefix = '来源'
        no_source = '请在汇报前补充可追溯的论文来源。'
    }
} else {
    [pscustomobject]@{
        title_subtitle = 'Journal Club | Editable presentation draft'
        source_prefix = 'Source'
        no_source = 'Add a traceable paper source before presenting.'
    }
}
$outputFull = Assert-PaperToJournalClubAllowedPath -Path $OutputPath -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'OutputPath'
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
$generatorPowerPointProcessId = $null
# 只有 AddPicture 成功返回后才记录。MCP 父进程据此导出长期保留的原图，
# 因而不会把 deck spec 中最终未插入的候选图片误保存下来。
$insertedFigureAssets = @()

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    # 创建新演示文稿前同样强制禁用 Office 自动化宏，避免将来流程中意外打开外部内容时降级。
    Set-PaperToJournalClubOfficeAutomationSecurity -Application $ppt -OfficeApplication 'Microsoft PowerPoint' -DisplayAlertsValue 1
    $ppt.Visible = -1
    try {
        $windowHandle = [int64]$ppt.HWND
        $generatorProcess = @(Get-Process -Name POWERPNT -ErrorAction Stop | Where-Object { [int64]$_.MainWindowHandle -eq $windowHandle } | Select-Object -First 1)
        if ($generatorProcess.Count -eq 1) { $generatorPowerPointProcessId = [int]$generatorProcess[0].Id }
    } catch { }
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

            $imagePath = [string]$slideSpec.suggested_image_path
            $hasApprovedImage = -not [string]::IsNullOrWhiteSpace($imagePath)
            $visualRole = Get-GeneratorVisualRole -Slide $slide
            $bullets = @(Get-GeneratorStringArray -Slide $slideSpec -PropertyName 'bullets' -MaximumItems $MaximumBulletsPerSlide -MaximumItemCharacters $MaximumBulletCharacters)
            $explanationPoints = @(Get-GeneratorStringArray -Slide $slideSpec -PropertyName 'explanation_points' -MaximumItems $MaximumExplanationPointsPerSlide -MaximumItemCharacters $MaximumExplanationPointCharacters)
            $resultAnalysis = @(Get-GeneratorStringArray -Slide $slideSpec -PropertyName 'result_analysis' -MaximumItems $MaximumResultAnalysisPointsPerSlide -MaximumItemCharacters $MaximumResultAnalysisPointCharacters)
            $sourceClaimIds = @($slideSpec.PSObject.Properties['source_claim_ids'] | ForEach-Object { @($_.Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) })
            $sourceSectionIds = @($slideSpec.PSObject.Properties['source_section_ids'] | ForEach-Object { @($_.Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) })
            $sourceFigureIds = @($slideSpec.PSObject.Properties['source_figure_ids'] | ForEach-Object { @($_.Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) })
            $sourceText = if ($slideSpec.source_text) {
                [string]$slideSpec.source_text
            } elseif ($sourceClaimIds.Count) {
                "$($labels.source_prefix) $($sourceClaimIds -join ', ')"
            } elseif ($sourceSectionIds.Count) {
                "$($labels.source_prefix) $($sourceSectionIds -join ', ')"
            } elseif ($sourceFigureIds.Count) {
                "$($labels.source_prefix) $($sourceFigureIds -join ', ')"
            } else {
                $labels.no_source
            }

            # 无图时让正文占满画布。这样不会把“还没有提取到图”伪装成一个空的图像框。
            if (-not $hasApprovedImage) {
                Add-EditableText -Slide $slide -Name "slide-$index-takeaway" -Text $slideSpec.takeaway -Left 52 -Top 122 -Width 825 -Height 74 -FontSize 24 -Color $primary -Bold $true | Out-Null
                # 结果分析已经把 interpretation/caveat 结构化展开；若再叠加 bullets 会让
                # 同一句显示两遍。无图页面只选择最有针对性的一组说明。
                $fullWidthPoints = if ($resultAnalysis.Count) { @($resultAnalysis) } elseif ($explanationPoints.Count) { @($explanationPoints) } else { @($bullets) }
                $fullWidthText = if ($fullWidthPoints.Count) { Join-GeneratorPointText -Points $fullWidthPoints } else { '' }
                if ($fullWidthText) {
                    Add-EditableText -Slide $slide -Name "slide-$index-analysis" -Text $fullWidthText -Left 55 -Top 218 -Width 825 -Height 230 -FontSize 18 -Color $primary | Out-Null
                }
                Add-EditableText -Slide $slide -Name "slide-$index-source" -Text $sourceText -Left 55 -Top 462 -Width 825 -Height 24 -FontSize 12 -Color $primary | Out-Null
            } else {
                if ([string]::IsNullOrWhiteSpace($assetDirectory)) { throw 'Deck spec image paths require a plugin-extracted temporary asset directory.' }

                # 系统结构图把图放在主视觉区，左侧以顺序说明其输入、流程和输出；实验图和表格
                # 则优先保留更大的原始数据区域，把对比结论放到单独的分析区，避免图与结论混杂。
                $imageFrame = switch ($visualRole) {
                    'system-architecture' { [pscustomobject]@{ left = 466; top = 122; width = 440; height = 286 } }
                    'result-table'        { [pscustomobject]@{ left = 430; top = 118; width = 476; height = 300 } }
                    'result-figure'       { [pscustomobject]@{ left = 430; top = 118; width = 476; height = 300 } }
                    default                { [pscustomobject]@{ left = 510; top = 135; width = 405; height = 230 } }
                }
                Add-AtomicFigureImage -Slide $slide -Name "slide-$index-figure-image" -ImagePath $imagePath -AssetDirectory $assetDirectory -Primary $primary `
                    -FrameLeft $imageFrame.left -FrameTop $imageFrame.top -MaximumWidth $imageFrame.width -MaximumHeight $imageFrame.height -VisualRole $visualRole | Out-Null

                $reportedFigureIds = @()
                foreach ($figureId in @($slideSpec.source_figure_ids) + @($slideSpec.suggested_figure_id)) {
                    $normalizedFigureId = ([string]$figureId).Trim()
                    if ($normalizedFigureId -and $normalizedFigureId -notin $reportedFigureIds) { $reportedFigureIds += $normalizedFigureId }
                }
                $insertedFigureAssets += [pscustomobject]@{
                    source_path = $imagePath
                    source_asset_id = ([string]$slideSpec.source_asset_id).Trim()
                    slide_id = ([string]$slideSpec.id).Trim()
                    slide_number = $index
                    figure_id = ([string]$slideSpec.suggested_figure_id).Trim()
                    figure_ids = @($reportedFigureIds)
                }

                switch ($visualRole) {
                    'system-architecture' {
                        Add-EditableText -Slide $slide -Name "slide-$index-takeaway" -Text $slideSpec.takeaway -Left 52 -Top 122 -Width 370 -Height 112 -FontSize 24 -Color $primary -Bold $true | Out-Null
                        $systemPoints = if ($explanationPoints.Count) { $explanationPoints } else { $bullets }
                        $systemText = Join-GeneratorPointText -Points $systemPoints
                        if ($systemText) {
                            Add-EditableText -Slide $slide -Name "slide-$index-system-explanation" -Text $systemText -Left 55 -Top 254 -Width 365 -Height 188 -FontSize 18 -Color $primary | Out-Null
                        }
                        Add-EditableText -Slide $slide -Name "slide-$index-source" -Text $sourceText -Left 468 -Top 423 -Width 438 -Height 34 -FontSize 12 -Color $primary | Out-Null
                    }
                    'result-table' {
                        Add-EditableText -Slide $slide -Name "slide-$index-takeaway" -Text $slideSpec.takeaway -Left 52 -Top 122 -Width 345 -Height 96 -FontSize 23 -Color $primary -Bold $true | Out-Null
                        $tablePoints = if ($resultAnalysis.Count) { $resultAnalysis } elseif ($explanationPoints.Count) { $explanationPoints } else { $bullets }
                        $tableText = Join-GeneratorPointText -Points $tablePoints
                        if ($tableText) {
                            Add-EditableText -Slide $slide -Name "slide-$index-result-analysis" -Text $tableText -Left 55 -Top 239 -Width 345 -Height 202 -FontSize 18 -Color $primary | Out-Null
                        }
                        Add-EditableText -Slide $slide -Name "slide-$index-source" -Text $sourceText -Left 430 -Top 430 -Width 476 -Height 30 -FontSize 12 -Color $primary | Out-Null
                    }
                    'result-figure' {
                        Add-EditableText -Slide $slide -Name "slide-$index-takeaway" -Text $slideSpec.takeaway -Left 52 -Top 122 -Width 345 -Height 96 -FontSize 23 -Color $primary -Bold $true | Out-Null
                        $resultPoints = if ($resultAnalysis.Count) { $resultAnalysis } elseif ($explanationPoints.Count) { $explanationPoints } else { $bullets }
                        $resultText = Join-GeneratorPointText -Points $resultPoints
                        if ($resultText) {
                            Add-EditableText -Slide $slide -Name "slide-$index-result-analysis" -Text $resultText -Left 55 -Top 239 -Width 345 -Height 202 -FontSize 18 -Color $primary | Out-Null
                        }
                        Add-EditableText -Slide $slide -Name "slide-$index-source" -Text $sourceText -Left 430 -Top 430 -Width 476 -Height 30 -FontSize 12 -Color $primary | Out-Null
                    }
                    default {
                        Add-EditableText -Slide $slide -Name "slide-$index-takeaway" -Text $slideSpec.takeaway -Left 52 -Top 123 -Width 405 -Height 142 -FontSize 24 -Color $primary -Bold $true | Out-Null
                        $genericPoints = if ($resultAnalysis.Count) { @($resultAnalysis) } elseif ($explanationPoints.Count) { @($explanationPoints) } else { @($bullets) }
                        $genericText = Join-GeneratorPointText -Points $genericPoints
                        if ($genericText) {
                            Add-EditableText -Slide $slide -Name "slide-$index-analysis" -Text $genericText -Left 55 -Top 292 -Width 395 -Height 165 -FontSize 18 -Color $primary | Out-Null
                        }
                        Add-EditableText -Slide $slide -Name "slide-$index-source" -Text $sourceText -Left 512 -Top 390 -Width 390 -Height 44 -FontSize 12 -Color $primary | Out-Null
                    }
                }
            }
        }
        Add-EditableText -Slide $slide -Name "slide-$index-footer" -Text "$($deck.paper.title) | $index / $slideTotal" -Left 48 -Top 503 -Width 860 -Height 18 -FontSize 9 -Color $primary | Out-Null
    }
    # ppSaveAsOpenXMLPresentation = 24; output remains editable in PowerPoint.
    $presentation.SaveAs($outputFull, 24)
    if (-not (Test-Path -LiteralPath $outputFull)) { throw "PowerPoint did not write the expected PPTX: $outputFull" }
    # 显式 PreviewDirectory 兼容旧的直调方式，并视作用户要求导出预览；没有任一显式
    # 请求时只返回结构质量审核，不在成品目录产生额外 PNG 文件。
    $shouldExportPreviews = ($ExportPreviews -or -not [string]::IsNullOrWhiteSpace($PreviewDirectory)) -and -not $SkipPreviewExport
    $resolvedPreviewDirectory = if ($shouldExportPreviews) {
        Resolve-GeneratorPreviewDirectory -RequestedDirectory $PreviewDirectory
    } else {
        $null
    }
    $qualityAuditDeferred = [bool]$DeferQualityAudit
    if ($DeferQualityAudit) {
        # MCP 宿主会在本子进程完全退出后执行只读审核，避免同一 COM 进程的刷新竞态。
        $quality = $null
    } else {
    if (-not $KeepOpen) {
        # 先关闭生成会话，确保最终审核针对磁盘中的交付物，而不是内存中的临时状态。
        try { $presentation.Close() } catch { }
        try { $ppt.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch { }
        $presentation = $null
        $ppt = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        # 只等待本脚本新建的 PowerPoint 进程退出，不干扰用户已打开的演示文稿。若 Office
        # 尚在后台卸载 COM server，过早重新打开会读到不完整的 Slides/Shapes collection。
        if ($generatorPowerPointProcessId) {
            try {
                $generatorProcess = Get-Process -Id $generatorPowerPointProcessId -ErrorAction Stop
                [void]$generatorProcess.WaitForExit(5000)
            } catch { }
        }
        Start-Sleep -Milliseconds 250
    }

    # 在全新的 PowerShell/COM 进程中重新打开 PPTX 审核。PowerPoint 在同一进程里先创建、
    # 保存、关闭再立刻读取时，部分版本会返回尚未刷新的 Shapes collection；独立审计进程
    # 同时验证成品文件可重开，避免把 COM 瞬态误报给用户。
    $auditScriptPath = Join-Path $PSScriptRoot 'audit-editable-pptx.ps1'
    $auditArguments = @('-NoProfile', '-STA', '-ExecutionPolicy', 'RemoteSigned', '-File', $auditScriptPath, '-PresentationPath', $outputFull)
    if ($shouldExportPreviews) {
        $auditArguments += '-ExportPreviews'
        if ($resolvedPreviewDirectory) { $auditArguments += @('-PreviewDirectory', $resolvedPreviewDirectory) }
    }
    $quality = $null
    for ($auditAttempt = 1; $auditAttempt -le 4; $auditAttempt++) {
        $auditOutput = @(& powershell.exe @auditArguments 2>&1)
        $auditExitCode = $LASTEXITCODE
        $auditJsonLine = @($auditOutput | ForEach-Object { [string]$_ } | Where-Object { $_.TrimStart().StartsWith('{') } | Select-Object -Last 1)
        if ($auditJsonLine.Count -ne 1) {
            throw "PowerPoint audit worker completed without its required JSON report: $($auditOutput -join [Environment]::NewLine)"
        }
        try {
            $quality = $auditJsonLine[0] | ConvertFrom-Json
        } catch {
            throw "PowerPoint audit worker returned invalid JSON. $($_.Exception.Message)"
        }
        if ($auditExitCode -notin @(0, 2)) {
            throw "PowerPoint audit worker failed with exit code ${auditExitCode}: $($auditOutput -join [Environment]::NewLine)"
        }

        # 仅对明确的 Office 启动同步态重试：空 slide collection，或所有错误都没有具体
        # slide/shape 名称。真实版式、文字溢出或图片错误不会被重试掩盖。
        $anonymousCollectionFailures = @($quality.findings | Where-Object {
            $_.severity -eq 'hard' -and $_.category -in @('slides', 'geometry', 'editability') -and $null -eq $_.slide -and $null -eq $_.shape
        })
        $auditedSlideCount = 0
        try { $auditedSlideCount = [int]$quality.presentation.slide_count } catch { }
        $retryableAuditState = ($auditedSlideCount -lt 1) -or ($anonymousCollectionFailures.Count -gt 0 -and $anonymousCollectionFailures.Count -eq @($quality.findings).Count)
        if (-not $retryableAuditState -or $auditAttempt -eq 4) { break }
        Start-Sleep -Milliseconds 750
    }
    }
    if (-not $DeferQualityAudit -and -not $quality.pass -and -not $DoNotFailOnAudit) {
        throw "PowerPoint quality audit failed: $($quality.findings | ConvertTo-Json -Compress)"
    }
    Write-Output (ConvertTo-Json @{ output_path = $outputFull; slide_count = $slideTotal; editable_contract = "native-text-shapes"; quality_audit = $quality; quality_audit_deferred = $qualityAuditDeferred; preview_directory = if ($shouldExportPreviews) { $resolvedPreviewDirectory } else { $null }; inserted_figure_assets = @($insertedFigureAssets) } -Depth 40 -Compress)
} finally {
    # SaveAs 成功后，部分 Office 版本会在后台窗口关闭时先断开 RPC。清理失败不能覆盖
    # 已写入并已审核的 PPTX 结果；仍逐项尝试关闭和释放其它 COM 对象，避免遗留进程。
    if ($presentation -and -not $KeepOpen) { try { $presentation.Close() } catch { } }
    if ($ppt -and -not $KeepOpen) { try { $ppt.Quit() } catch { } }
    if ($presentation) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch { } }
    if ($ppt) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch { } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
