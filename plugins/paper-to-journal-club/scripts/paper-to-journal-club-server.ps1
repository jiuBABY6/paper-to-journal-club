<#
  无 Node 的 MCP 服务。

  这是插件的唯一运行时入口：标准输入/输出处理 MCP JSON-RPC，
  PDF 文本由发布包中的自包含 parser.exe 或 Microsoft Word COM 读取，
  PPT 由同目录的 PowerPoint COM 脚本生成。
#>
param(
    [switch]$Demo,
    [string]$DemoInputPath,
    [string]$DemoOutputPath
)

$ErrorActionPreference = "Stop"
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$ScriptRoot = $PSScriptRoot
$PluginRoot = Split-Path -Parent $ScriptRoot
$PluginManifestPath = Join-Path $PluginRoot '.codex-plugin\plugin.json'
$ParserPath = Join-Path $PluginRoot "assets\paper-parser.exe"
$PptGeneratorPath = Join-Path $ScriptRoot "generate-editable-pptx.ps1"
$PowerPointQualityPath = Join-Path $ScriptRoot "powerpoint-quality.ps1"
$PdfVisualRendererPath = Join-Path $ScriptRoot "render-pdf-page.ps1"
$SupportedProtocolVersions = @('2024-11-05', '2025-03-26', '2025-06-18')
$MaximumPaperBytes = 50MB
$ParserTimeoutMilliseconds = 120000
# MCP stdio 没有浏览器或 HTTP 网关代为限流，必须在服务端自行限制请求、对象图和生成规模。
$MaximumMcpRequestCharacters = 1MB
$MaximumToolArgumentCharacters = 768KB
$MaximumToolObjectDepth = 20
$MaximumToolObjectProperties = 96
$MaximumToolArrayItems = 256
$MaximumToolObjectNodes = 2500
$MaximumDeckSlides = 30
$MaximumBulletsPerSlide = 3
$MaximumTextCharactersPerSlide = 8000
$MaximumNonTitleSlideTitleCharacters = 64
$MaximumNonTitleSlideTitleCharactersChinese = 28
$MaximumTakeawayCharacters = 180
$MaximumTakeawayCharactersChinese = 90
$MaximumBulletCharacters = 120
$MaximumBulletCharactersChinese = 60
$MaximumExplanationPointsPerSlide = 4
$MaximumParserOutputCharacters = 4MB
$MaximumParserErrorCharacters = 64KB
$PdfVisualRenderTimeoutMilliseconds = 30000
$MaximumPdfVisualRenderOutputCharacters = 16KB
$MaximumPdfVisualRenderErrorCharacters = 32KB
$MaximumRenderedVisualsPerPaper = 12
$MaximumRenderedVisualBytesPerPaper = 100MB
$MaximumRenderedVisualPixels = 8000000
$MaximumExtractedTextCharacters = 750000
# cleanup_paper_assets 只能删除由 analyse_paper 成功创建并标记的临时资产目录。这样即使
# 用户把 PPTX 输出到插件临时根，长期保留的 <PPT名>_assets 也不会被误删。
$TemporaryPaperAssetMarkerName = '.paper-to-journal-club-temporary-assets'
$TemporaryPaperAssetMarkerContent = 'paper-to-journal-club-temporary-assets-v1'

# MCP 握手返回的版本必须与插件清单一致，避免修复发布时遗漏同步硬编码版本号。
try {
    $PluginManifest = Get-Content -LiteralPath $PluginManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $PluginVersion = [string]$PluginManifest.version
    if ($PluginVersion -notmatch '^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$') {
        throw 'version is missing or is not a semantic version.'
    }
} catch {
    throw "Could not load the plugin version from ${PluginManifestPath}: $($_.Exception.Message)"
}

# 与生成器共用 PowerPoint COM 检查和原生渲染审计，避免 MCP 层声称了不存在的能力。
. $PowerPointQualityPath

# 组会汇报的默认必备模块。使用稳定的英文 ID，避免 Windows PowerShell 5.1
# 在不同系统代码页下处理中文 JSON 键名时出现不兼容。
$DefaultRequiredSections = @(
    'background',
    'innovation',
    'methods',
    'experimental_data',
    'limitations',
    'future_directions'
)

$KnownJournalClubSections = @(
    'background',
    'innovation',
    'methods',
    'experimental_data',
    'limitations',
    'future_directions'
)

function Get-PropertyValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $Default
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return $property.Value
}

function Test-PropertyExists {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $false }
    if ($Object -is [Collections.IDictionary]) { return $Object.Contains($Name) }
    return $null -ne $Object.PSObject.Properties[$Name]
}

function Get-StrictBoolean {
    param($Object, [string]$Name, [bool]$Default = $false)

    if (-not (Test-PropertyExists $Object $Name)) { return $Default }
    $value = Get-PropertyValue $Object $Name
    # PowerShell 的 [bool]'false' 会得到 $true；因此绝不做字符串或数字转换。
    if ($null -eq $value -or $value -isnot [bool]) {
        throw "$Name must be a JSON boolean (true or false), not a string, number, or null."
    }
    return $value
}

function Assert-McpObject {
    param($Value, [string]$ParameterName)
    if ($null -eq $Value -or ($Value -isnot [Collections.IDictionary] -and $Value -isnot [System.Management.Automation.PSCustomObject])) {
        throw "$ParameterName must be a JSON object."
    }
}

function Get-McpObjectPropertyNames {
    param($Value)
    Assert-McpObject -Value $Value -ParameterName 'arguments'
    if ($Value -is [Collections.IDictionary]) { return @($Value.Keys | ForEach-Object { [string]$_ }) }
    return @($Value.PSObject.Properties | ForEach-Object { $_.Name })
}

function Assert-OnlyKnownArguments {
    param($Arguments, [string[]]$AllowedNames, [string]$ToolName)
    foreach ($name in @(Get-McpObjectPropertyNames $Arguments)) {
        if ($name -notin $AllowedNames) { throw "$ToolName received an unsupported argument: $name" }
    }
}

function Assert-StringArgument {
    param($Arguments, [string]$Name, [bool]$Required = $false, [int]$MaximumLength = 4096)
    if (-not (Test-PropertyExists $Arguments $Name)) {
        if ($Required) { throw "$Name is required." }
        return
    }
    $value = Get-PropertyValue $Arguments $Name
    if ($value -isnot [string] -or [string]::IsNullOrWhiteSpace($value)) { throw "$Name must be a non-empty string." }
    if ($value.Length -gt $MaximumLength) { throw "$Name exceeds the $MaximumLength-character safety limit." }
}

function Assert-IntegerArgument {
    param($Arguments, [string]$Name, [int]$Minimum, [int]$Maximum, [int]$Default)
    if (-not (Test-PropertyExists $Arguments $Name)) { return $Default }
    $value = Get-PropertyValue $Arguments $Name
    if ($value -isnot [byte] -and $value -isnot [sbyte] -and $value -isnot [int16] -and $value -isnot [uint16] -and $value -isnot [int32] -and $value -isnot [uint32] -and $value -isnot [int64]) {
        throw "$Name must be an integer."
    }
    if ([int64]$value -lt $Minimum -or [int64]$value -gt $Maximum) { throw "$Name must be between $Minimum and $Maximum." }
    return [int]$value
}

function Assert-StringArrayArgument {
    param($Arguments, [string]$Name, [int]$MinimumItems = 0, [int]$MaximumItems = 32, [int]$MaximumItemLength = 128)
    if (-not (Test-PropertyExists $Arguments $Name)) { return }
    $value = Get-PropertyValue $Arguments $Name
    if ($value -is [string] -or $value -isnot [Collections.IEnumerable]) { throw "$Name must be a JSON array." }
    $items = @($value | ForEach-Object { $_ })
    if ($items.Count -lt $MinimumItems -or $items.Count -gt $MaximumItems) { throw "$Name must contain between $MinimumItems and $MaximumItems items." }
    foreach ($item in $items) {
        if ($item -isnot [string] -or [string]::IsNullOrWhiteSpace($item) -or $item.Length -gt $MaximumItemLength) {
            throw "$Name items must be non-empty strings no longer than $MaximumItemLength characters."
        }
    }
}

function Assert-McpValueLimits {
    param($Value, [string]$Path = '$', [int]$Depth = 0, [ref]$State)

    if ($Depth -gt $MaximumToolObjectDepth) { throw "$Path exceeds the $MaximumToolObjectDepth-level JSON nesting limit." }
    $State.Value.nodes++
    if ($State.Value.nodes -gt $MaximumToolObjectNodes) { throw "Tool arguments exceed the $MaximumToolObjectNodes-node safety limit." }
    if ($null -eq $Value -or $Value -is [bool] -or $Value -is [byte] -or $Value -is [sbyte] -or $Value -is [int16] -or $Value -is [uint16] -or $Value -is [int32] -or $Value -is [uint32] -or $Value -is [int64] -or $Value -is [double] -or $Value -is [decimal]) { return }
    if ($Value -is [string]) {
        $State.Value.characters += $Value.Length
        if ($State.Value.characters -gt $MaximumToolArgumentCharacters) { throw "Tool arguments exceed the $MaximumToolArgumentCharacters-character safety limit." }
        return
    }
    if ($Value -is [Collections.IDictionary] -or $Value -is [System.Management.Automation.PSCustomObject]) {
        $names = @(Get-McpObjectPropertyNames $Value)
        if ($names.Count -gt $MaximumToolObjectProperties) { throw "$Path has more than $MaximumToolObjectProperties properties." }
        foreach ($name in $names) {
            if ($name.Length -gt 128) { throw "$Path contains an overlong property name." }
            Assert-McpValueLimits -Value (Get-PropertyValue $Value $name) -Path "$Path.$name" -Depth ($Depth + 1) -State $State
        }
        return
    }
    if ($Value -is [Collections.IEnumerable]) {
        $items = @($Value | ForEach-Object { $_ })
        if ($items.Count -gt $MaximumToolArrayItems) { throw "$Path has more than $MaximumToolArrayItems array items." }
        for ($index = 0; $index -lt $items.Count; $index++) {
            Assert-McpValueLimits -Value $items[$index] -Path "$Path[$index]" -Depth ($Depth + 1) -State $State
        }
        return
    }
    throw "$Path contains an unsupported JSON value type."
}

function Test-TextContainsChinese {
    param([string]$Text)
    return -not [string]::IsNullOrWhiteSpace($Text) -and $Text -match '[\u4e00-\u9fff]'
}

function Get-ReadableTextLimit {
    param([string]$Text, [int]$ChineseLimit, [int]$OtherLimit)
    if (Test-TextContainsChinese $Text) { return $ChineseLimit }
    return $OtherLimit
}

function Assert-PresentationTextLength {
    param(
        [string]$Text,
        [string]$PropertyName,
        [int]$ChineseLimit,
        [int]$OtherLimit
    )

    if ($Text -isnot [string]) { throw "$PropertyName must be a string." }
    $limit = Get-ReadableTextLimit -Text $Text -ChineseLimit $ChineseLimit -OtherLimit $OtherLimit
    if ($Text.Length -gt $limit) {
        throw "$PropertyName is too long for a readable slide. Limit: $limit characters."
    }
}

function Assert-DeckSpecificationLimits {
    param($Deck)
    Assert-McpObject -Value $Deck -ParameterName 'deck_spec'
    $slides = Get-PropertyValue $Deck 'slides'
    if ($slides -is [string] -or $slides -isnot [Collections.IEnumerable]) { throw 'deck_spec.slides must be a JSON array.' }
    $slideItems = @($slides | ForEach-Object { $_ })
    if ($slideItems.Count -lt 1 -or $slideItems.Count -gt $MaximumDeckSlides) { throw "deck_spec.slides must contain between 1 and $MaximumDeckSlides slides." }
    foreach ($slide in $slideItems) {
        Assert-McpObject -Value $slide -ParameterName 'deck_spec.slides item'
        $slideKind = [string](Get-PropertyValue $slide 'kind' '')
        foreach ($propertyName in @('title', 'takeaway', 'subtitle', 'source_text', 'suggested_image_path', 'source_asset_id', 'figure_label', 'source_visual_id', 'visual_role_source')) {
            if (Test-PropertyExists $slide $propertyName) {
                $text = Get-PropertyValue $slide $propertyName
                if ($text -isnot [string] -or $text.Length -gt $MaximumTextCharactersPerSlide) {
                    throw "deck_spec slide $propertyName must be a string no longer than $MaximumTextCharactersPerSlide characters."
                }
            }
        }
        $title = [string](Get-PropertyValue $slide 'title' '')
        if ($slideKind -ne 'title') {
            Assert-PresentationTextLength -Text $title -PropertyName 'deck_spec slide title' -ChineseLimit $MaximumNonTitleSlideTitleCharactersChinese -OtherLimit $MaximumNonTitleSlideTitleCharacters
            if ($title -match '[:：?？]') {
                throw 'Non-title slide headings must be declarative and must not contain a colon or question mark.'
            }
        }
        if (Test-PropertyExists $slide 'takeaway') {
            Assert-PresentationTextLength -Text ([string](Get-PropertyValue $slide 'takeaway' '')) -PropertyName 'deck_spec slide takeaway' -ChineseLimit $MaximumTakeawayCharactersChinese -OtherLimit $MaximumTakeawayCharacters
        }
        if (Test-PropertyExists $slide 'bullets') {
            $bullets = Get-PropertyValue $slide 'bullets'
            # Windows PowerShell 的 ConvertFrom-Json 会把单元素数组还原成标量；
            # 兼容本插件自己产生的 deck-spec，同时仍限制条数和文本总量。
            if ($null -eq $bullets) { $bulletItems = @() }
            elseif ($bullets -is [string]) { $bulletItems = @($bullets) }
            elseif ($bullets -is [Collections.IEnumerable]) { $bulletItems = @($bullets | ForEach-Object { $_ }) }
            else { throw 'deck_spec slide bullets must be a JSON array or a single string.' }
            if ($bulletItems.Count -gt $MaximumBulletsPerSlide) { throw "deck_spec slide bullets may contain at most $MaximumBulletsPerSlide items." }
            foreach ($bullet in $bulletItems) {
                Assert-PresentationTextLength -Text $bullet -PropertyName 'deck_spec slide bullet' -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters
            }
        }
        if (Test-PropertyExists $slide 'explanation_points') {
            $points = Get-PropertyValue $slide 'explanation_points'
            if ($points -is [string] -or $points -isnot [Collections.IEnumerable]) { throw 'deck_spec slide explanation_points must be a JSON array.' }
            $pointItems = @($points | ForEach-Object { $_ })
            if ($pointItems.Count -lt 2 -or $pointItems.Count -gt $MaximumExplanationPointsPerSlide) {
                throw "deck_spec slide explanation_points must contain between 2 and $MaximumExplanationPointsPerSlide items."
            }
            foreach ($point in $pointItems) {
                $pointText = if ($point -is [string]) { $point } else { [string](Get-PropertyValue $point 'text' '') }
                Assert-PresentationTextLength -Text $pointText -PropertyName 'deck_spec slide explanation point' -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters
            }
        }
        if (Test-PropertyExists $slide 'result_analysis') {
            $analysis = Get-PropertyValue $slide 'result_analysis'
            Assert-McpObject -Value $analysis -ParameterName 'deck_spec slide result_analysis'
            foreach ($name in @('comparison', 'interpretation', 'caveat')) {
                if (Test-PropertyExists $analysis $name) {
                    Assert-PresentationTextLength -Text ([string](Get-PropertyValue $analysis $name '')) -PropertyName "deck_spec slide result_analysis.$name" -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters
                }
            }
        }
        if (Test-PropertyExists $slide 'visual_source_page') {
            $visualSourcePage = Get-PropertyValue $slide 'visual_source_page'
            if ($visualSourcePage -isnot [int] -and $visualSourcePage -isnot [int64]) {
                throw 'deck_spec slide visual_source_page must be an integer.'
            }
            if ([int]$visualSourcePage -lt 1) { throw 'deck_spec slide visual_source_page must be at least 1.' }
        }
        if (Test-PropertyExists $slide 'visual_evidence') {
            $visualEvidence = Get-PropertyValue $slide 'visual_evidence'
            Assert-McpObject -Value $visualEvidence -ParameterName 'deck_spec slide visual_evidence'
            $excerpt = Get-PropertyValue $visualEvidence 'excerpt' ''
            if ($excerpt -isnot [string] -or [string]::IsNullOrWhiteSpace($excerpt)) {
                throw 'deck_spec slide visual_evidence.excerpt must be a non-empty string.'
            }
        }
    }
}

function Get-ValidatedPaperPath {
    param([string]$FilePath)

    if ([string]::IsNullOrWhiteSpace($FilePath)) { throw 'file_path is required.' }
    # 论文只能来自用户资料目录或显式配置的资料根目录；示例目录仅供插件自检使用。
    $readRoots = @((Get-PaperToJournalClubUserDataRoots) + (Get-PaperToJournalClubTemporaryRoot) + (Join-Path $PluginRoot 'examples'))
    $absolutePath = Assert-PaperToJournalClubAllowedPath -Path $FilePath -AllowedRoots $readRoots -ParameterName 'file_path'
    if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) { throw "Paper file was not found: $absolutePath" }
    $extension = [IO.Path]::GetExtension($absolutePath).ToLowerInvariant()
    if ($extension -notin @('.pdf', '.txt', '.md', '.tex')) {
        throw 'Supported input files are PDF, TXT, Markdown, and TeX.'
    }
    $length = (Get-Item -LiteralPath $absolutePath).Length
    if ($length -gt $MaximumPaperBytes) {
        throw "Paper is larger than the $([int]($MaximumPaperBytes / 1MB)) MB safety limit."
    }
    return $absolutePath
}

function Resolve-RequestedOutputPath {
    param(
        [string]$Path,
        [string]$RequiredExtension,
        [bool]$Overwrite = $false
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'output_path is required.' }
    $writeRoots = @(Get-PaperToJournalClubApprovedWriteRoots)
    # run-demo 是开发演示入口，允许它写入插件内的示例目录；正式 MCP 调用绝不包含该目录。
    if ($Demo) { $writeRoots += (Join-Path $PluginRoot 'examples') }
    $absolutePath = Assert-PaperToJournalClubAllowedPath -Path $Path -AllowedRoots $writeRoots -ParameterName 'output_path'
    if ([IO.Path]::GetExtension($absolutePath).ToLowerInvariant() -ne $RequiredExtension.ToLowerInvariant()) {
        throw "output_path must end with $RequiredExtension."
    }
    if ((Test-Path -LiteralPath $absolutePath -PathType Leaf) -and -not $Overwrite) {
        throw "Output already exists. Set overwrite=true only when the selected file may be replaced: $absolutePath"
    }
    $parent = Split-Path -Parent $absolutePath
    if ([string]::IsNullOrWhiteSpace($parent)) { throw 'output_path must include a writable directory.' }
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
    return $absolutePath
}

function Get-ExistingPowerPointPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'file_path is required.' }
    $absolutePath = Assert-PaperToJournalClubAllowedPath -Path $Path -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'file_path'
    if ([IO.Path]::GetExtension($absolutePath).ToLowerInvariant() -ne '.pptx') { throw 'file_path must point to a .pptx file.' }
    if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) { throw "Presentation was not found: $absolutePath" }
    return $absolutePath
}

function Test-DirectoryIsEmptyOrMissing {
    param([string]$Directory)
    if (-not (Test-Path -LiteralPath $Directory)) { return $true }
    return $null -eq (Get-ChildItem -LiteralPath $Directory -Force | Select-Object -First 1)
}

function ConvertTo-NonEmptyStringArray {
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value | ForEach-Object {
        if ($null -ne $_) {
            $text = ([string]$_).Trim()
            if ($text) { $text }
        }
    })
}

function Resolve-RequiredSections {
    param($RequestedSections)

    # 用户可按汇报场景裁剪模块；未提供配置时始终采用六个组会必备模块。
    if ($null -eq $RequestedSections) { return @($DefaultRequiredSections) }
    $requested = @(ConvertTo-NonEmptyStringArray $RequestedSections)
    if ($requested.Count -eq 0) {
        throw 'required_sections cannot be empty. Omit it to use the journal-club default sections.'
    }

    $resolved = @()
    foreach ($item in $requested) {
        $sectionId = ($item.Trim().ToLowerInvariant() -replace '[\s-]+', '_')
        if ($sectionId -notin $KnownJournalClubSections) {
            throw "Unknown required_sections value: $item. Allowed values: $($KnownJournalClubSections -join ', ')."
        }
        if ($sectionId -notin $resolved) { $resolved += $sectionId }
    }
    return @($resolved)
}

function Get-ObjectIds {
    param($Items)
    $ids = @{}
    foreach ($item in @($Items)) {
        $id = Get-PropertyValue $item 'id'
        if ($id) { $ids[[string]$id] = $true }
    }
    return $ids
}

function Normalize-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    # 将连续空行和行尾空白归一化；分层括号避免 PowerShell 将连续 -replace 误解析。
    $withoutCarriageReturns = $Text -replace "`r", ""
    $withoutTrailingWhitespace = $withoutCarriageReturns -replace "[ `t]+`n", "`n"
    return ($withoutTrailingWhitespace -replace "`n{3,}", "`n`n").Trim()
}

function Split-PaperSentences {
    param([string]$Text)
    # Protect abbreviations so "Fig. 2" remains part of the evidence sentence.
    $protected = (Normalize-Text $Text) -replace '\b(Fig|fig|Dr|dr|e\.g|i\.e)\.', '$1<dot>'
    return @($protected -split '(?<=[.!?。！？])\s+|\n+' |
        ForEach-Object { ($_ -replace '<dot>', '.').Trim() } |
        Where-Object { $_.Length -ge 24 -and $_.Length -le 420 })
}

function Get-WordPdfText {
    param([string]$FilePath)
    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        # msoAutomationSecurityForceDisable = 3；Word 仅作为解析器缺失时的受限回退，
        # 打开不可信 PDF 前仍必须关闭宏和链接更新。
        Set-PaperToJournalClubOfficeAutomationSecurity -Application $word -OfficeApplication 'Microsoft Word' -DisplayAlertsValue 0
        try { $word.Options.UpdateLinksAtOpen = $false } catch { }
        $word.Visible = $false
        $document = $word.Documents.Open($FilePath, $false, $true, $false)
        return Normalize-Text $document.Content.Text
    } finally {
        if ($document) { $document.Close(0); [Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null }
        if ($word) { $word.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    }
}

function New-PaperAssetDirectory {
    param([string]$PaperPath, [string]$RequestedDirectory)
    $temporaryRoot = Get-PaperToJournalClubTemporaryRoot
    if ($RequestedDirectory) {
        # 提取资产默认是短生命周期临时数据；不允许 MCP 直接向任意研究目录批量落盘图片。
        $directory = Assert-PaperToJournalClubAllowedPath -Path $RequestedDirectory -AllowedRoots @($temporaryRoot) -ParameterName 'asset_output_dir'
        if (-not (Test-DirectoryIsEmptyOrMissing $directory)) {
            throw "asset_output_dir must be a new or empty directory so existing assets are not overwritten: $directory"
        }
    } else {
        # 默认使用用户临时目录，避免向已安装的插件目录写入运行时数据。
        $safeStem = ([IO.Path]::GetFileNameWithoutExtension($PaperPath) -replace '[^a-zA-Z0-9._-]', '-')
        if (-not $safeStem) { $safeStem = 'paper' }
        $directory = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $temporaryRoot "$safeStem-$([Guid]::NewGuid().ToString('N'))") -AllowedRoots @($temporaryRoot) -ParameterName 'default asset_output_dir'
    }
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
    return $directory
}

function Write-TemporaryPaperAssetMarker {
    param([string]$Directory)

    # parser.exe 要求目标目录初始为空，因此标记必须在解析成功、资产限制检查完成后写入。
    $temporaryRoot = Get-PaperToJournalClubTemporaryRoot
    $safeDirectory = Assert-PaperToJournalClubAllowedPath -Path $Directory -AllowedRoots @($temporaryRoot) -ParameterName 'temporary paper asset directory'
    $markerPath = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $safeDirectory $TemporaryPaperAssetMarkerName) -AllowedRoots @($temporaryRoot) -ParameterName 'temporary paper asset ownership marker'
    if (Test-Path -LiteralPath $markerPath) { throw "Temporary paper asset ownership marker already exists: $markerPath" }

    $stream = $null
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($TemporaryPaperAssetMarkerContent)
        $stream = [IO.File]::Open($markerPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}

function Test-TemporaryPaperAssetMarker {
    param([string]$Directory)

    $temporaryRoot = Get-PaperToJournalClubTemporaryRoot
    $markerPath = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $Directory $TemporaryPaperAssetMarkerName) -AllowedRoots @($temporaryRoot) -ParameterName 'temporary paper asset ownership marker'
    if (-not (Test-Path -LiteralPath $markerPath -PathType Leaf)) { return $false }
    Assert-PaperToJournalClubNoReparsePoint -Path $markerPath -ParameterName 'temporary paper asset ownership marker'
    try {
        return ([IO.File]::ReadAllText($markerPath, [Text.Encoding]::UTF8) -eq $TemporaryPaperAssetMarkerContent)
    } catch {
        return $false
    }
}

function Remove-TemporaryPaperAssets {
    param($Arguments)
    $directory = Get-PropertyValue $Arguments 'asset_output_dir'
    $confirmed = Get-StrictBoolean -Object $Arguments -Name 'confirm' -Default $false
    if (-not $confirmed) { throw 'Set confirm=true before deleting extracted temporary paper assets.' }
    if ([string]::IsNullOrWhiteSpace($directory)) { throw 'asset_output_dir is required.' }

    $temporaryRoot = Get-PaperToJournalClubTemporaryRoot
    $target = Assert-PaperToJournalClubAllowedPath -Path $directory -AllowedRoots @($temporaryRoot) -ParameterName 'asset_output_dir'
    if (-not (Test-Path -LiteralPath $target)) {
        return [pscustomobject]@{ deleted = $false; asset_output_dir = $target; note = 'Directory was already absent.' }
    }
    # 仅凭“位于临时根”不足以说明目录可删：用户也可能把生成的 PPTX 与长期图片
    # 资产放在这里。因此要求 analyse_paper 在成功解析后写入的所有权标记。
    if (-not (Test-TemporaryPaperAssetMarker -Directory $target)) {
        throw 'cleanup_paper_assets only deletes a temporary paper-asset directory created by analyse_paper; this directory has no valid ownership marker.'
    }
    # 删除前再遍历一次，拒绝期间被替换成 junction/symlink 的子项。
    foreach ($item in @(Get-ChildItem -LiteralPath $target -Force -Recurse)) {
        Assert-PaperToJournalClubNoReparsePoint -Path $item.FullName -ParameterName 'temporary paper asset cleanup item'
    }
    Remove-Item -LiteralPath $target -Recurse -Force
    return [pscustomobject]@{ deleted = $true; asset_output_dir = $target }
}

function Initialize-BoundedProcessTextReader {
    if ($null -ne ('PaperToJournalClub.BoundedTextReader' -as [type])) { return }

    # ReadToEndAsync 会在限制生效前完整分配子进程输出。这个小型托管读取器按 8 KiB 分块读取，
    # 达到阈值后立即完成任务，让父进程终止异常输出的 parser.exe。
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace PaperToJournalClub
{
    public sealed class BoundedReadResult
    {
        public string Text;
        public bool Exceeded;
    }

    public static class BoundedTextReader
    {
        public static async Task<BoundedReadResult> ReadAsync(StreamReader reader, int maximumCharacters)
        {
            var buffer = new char[8192];
            var builder = new StringBuilder(Math.Min(maximumCharacters, buffer.Length));
            while (true)
            {
                var read = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                if (read == 0)
                {
                    return new BoundedReadResult { Text = builder.ToString(), Exceeded = false };
                }
                if (builder.Length + read > maximumCharacters)
                {
                    var remaining = maximumCharacters - builder.Length;
                    if (remaining > 0) { builder.Append(buffer, 0, remaining); }
                    return new BoundedReadResult { Text = builder.ToString(), Exceeded = true };
                }
                builder.Append(buffer, 0, read);
            }
        }
    }
}
'@
}

function Invoke-PaperParserPackage {
    param([string]$PaperPath, [string]$AssetDirectory)

    if (-not (Test-Path -LiteralPath $ParserPath -PathType Leaf)) {
        throw "Bundled PDF parser was not found: $ParserPath"
    }
    # Windows 路径不允许双引号；明确拒绝后再构造进程参数，避免子进程参数被意外拆分。
    if ($PaperPath.Contains('"') -or $AssetDirectory.Contains('"')) {
        throw 'Input and asset paths may not contain quotation marks.'
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $ParserPath
    $startInfo.Arguments = 'extract-package "{0}" "{1}"' -f $PaperPath, $AssetDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) { throw 'Bundled PDF parser did not start.' }
        Initialize-BoundedProcessTextReader
        $standardOutputTask = [PaperToJournalClub.BoundedTextReader]::ReadAsync($process.StandardOutput, $MaximumParserOutputCharacters)
        $standardErrorTask = [PaperToJournalClub.BoundedTextReader]::ReadAsync($process.StandardError, $MaximumParserErrorCharacters)
        $deadline = [DateTime]::UtcNow.AddMilliseconds($ParserTimeoutMilliseconds)
        while (-not $process.HasExited) {
            if ($standardOutputTask.IsCompleted -and $standardOutputTask.GetAwaiter().GetResult().Exceeded) {
                try { $process.Kill() } catch { }
                throw "Bundled PDF parser exceeded the $MaximumParserOutputCharacters-character stdout safety limit."
            }
            if ($standardErrorTask.IsCompleted -and $standardErrorTask.GetAwaiter().GetResult().Exceeded) {
                try { $process.Kill() } catch { }
                throw "Bundled PDF parser exceeded the $MaximumParserErrorCharacters-character stderr safety limit."
            }
            if ([DateTime]::UtcNow -ge $deadline) {
                try { $process.Kill() } catch { }
                throw "PDF parsing exceeded the $([int]($ParserTimeoutMilliseconds / 1000))-second safety limit."
            }
            Start-Sleep -Milliseconds 25
        }
        if (-not $process.WaitForExit(5000)) {
            try { $process.Kill() } catch { }
            throw 'Bundled PDF parser did not finish flushing its output after exit.'
        }
        $standardOutputResult = $standardOutputTask.GetAwaiter().GetResult()
        $standardErrorResult = $standardErrorTask.GetAwaiter().GetResult()
        if ($standardOutputResult.Exceeded) { throw "Bundled PDF parser exceeded the $MaximumParserOutputCharacters-character stdout safety limit." }
        if ($standardErrorResult.Exceeded) { throw "Bundled PDF parser exceeded the $MaximumParserErrorCharacters-character stderr safety limit." }
        $standardOutput = $standardOutputResult.Text
        $standardError = $standardErrorResult.Text
        if ($process.ExitCode -ne 0) {
            throw "Bundled PDF parser failed: $standardError"
        }
        if ([string]::IsNullOrWhiteSpace($standardOutput)) { throw 'Bundled PDF parser returned no JSON package.' }
        try {
            return $standardOutput | ConvertFrom-Json
        } catch {
            throw "Bundled PDF parser returned invalid JSON. $($_.Exception.Message)"
        }
    } finally {
        if ($process) { $process.Dispose() }
    }
}

function Assert-ExtractedAssetDirectoryLimits {
    param([string]$Directory)

    $safeDirectory = Assert-PaperToJournalClubAllowedPath -Path $Directory -AllowedRoots @((Get-PaperToJournalClubTemporaryRoot)) -ParameterName 'parser asset directory'
    if (-not (Test-Path -LiteralPath $safeDirectory -PathType Container)) { throw 'Bundled PDF parser did not create its requested asset directory.' }
    $files = @(Get-ChildItem -LiteralPath $safeDirectory -File -Recurse -Force)
    if ($files.Count -gt 100) { throw 'Bundled PDF parser extracted more than the 100-file asset safety limit.' }
    [int64]$totalBytes = 0
    foreach ($file in $files) {
        Assert-PaperToJournalClubNoReparsePoint -Path $file.FullName -ParameterName 'parser asset file'
        $totalBytes += [int64]$file.Length
        if ($totalBytes -gt 200MB) { throw 'Bundled PDF parser extracted more than the 200 MB asset safety limit.' }
    }
    return $safeDirectory
}

function Find-PageNumberForText {
    param([string]$Text, $Pages)
    if (-not $Text) { return $null }
    $normalizedText = (Normalize-Text $Text -replace '\s+', ' ').Trim()
    if (-not $normalizedText) { return $null }

    # 图号附近的上下文可能恰好跨页。不要因一段 380 字符的上下文无法完整出现在某页而
    # 丢失页码；改用若干足够长的重叠片段定位，仍避免用过短的常见词作不可靠匹配。
    $needles = @()
    if ($normalizedText.Length -le 180) {
        $needles += $normalizedText
    } else {
        for ($offset = 0; $offset -lt $normalizedText.Length; $offset += 90) {
            $length = [Math]::Min(160, $normalizedText.Length - $offset)
            if ($length -lt 48) { break }
            $needles += $normalizedText.Substring($offset, $length)
            if ($offset + $length -ge $normalizedText.Length) { break }
        }
    }
    foreach ($page in @($Pages)) {
        $pageText = (Normalize-Text ([string](Get-PropertyValue $page 'Text' '')) -replace '\s+', ' ').Trim()
        foreach ($needle in $needles) {
            if ($pageText.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return Get-PageNumberValue $page
            }
        }
    }
    return $null
}

function Get-PaperExtraction {
    param([string]$FilePath, [string]$AssetOutputDirectory)
    $absolutePath = Get-ValidatedPaperPath $FilePath
    $extension = [IO.Path]::GetExtension($absolutePath).ToLowerInvariant()
    if ($extension -in @('.txt', '.md', '.tex')) {
        $text = Normalize-Text (Get-Content -Raw -Encoding UTF8 -LiteralPath $absolutePath)
        if ($text.Length -gt $MaximumExtractedTextCharacters) { throw "Paper text exceeds the $MaximumExtractedTextCharacters-character safety limit." }
        return [pscustomobject]@{
            text = $text
            pages = @([pscustomobject]@{ PageNumber = 1; Text = $text; Assets = @() })
            asset_directory = $null
            assets_truncated = $false
            extraction_method = 'plain-text'
        }
    }
    if ($extension -ne '.pdf') { throw "Supported input files are PDF, TXT, Markdown, and TeX." }
    if (Test-Path -LiteralPath $ParserPath) {
        $assetDirectory = New-PaperAssetDirectory $absolutePath $AssetOutputDirectory
        try {
            # extract-package 会同时给出逐页文本与真实导出的图片资产；图号与图片的对应关系不在这里臆测。
            $package = Invoke-PaperParserPackage -PaperPath $absolutePath -AssetDirectory $assetDirectory
            $reportedAssetDirectory = Get-PropertyValue $package 'asset_directory' $assetDirectory
            $safeAssetDirectory = Assert-ExtractedAssetDirectoryLimits $reportedAssetDirectory
            if (-not $safeAssetDirectory.Equals($assetDirectory, [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Bundled PDF parser returned an unexpected asset directory.'
            }
            $text = Normalize-Text (Get-PropertyValue $package 'text' '')
            if ($text.Length -gt $MaximumExtractedTextCharacters) { throw "Extracted PDF text exceeds the $MaximumExtractedTextCharacters-character safety limit." }
            if ($text.Length -ge 120) {
                Write-TemporaryPaperAssetMarker -Directory $safeAssetDirectory
                return [pscustomobject]@{
                    text = $text
                    pages = @(Get-PropertyValue $package 'pages' @())
                    asset_directory = $safeAssetDirectory
                    assets_truncated = Get-StrictBoolean -Object $package -Name 'assets_truncated' -Default $false
                    extraction_method = 'paper-parser-package'
                }
            }
            throw 'The PDF contains too little extractable text. It may be a scan and requires OCR.'
        } catch {
            # 发布包已携带解析器；失败时直接报告，避免再启动无法设置超时的 Word COM 回退。
            throw "Bundled PDF parser could not safely extract this paper. $($_.Exception.Message)"
        }
    }
    # Word COM 不能可靠地终止卡住的 PDF 打开操作，因此发布版默认不把它当作自动回退。
    # 仅由管理员在启动 Codex 前显式设置环境变量后才启用，且仍保持宏禁用设置。
    if ($env:PAPER_TO_JOURNAL_CLUB_ALLOW_WORD_PDF_FALLBACK -ne '1') {
        throw 'Bundled PDF parser is unavailable. Word COM PDF fallback is disabled by default because it has no reliable watchdog; restore paper-parser.exe or explicitly set PAPER_TO_JOURNAL_CLUB_ALLOW_WORD_PDF_FALLBACK=1 before starting Codex.'
    }
    try {
        $wordText = Get-WordPdfText $absolutePath
        if ($wordText.Length -ge 120) {
            return [pscustomobject]@{ text = $wordText; pages = @(); asset_directory = $null; extraction_method = 'word-com-fallback' }
        }
    } catch {
        throw "PDF extraction failed. Install the released paper-parser.exe asset, or install Microsoft Word for the Office fallback. Details: $($_.Exception.Message)"
    }
    throw "The PDF contains too little extractable text. It may be a scan and requires OCR."
}

function Get-CanonicalPaperSectionName {
    param([string]$Title)
    switch -Regex ($Title.Trim().ToLowerInvariant()) {
        '^(abstract|摘要)$' { return 'abstract' }
        '^(introduction|引言|前言|研究背景)$' { return 'introduction' }
        '^(method|methods|materials? and methods?|方法|材料与方法|研究方法)$' { return 'methods' }
        '^(result|results|结果)$' { return 'results' }
        '^(discussion|讨论)$' { return 'discussion' }
        '^(conclusion|conclusions|结论)$' { return 'conclusions' }
        '^(references|参考文献)$' { return 'references' }
        default { return 'section' }
    }
}

function Get-PaperSections {
    param([string]$Text)
    # 兼容常见中英文标题及“1. Introduction”类编号标题；输出 ID 仍保持稳定的英文类别。
    $pattern = '^(?:\s*\d+(?:\.\d+)*[.)]?\s*)?(abstract|introduction|methods?|materials? and methods?|results?|discussion|conclusions?|references|摘要|引言|前言|研究背景|方法|材料与方法|研究方法|结果|讨论|结论|参考文献)\s*$'
    $headingMatches = [regex]::Matches($Text, $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Multiline)
    if ($headingMatches.Count -eq 0) { return @([pscustomobject]@{ id = 'full-text'; title = 'Full text'; text = $Text }) }
    $sections = @()
    for ($index = 0; $index -lt $headingMatches.Count; $index++) {
        $start = $headingMatches[$index].Index + $headingMatches[$index].Length
        $end = if ($index + 1 -lt $headingMatches.Count) { $headingMatches[$index + 1].Index } else { $Text.Length }
        $body = Normalize-Text $Text.Substring($start, $end - $start)
        if ($body.Length -gt 0) {
            $title = $headingMatches[$index].Groups[1].Value
            $canonicalTitle = Get-CanonicalPaperSectionName $title
            $sections += [pscustomobject]@{ id = "$canonicalTitle-$($index + 1)"; title = $title; text = $body }
        }
    }
    return $sections
}

function Get-PaperMetadata {
    param([string]$Text, [string]$FallbackName)
    $lines = @((Normalize-Text $Text) -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -First 35)
    $title = $lines | Where-Object { $_.Length -ge 15 -and $_.Length -le 220 -and $_ -notmatch '^(abstract|摘要)$' } | Select-Object -First 1
    if (-not $title) { $title = $FallbackName }
    $doiMatch = [regex]::Match($Text, '10\.\d{4,9}\/[\w.()/:;-]+', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $yearMatch = [regex]::Match($Text, '\b(19|20)\d{2}\b')
    return [pscustomobject]@{ title = $title; doi = if ($doiMatch.Success) { $doiMatch.Value } else { $null }; year = if ($yearMatch.Success) { $yearMatch.Value } else { $null } }
}

function Get-FigureReferences {
    param([string]$Text, $Pages)
    $figureMatches = [regex]::Matches($Text, '(?:\b(?:fig(?:ure)?s?\.?)\s*|图\s*)(\d+)([a-zA-Z])?', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # 同一个图号通常会在正文、图注和讨论中反复出现。保留所有出现页，才能优先把
    # 图号关联到真正含有导出位图的页面，而不是机械地使用第一次正文引用所在的页。
    $byId = @{}
    $orderedIds = @()
    foreach ($match in $figureMatches) {
        $id = "fig-$($match.Groups[1].Value)$($match.Groups[2].Value)".ToLowerInvariant()
        # 只限制“不同图号”的数量，不能在达到 30 个图号后停止收集已有图号的后续
        # 图注出现；否则反而会丢失真正包含图像资产的页码。
        if (-not $byId.ContainsKey($id)) {
            if ($orderedIds.Count -ge 30) { continue }
            $byId[$id] = @()
            $orderedIds += $id
        }
        # 优先截取同一行/紧邻图号的图注片段。不能把整段 Results 或下一段 Examples
        # 混入分类依据，否则会将 Figure 4 的性能曲线误判成案例图，或把相似 meme 图
        # 误判为主实验结果。
        $captionStart = $match.Index
        while ($captionStart -gt 0 -and $Text[$captionStart - 1] -notin @("`n", "`r")) { $captionStart-- }
        $captionEnd = $match.Index + $match.Length
        while ($captionEnd -lt $Text.Length -and $Text[$captionEnd] -notin @("`n", "`r")) { $captionEnd++ }
        $captionLength = [Math]::Min(520, [Math]::Max(0, $captionEnd - $captionStart))
        $context = if ($captionLength -gt 0) {
            Normalize-Text $Text.Substring($captionStart, $captionLength)
        } else {
            Normalize-Text $match.Value
        }
        $occurrence = [pscustomobject]@{
            label = $match.Value
            context = $context
            source_page = Find-PageNumberForText $context $Pages
        }
        $byId[$id] += $occurrence
    }

    $figures = @()
    foreach ($id in $orderedIds) {
        $occurrences = @($byId[$id])
        $sourcePages = @($occurrences | ForEach-Object { $_.source_page } | Where-Object { $null -ne $_ } | Select-Object -Unique)
        # 若一个图号在多个页面被引用，优先选择“仅有的含原始位图页面”。这只是缩小候选，
        # 后续仍要求该页只有一个图号和一张合规图片才会自动插入。
        $pagesWithAssets = @($sourcePages | Where-Object { @((Get-PageAssetsForNumber -Pages $Pages -PageNumber ([int]$_))).Count -gt 0 })
        $sourcePage = if ($pagesWithAssets.Count -eq 1) {
            [int]$pagesWithAssets[0]
        } elseif ($sourcePages.Count -eq 1) {
            [int]$sourcePages[0]
        } else {
            $null
        }
        $selectedOccurrence = @($occurrences | Where-Object { $_.source_page -eq $sourcePage } | Select-Object -First 1)
        $fallbackOccurrence = @($occurrences | Select-Object -First 1)
        $representativeOccurrence = if ($selectedOccurrence.Count) { $selectedOccurrence[0] } else { $fallbackOccurrence[0] }
        $figures += [pscustomobject]@{
            id = $id
            label = $representativeOccurrence.label
            context = $representativeOccurrence.context
            source_page = $sourcePage
            source_pages = @($sourcePages)
        }
    }
    return $figures
}

function Get-TableReferences {
    param([string]$Text, $Pages)

    # 表格通常是 PDF 的文字和矢量线条，PdfPig 无法把它当作嵌入位图导出。这里先保留
    # 表号、上下文和页码，后续只允许用户确认后的页面裁剪进入实验数据页。
    $tableMatches = [regex]::Matches($Text, '(?:\b(?:table|tab\.)\s*|表\s*)(\d+)([a-zA-Z])?', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $byId = @{}
    $orderedIds = @()
    foreach ($match in $tableMatches) {
        $id = "table-$($match.Groups[1].Value)$($match.Groups[2].Value)".ToLowerInvariant()
        if (-not $byId.ContainsKey($id)) {
            if ($orderedIds.Count -ge 30) { continue }
            $byId[$id] = @()
            $orderedIds += $id
        }
        # 表格与 Figure 一样使用本行标题/图注，避免邻近正文里无关的关键词改变角色。
        $captionStart = $match.Index
        while ($captionStart -gt 0 -and $Text[$captionStart - 1] -notin @("`n", "`r")) { $captionStart-- }
        $captionEnd = $match.Index + $match.Length
        while ($captionEnd -lt $Text.Length -and $Text[$captionEnd] -notin @("`n", "`r")) { $captionEnd++ }
        $captionLength = [Math]::Min(520, [Math]::Max(0, $captionEnd - $captionStart))
        $context = if ($captionLength -gt 0) {
            Normalize-Text $Text.Substring($captionStart, $captionLength)
        } else {
            Normalize-Text $match.Value
        }
        $byId[$id] += [pscustomobject]@{
            label = $match.Value
            context = $context
            source_page = Find-PageNumberForText $context $Pages
        }
    }

    $tables = @()
    foreach ($id in $orderedIds) {
        $occurrences = @($byId[$id])
        $sourcePages = @($occurrences | ForEach-Object { $_.source_page } | Where-Object { $null -ne $_ } | Select-Object -Unique)
        $sourcePage = if ($sourcePages.Count -eq 1) { [int]$sourcePages[0] } else { $null }
        $tables += [pscustomobject]@{
            id = $id
            label = [string]$occurrences[0].label
            context = [string]$occurrences[0].context
            source_page = $sourcePage
            source_pages = @($sourcePages)
            selection_mode = 'user-confirmed-page-crop-required'
        }
    }
    return @($tables)
}

function Get-JournalClubVisualClassification {
    <#
      组会主线只需要帮助听众理解“怎么做、效果如何、模块是否必要”的图表。
      本函数故意不尝试从所有 PDF 图片中猜测用途：只在图注或邻近文本出现强信号时
      才放行。无法可靠判断时，默认不渲染、不插入，交给用户显式审阅。
    #>
    param(
        [ValidateSet('figure', 'table')]
        [string]$VisualKind,
        [string]$CaptionOrContext
    )

    $text = Normalize-Text $CaptionOrContext
    $unknown = [pscustomobject]@{
        journal_club_role = 'excluded-or-unknown'
        eligible_for_journal_club_visual = $false
        selection_guidance = '未能从图注或邻近文本可靠判定为方法、主实验或消融图表，默认不自动渲染或插入。'
    }
    if ([string]::IsNullOrWhiteSpace($text)) { return $unknown }

    # 数据集规模、类别分布一类表格通常用于实验设置，而不是结果分析。它们保留为
    # 证据对象，但不占用默认的“主实验/消融”视觉位置。
    $datasetStatisticsPattern = '(?i)\b(dataset|data set|corpus)\s*(?:statistics|overview|description|profile)\b|\b(?:statistics|distribution)\s+of\s+(?:the\s+)?(?:dataset|data set|corpus)\b|\bclass(?:\s+label)?\s+distribution\b|数据集统计|数据集概览|样本统计|类别分布|数据分布'
    if ($VisualKind -eq 'table' -and $text -match $datasetStatisticsPattern) {
        return [pscustomobject]@{
            journal_club_role = 'excluded-or-unknown'
            eligible_for_journal_club_visual = $false
            selection_guidance = '这是数据集或样本统计表，默认不作为方法、主实验或消融实验的插图；可在实验设置需要时由用户另行选择。'
        }
    }

    # Top-K、超参数等“效果曲线”通常包含词 retrieval，但那是方法模块的名称，
    # 不是相似样例展示。先识别这一类定量趋势，避免被下方的案例关键词吞掉。
    $quantitativeEffectPattern = '(?i)\b(?:effect|impact|sensitivity|influence)\s+of\s+(?:top\s*-?\s*k|k\b|parameter|hyperparameter|threshold|number of)\b|\b(?:performance|metric)\s+(?:curve|trend|versus|vs\.?)\b|\btop\s*-?\s*k\b|Top\s*-?\s*K|参数(?:敏感性|影响)|性能曲线|性能趋势|不同(?:TopK|Top-K|参数)'
    if ($text -match $quantitativeEffectPattern) {
        return [pscustomobject]@{
            journal_club_role = 'main-result'
            eligible_for_journal_club_visual = $true
            selection_guidance = '可作为主实验的参数敏感性或性能趋势图；应解释横纵轴、最佳区间以及为何该设置被采用。'
        }
    }

    # 案例、错误预测与相似样例容易让汇报偏离方法和可量化结果。定量参数曲线优先于
    # “Similar Sample Retrieval”这种模块名；其余案例即使出现 performance 也必须排除。
    $casePattern = '(?i)\b(example|examples|case|cases|case study|qualitative|correct(?:ly)?|wrong(?:ly)?|misclassif(?:y|ication)|failure cases?|error analysis|similar meme|similar samples?|retriev(?:ed|al)\s+(?:meme|sample)s?|illustrative|visuali[sz]ation)\b|\bsimilar\s+memes?\s+(?:retriev(?:ed|al)|for|from|of|by)\b|案例|示例|样例|正确预测|错误预测|误判|错误分析|失败案例|相似(?:模因|样本|图片)|相似.*meme'
    if ($text -match $casePattern) {
        return [pscustomobject]@{
            journal_club_role = 'case-analysis'
            eligible_for_journal_club_visual = $false
            selection_guidance = '这是案例、错误预测或相似样例内容，默认不用于组会主线图表；只有用户明确要求案例讨论时才单独处理。'
        }
    }

    $ablationPattern = '(?i)\b(ablation|ablate|component analysis|module analysis|effect of (?:each |individual )?(?:component|module)|without (?:the )?(?:module|component)|remov(?:e|al) of)\b|消融|模块分析|组件分析|去除(?:模块|组件)'
    if ($text -match $ablationPattern) {
        return [pscustomobject]@{
            journal_club_role = 'ablation'
            eligible_for_journal_club_visual = $true
            selection_guidance = '可作为消融实验图表；幻灯片应说明移除、替换或比较各模块后，哪些指标发生变化及其边界。'
        }
    }

    # 系统结构角色只授予 Figure。表格即使出现在方法章节，也不足以证明它是可讲解的
    # 系统框图，避免把参数表误认为方法示意图。
    $methodsPattern = '(?i)\b(?:system|model|framework|method|approach)\s+(?:architecture|overview|workflow|pipeline)\b|\b(?:architecture|workflow|pipeline|method overview|framework overview|system overview|model overview|proposed framework|overview of (?:our|the) (?:framework|method|approach))\b|系统结构|整体框架|框架图|流程图|方法概览|模型结构|系统概览|工作流程'
    if ($VisualKind -eq 'figure' -and $text -match $methodsPattern) {
        return [pscustomobject]@{
            journal_club_role = 'methods-system'
            eligible_for_journal_club_visual = $true
            selection_guidance = '可作为方法页系统结构图；应结合方法原文讲清输入、关键模块、信息流、输出或验证步骤。'
        }
    }

    $resultPattern = '(?i)\b(?:main|overall|primary)?\s*(?:results?|performance|evaluation|benchmark|comparison|comparisons?)\b|\b(?:accuracy|macro[- ]?f1|f1[- ]?score|auc|precision|recall)\b|实验结果|性能比较|性能评估|基准测试|主实验|对比实验|主要结果|模型性能|(?:显著|明显)?(?:提高|提升|改善).*(?:预测|性能|准确率|指标)|(?:反应|通路|表型).*(?:增加|降低|改善|提高)'
    if ($text -match $resultPattern) {
        return [pscustomobject]@{
            journal_club_role = 'main-result'
            eligible_for_journal_club_visual = $true
            selection_guidance = '可作为主实验图表；幻灯片应依据图注和结果原文说明比较对象、指标方向、优势或不足，避免只呈现正面结果。'
        }
    }

    return $unknown
}

function Set-JournalClubVisualClassification {
    <#
      无论 evidence pack 来自本插件解析器、旧版本缓存还是 JSON 往返，本步骤都重新根据
      原始图注/邻近文本计算资格。不能信任调用方手写的 eligible=true，以免案例图通过
      一个布尔字段绕过默认排除规则。
    #>
    param(
        $Visuals,
        [ValidateSet('figure', 'table')]
        [string]$VisualKind
    )

    $items = @($Visuals | Where-Object { $null -ne $_ })
    foreach ($item in $items) {
        $classification = Get-JournalClubVisualClassification -VisualKind $VisualKind -CaptionOrContext ([string](Get-PropertyValue $item 'context' ''))
        $item | Add-Member -NotePropertyName 'journal_club_role' -NotePropertyValue $classification.journal_club_role -Force
        $item | Add-Member -NotePropertyName 'eligible_for_journal_club_visual' -NotePropertyValue ([bool]$classification.eligible_for_journal_club_visual) -Force
        $item | Add-Member -NotePropertyName 'selection_guidance' -NotePropertyValue $classification.selection_guidance -Force
        # 把分类所依据的原文片段与图表实体放在一起，后续 deck spec 可以同时追溯
        # 图号、页码、角色判断和论文原文，而不是只留下一个图像文件路径。
        $item | Add-Member -NotePropertyName 'classification_evidence' -NotePropertyValue ([pscustomobject]@{
            source_page = Get-PropertyValue $item 'source_page'
            excerpt = ConvertTo-ReadableSlideText -Text ([string](Get-PropertyValue $item 'context' '')) -ChineseLimit 120 -OtherLimit 360
            basis = 'caption-or-nearby-text-keywords'
        }) -Force
    }
    return @($items)
}

function Test-JournalClubVisualEligibility {
    param(
        $Visual,
        [string[]]$AllowedRoles = @()
    )

    if ($null -eq $Visual) { return $false }
    $visualId = [string](Get-PropertyValue $Visual 'id' '')
    $visualKind = if ($visualId -like 'table-*') { 'table' } else { 'figure' }
    # 重新依图注/邻近文本判断，绝不信任手工 evidence pack 填写的 true。这样旧版
    # evidence pack 缺少新字段时仍可安全兼容，而伪造 eligible=true 也不会放行。
    $classification = Get-JournalClubVisualClassification -VisualKind $visualKind -CaptionOrContext ([string](Get-PropertyValue $Visual 'context' ''))
    if (-not $classification.eligible_for_journal_club_visual) { return $false }
    $role = [string]$classification.journal_club_role
    if ($role -notin @('methods-system', 'main-result', 'ablation')) { return $false }
    if (@($AllowedRoles).Count -gt 0 -and $role -notin $AllowedRoles) { return $false }
    return $true
}

function Get-JournalClubVisualTraceability {
    <#
      设计阶段不仅需要一个 bool，还需要将角色判断的原文依据带到 slide。对于旧版或
      用户手工构造的 evidence pack，补建最小证据对象，避免产生无法回答“这张图从
      哪里来、为什么被选中”的 deck spec。
    #>
    param($Visual)

    if ($null -eq $Visual) {
        return [pscustomobject]@{ visual_id = ''; role = ''; source_page = 0; evidence = $null }
    }
    $visualId = [string](Get-PropertyValue $Visual 'id' '')
    $visualKind = if ($visualId -like 'table-*') { 'table' } else { 'figure' }
    $classification = Get-JournalClubVisualClassification -VisualKind $visualKind -CaptionOrContext ([string](Get-PropertyValue $Visual 'context' ''))
    $sourcePage = Get-PropertyValue $Visual 'source_page' 0
    try { $sourcePage = [int]$sourcePage } catch { $sourcePage = 0 }
    $evidence = Get-PropertyValue $Visual 'classification_evidence'
    if ($null -eq $evidence) {
        $evidence = [pscustomobject]@{
            source_page = if ($sourcePage -ge 1) { $sourcePage } else { $null }
            excerpt = ConvertTo-ReadableSlideText -Text ([string](Get-PropertyValue $Visual 'context' '')) -ChineseLimit 120 -OtherLimit 360
            basis = 'caption-or-nearby-text-keywords'
        }
    }
    return [pscustomobject]@{
        visual_id = $visualId
        role = [string]$classification.journal_club_role
        source_page = $sourcePage
        evidence = $evidence
    }
}

function Get-PageNumberValue {
    param($Page)

    $value = Get-PropertyValue $Page 'PageNumber'
    if ($null -eq $value) { $value = Get-PropertyValue $Page 'page_number' }
    if ($null -eq $value) { return $null }
    try {
        $pageNumber = [int]$value
        if ($pageNumber -ge 1) { return $pageNumber }
    } catch { }
    return $null
}

function Get-PageAssetsForNumber {
    param($Pages, [int]$PageNumber)

    if ($PageNumber -lt 1) { return @() }
    foreach ($page in @($Pages)) {
        if ((Get-PageNumberValue $page) -ne $PageNumber) { continue }
        $assets = Get-PropertyValue $page 'Assets'
        if ($null -eq $assets) { $assets = Get-PropertyValue $page 'assets' @() }
        return @($assets)
    }
    return @()
}

function Get-PageAssetCandidates {
    param($Pages, [int]$PageNumber, [string]$AssetDirectory)

    return @(
        Get-PageRasterAssetCandidates -Pages $Pages -PageNumber $PageNumber -AssetDirectory $AssetDirectory |
            ForEach-Object { [string]$_.path }
    )
}

function Get-PageRasterAssetCandidates {
    <#
      返回“页面中的已验证位图资产”，而不仅是路径。asset_id 是解析器为 PDF 内嵌图像
      分配的稳定编号，例如 page-03-image-01。保留它可让 evidence pack 明确展示
      Fig. 2 与实际插入 PowerPoint 的哪张原始图对应，避免后续只能按文件名猜测。
    #>
    param($Pages, [int]$PageNumber, [string]$AssetDirectory)

    if ($PageNumber -lt 1 -or [string]::IsNullOrWhiteSpace($AssetDirectory)) { return @() }
    $candidates = @()
    $seenPaths = @{}
    foreach ($asset in @(Get-PageAssetsForNumber -Pages $Pages -PageNumber $PageNumber)) {
        $path = Get-PropertyValue $asset 'Path'
        if ([string]::IsNullOrWhiteSpace($path)) { $path = Get-PropertyValue $asset 'path' }
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        try {
            # 导出图片来自不可信 PDF；验证扩展名、签名、像素和临时目录边界后才允许成为候选。
            $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath ([string]$path) -AllowedRoots @($AssetDirectory) -ParameterName 'parser-extracted figure asset'
            if ($seenPaths.ContainsKey($image.path)) { continue }
            $seenPaths[$image.path] = $true
            $assetId = [string](Get-PropertyValue $asset 'Id' '')
            if ([string]::IsNullOrWhiteSpace($assetId)) { $assetId = [string](Get-PropertyValue $asset 'id' '') }
            $candidates += [pscustomobject]@{
                asset_id = $assetId.Trim()
                source_page = $PageNumber
                path = $image.path
                bytes = [int64]$image.bytes
            }
        } catch {
            # 单张图片坏掉、过大或在解析后被删除时，不影响正文证据包；只是不把它插入 PPT。
        }
    }
    return @($candidates)
}

function Add-FigureAssetCandidates {
    param($Figures, $Pages, [string]$AssetDirectory)

    $figureItems = @($Figures)
    foreach ($figure in $figureItems) {
        $pageNumber = Get-PropertyValue $figure 'source_page'
        $safePageNumber = if ($null -ne $pageNumber) { [int]$pageNumber } else { 0 }
        $candidateAssets = @(Get-PageRasterAssetCandidates -Pages $Pages -PageNumber $safePageNumber -AssetDirectory $AssetDirectory)
        $candidates = @($candidateAssets | ForEach-Object { [string]$_.path })
        $pageFigures = if ($safePageNumber -ge 1) {
            @($figureItems | Where-Object { (Get-PropertyValue $_ 'source_page') -eq $safePageNumber })
        } else {
            @()
        }

        $matchKind = if ($candidates.Count -eq 0) {
            'none'
        } elseif ($pageFigures.Count -eq 1 -and $candidates.Count -eq 1) {
            'same-page-single-raster'
        } else {
            'ambiguous'
        }
        $automaticImagePath = if ($matchKind -eq 'same-page-single-raster') { $candidates[0] } else { $null }
        $automaticBinding = $null
        if ($automaticImagePath -and $candidateAssets.Count -eq 1) {
            $assetId = [string]$candidateAssets[0].asset_id
            $figureId = [string](Get-PropertyValue $figure 'id' '')
            # 旧 parser package 或手工 evidence pack 可能没有 asset id。保留原有自动插图兼容，
            # 但不伪造一条不可审计的“编号到资产”映射。
            if (-not [string]::IsNullOrWhiteSpace($assetId) -and -not [string]::IsNullOrWhiteSpace($figureId)) {
                $automaticBinding = [pscustomobject]@{
                    figure_id = $figureId
                    asset_id = $assetId
                    source_page = $safePageNumber
                    path = $automaticImagePath
                    mode = 'automatic-figure-number-single-raster'
                }
            }
        }

        # 候选路径始终保留在 evidence pack，供用户在多图页手动确认；只有一图一资产时才写入
        # automatic_image_path，后续设计步骤绝不把候选列表的第一张图当作事实关联。
        $figure | Add-Member -NotePropertyName 'figure_asset_candidates' -NotePropertyValue @($candidates) -Force
        $figure | Add-Member -NotePropertyName 'asset_match' -NotePropertyValue $matchKind -Force
        foreach ($transientProperty in @('automatic_image_path', 'automatic_binding')) {
            if ($figure.PSObject.Properties[$transientProperty]) { [void]$figure.PSObject.Properties.Remove($transientProperty) }
        }
        if ($automaticImagePath) {
            $figure | Add-Member -NotePropertyName 'automatic_image_path' -NotePropertyValue $automaticImagePath -Force
        }
        if ($null -ne $automaticBinding) {
            $figure | Add-Member -NotePropertyName 'automatic_binding' -NotePropertyValue $automaticBinding -Force
        }
    }
    return @($figureItems)
}

function Get-ClaimComparisonKind {
    param([string]$Text)

    # 只有论文原文直接出现比较性措辞时，才允许在汇报中写“更优”“未改善”或“无明确差异”。
    # 单独的 increase/decrease 不能说明指标方向是否有利，因此归为观察而非好坏判断。
    if ($Text -match '(?i)inferior|worse\s+than|failed\s+to|no\s+improvement|did\s+not\s+improve|较差|劣于|未改善|没有改善|效果不佳') {
        return 'reported-worse-or-no-benefit'
    }
    if ($Text -match '(?i)no\s+(?:statistically\s+)?significant(?:\s+differences?)?|not\s+significant|did\s+not\s+differ|no\s+difference|无显著(?:性)?差异|未见显著差异|无明显差异') {
        return 'no-clear-difference'
    }
    if ($Text -match '(?i)outperform(?:ed|s|ing)?|superior|better\s+than|improv(?:e|ed|ement)|benefit(?:ed|s)?|优于|更好|改善|疗效提高|表现更佳') {
        return 'reported-better'
    }
    return 'reported-observation'
}

function Get-CandidateClaims {
    param($Sections, $Pages)
    $resultSections = @($Sections | Where-Object { $_.title -match 'result|结果' })
    $sourceSections = if ($resultSections.Count) { $resultSections } else { $Sections }
    $claims = @()
    foreach ($section in $sourceSections) {
        foreach ($sentence in (Split-PaperSentences $section.text)) {
            # 结果候选不能只收集“改善/提高”。显式的不利、无改善或失败结果同样是组会
            # 需要讨论的证据，后续会与正面比较一起进入短汇报叙事。
            if ($sentence -notmatch 'significant|increase|decrease|improve|inferior|worse|failed|no\s+improvement|associated|demonstrate|show|suggest|support|显著|增加|减少|提高|降低|改善|未改善|没有改善|较差|劣于|效果不佳|失败|相关|表明|显示|提示|支持') { continue }
            $claims += [pscustomobject]@{
                id = "claim-{0:D2}" -f ($claims.Count + 1)
                text = $sentence
                evidence = @([pscustomobject]@{
                    section_id = $section.id
                    section_title = $section.title
                    page_number = Find-PageNumberForText $sentence $Pages
                    excerpt = $sentence
                })
                confidence = 'needs-review'
                comparison_kind = Get-ClaimComparisonKind $sentence
                note = 'Automatically extracted candidate claim. Verify the source wording before making slides.'
            }
            if ($claims.Count -ge 8) { return $claims }
        }
    }
    return $claims
}

function Invoke-AnalysePaper {
    param($Arguments)
    $filePath = Get-PropertyValue $Arguments 'file_path'
    if (-not $filePath) { throw 'file_path is required.' }
    $assetOutputDirectory = Get-PropertyValue $Arguments 'asset_output_dir'
    $extraction = Get-PaperExtraction $filePath $assetOutputDirectory
    $text = $extraction.text
    $pages = @(Get-PropertyValue $extraction 'pages' @())
    $sections = Get-PaperSections $text
    $fallback = [IO.Path]::GetFileNameWithoutExtension($filePath)
    $assetDirectory = Get-PropertyValue $extraction 'asset_directory'
    $figures = @(Get-FigureReferences $text $pages)
    $figures = @(Add-FigureAssetCandidates -Figures $figures -Pages $pages -AssetDirectory $assetDirectory)
    $tables = @(Get-TableReferences $text $pages)
    # 在 evidence pack 阶段固化图表的组会角色。后续设计、渲染和插图还会重复核验，
    # 但先把判定、依据与建议返回给调用方，方便用户看到为何某张图没有被默认采用。
    $figures = @(Set-JournalClubVisualClassification -Visuals $figures -VisualKind 'figure')
    $tables = @(Set-JournalClubVisualClassification -Visuals $tables -VisualKind 'table')
    return [pscustomobject]@{
        schema_version = '0.3'
        source_file = [IO.Path]::GetFullPath($filePath)
        paper = Get-PaperMetadata $text $fallback
        extraction = [pscustomobject]@{
            method = Get-PropertyValue $extraction 'extraction_method'
            asset_directory = Get-PropertyValue $extraction 'asset_directory'
            assets_truncated = Get-StrictBoolean -Object $extraction -Name 'assets_truncated' -Default $false
            pages = @($pages | ForEach-Object {
                [pscustomobject]@{
                    page_number = Get-PropertyValue $_ 'PageNumber'
                    excerpt = ((Get-PropertyValue $_ 'Text' '')).Substring(0, [Math]::Min(1200, ((Get-PropertyValue $_ 'Text' '')).Length))
                    assets = @((Get-PropertyValue $_ 'Assets' @()) | ForEach-Object {
                        [pscustomobject]@{
                            id = Get-PropertyValue $_ 'Id'
                            page_number = Get-PropertyValue $_ 'PageNumber'
                            path = Get-PropertyValue $_ 'Path'
                            bytes = Get-PropertyValue $_ 'Bytes'
                        }
                    })
                }
            })
        }
        sections = @($sections | ForEach-Object {
            $sectionExcerptLimit = if ($_.title -match 'method|material|方法') { 4000 } else { 1600 }
            [pscustomobject]@{
                id = $_.id
                title = $_.title
                source_page = Find-PageNumberForText $_.text $pages
                # 方法段保留更长的可回链摘录，以便系统结构页解释输入、模块、输出与验证，
                # 同时保持整份 evidence pack 的大小上限。
                excerpt = $_.text.Substring(0, [Math]::Min($sectionExcerptLimit, $_.text.Length))
            }
        })
        claims = Get-CandidateClaims $sections $pages
        figures = $figures
        tables = $tables
        journal_club_defaults = [pscustomobject]@{ required_sections = $DefaultRequiredSections }
        ambiguities = @('Candidate claims are not publication-ready facts until a reviewer confirms the original paper.')
    }
}

function Get-VisualEntityById {
    param($EvidencePack, [string]$VisualId)

    foreach ($collectionName in @('figures', 'tables')) {
        foreach ($item in @(Get-PropertyValue $EvidencePack $collectionName @())) {
            if ([string](Get-PropertyValue $item 'id' '') -eq $VisualId) { return $item }
        }
    }
    return $null
}

function Get-RenderedVisualDirectory {
    param($EvidencePack)

    $assetDirectory = Get-SafeEvidenceAssetDirectory $EvidencePack
    if ([string]::IsNullOrWhiteSpace($assetDirectory)) {
        throw 'render_paper_visual requires an evidence_pack produced from a PDF with a plugin-owned temporary asset directory.'
    }
    if (-not (Test-TemporaryPaperAssetMarker -Directory $assetDirectory)) {
        throw 'render_paper_visual requires a valid temporary paper-asset directory created by analyse_paper.'
    }
    $directory = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $assetDirectory 'rendered-visuals') -AllowedRoots @($assetDirectory) -ParameterName 'rendered visual directory'
    if (Test-Path -LiteralPath $directory) {
        Assert-PaperToJournalClubNoReparsePoint -Path $directory -ParameterName 'rendered visual directory'
    } else {
        New-Item -ItemType Directory -Path $directory -ErrorAction Stop | Out-Null
        Assert-PaperToJournalClubAllowedPath -Path $directory -AllowedRoots @($assetDirectory) -ParameterName 'rendered visual directory'
    }
    return $directory
}

function Invoke-PdfVisualRenderer {
    param(
        [string]$PdfPath,
        [int]$PageNumber,
        [string]$OutputPath,
        [int]$Width,
        [int]$Height,
        $Crop
    )

    if (-not (Test-Path -LiteralPath $PdfVisualRendererPath -PathType Leaf)) {
        throw "Bundled PDF visual renderer was not found: $PdfVisualRendererPath"
    }
    foreach ($path in @($PdfPath, $OutputPath, $PdfVisualRendererPath)) {
        if ($path.Contains('"')) { throw 'Renderer paths may not contain quotation marks.' }
    }

    $arguments = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $PdfVisualRendererPath,
        '-PdfPath', $PdfPath, '-PageNumber', [string]$PageNumber, '-OutputPath', $OutputPath,
        '-Width', [string]$Width, '-Height', [string]$Height)
    if ($null -ne $Crop) {
        # 子进程参数必须使用不受本机小数点区域设置影响的表示法；例如中文系统不能把
        # 0.25 序列化为 PowerShell 参数解析器无法识别的 0,25。
        $invariantCulture = [Globalization.CultureInfo]::InvariantCulture
        $arguments += @('-CropX', [Convert]::ToString([double](Get-PropertyValue $Crop 'x'), $invariantCulture), '-CropY', [Convert]::ToString([double](Get-PropertyValue $Crop 'y'), $invariantCulture),
            '-CropWidth', [Convert]::ToString([double](Get-PropertyValue $Crop 'width'), $invariantCulture), '-CropHeight', [Convert]::ToString([double](Get-PropertyValue $Crop 'height'), $invariantCulture))
    }
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'powershell.exe'
    # ProcessStartInfo.ArgumentList 在 Windows PowerShell 5.1 不可用。每项均是服务端生成的
    # 数字或已拒绝双引号的路径，因此以双引号包裹后不会发生参数注入。
    $startInfo.Arguments = ($arguments | ForEach-Object {
        if ($_ -match '\s') { '"{0}"' -f $_ } else { $_ }
    }) -join ' '
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) { throw 'PDF visual renderer did not start.' }
        Initialize-BoundedProcessTextReader
        $stdoutTask = [PaperToJournalClub.BoundedTextReader]::ReadAsync($process.StandardOutput, $MaximumPdfVisualRenderOutputCharacters)
        $stderrTask = [PaperToJournalClub.BoundedTextReader]::ReadAsync($process.StandardError, $MaximumPdfVisualRenderErrorCharacters)
        $deadline = [DateTime]::UtcNow.AddMilliseconds($PdfVisualRenderTimeoutMilliseconds)
        while (-not $process.HasExited) {
            if ($stdoutTask.IsCompleted -and $stdoutTask.GetAwaiter().GetResult().Exceeded) { try { $process.Kill() } catch { }; throw 'PDF visual renderer exceeded its stdout safety limit.' }
            if ($stderrTask.IsCompleted -and $stderrTask.GetAwaiter().GetResult().Exceeded) { try { $process.Kill() } catch { }; throw 'PDF visual renderer exceeded its stderr safety limit.' }
            if ([DateTime]::UtcNow -ge $deadline) { try { $process.Kill() } catch { }; throw 'PDF visual rendering exceeded the 30-second safety limit.' }
            Start-Sleep -Milliseconds 25
        }
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        if ($stdout.Exceeded -or $stderr.Exceeded) { throw 'PDF visual renderer exceeded an output safety limit.' }
        if ($process.ExitCode -ne 0) { throw "PDF visual renderer failed: $($stderr.Text)" }
        $jsonLine = @($stdout.Text -split "`r?`n" | Where-Object { $_.TrimStart().StartsWith('{') } | Select-Object -Last 1)
        if ($jsonLine.Count -ne 1) { throw 'PDF visual renderer returned no JSON report.' }
        return $jsonLine[0] | ConvertFrom-Json
    } finally {
        if ($process) { $process.Dispose() }
    }
}

function Invoke-RenderPaperVisual {
    param($Arguments)

    $evidencePack = Get-PropertyValue $Arguments 'evidence_pack'
    Assert-McpObject -Value $evidencePack -ParameterName 'evidence_pack'
    $visualId = [string](Get-PropertyValue $Arguments 'visual_id' '')
    $entity = Get-VisualEntityById -EvidencePack $evidencePack -VisualId $visualId
    if ($null -eq $entity) { throw "visual_id must reference a figure or table in evidence_pack: $visualId" }
    # 渲染本身会在临时目录写入文件；因此不能把它当成“看看案例图”的无害操作。
    # 只有已由图注/邻近文本保守认定为方法、主实验或消融的图表才允许进入该路径。
    if (-not (Test-JournalClubVisualEligibility -Visual $entity)) {
        $role = [string](Get-PropertyValue $entity 'journal_club_role' 'excluded-or-unknown')
        throw "visual_id is not eligible for the default journal-club visual workflow (role: $role). Case-analysis, dataset-statistics, and unknown visuals require no rendering or insertion by default."
    }
    if (-not (Test-PropertyExists $Arguments 'page_number')) { throw 'page_number is required.' }
    $requestedPage = Assert-IntegerArgument -Arguments $Arguments -Name 'page_number' -Minimum 1 -Maximum 10000 -Default 1
    $sourcePage = Get-PropertyValue $entity 'source_page'
    $sourcePages = @((Get-PropertyValue $entity 'source_pages' @()) | ForEach-Object {
        try { [int]$_ } catch { $null }
    } | Where-Object { $_ -ge 1 } | Select-Object -Unique)
    if ($null -ne $sourcePage) {
        if ($requestedPage -ne [int]$sourcePage) { throw 'page_number must equal the reviewed source_page for visual_id.' }
    } elseif ($requestedPage -notin $sourcePages) {
        throw "page_number must be one of the reviewed source_pages for ambiguous ${visualId}: $($sourcePages -join ', ')."
    }
    $pageNumber = $requestedPage
    $sourceFile = [string](Get-PropertyValue $evidencePack 'source_file' '')
    $safePdfPath = Get-ValidatedPaperPath $sourceFile
    if ([IO.Path]::GetExtension($safePdfPath).ToLowerInvariant() -ne '.pdf') { throw 'render_paper_visual is available only for PDF evidence packs.' }
    $width = Assert-IntegerArgument -Arguments $Arguments -Name 'width' -Minimum 400 -Maximum 3200 -Default 1600
    $height = Assert-IntegerArgument -Arguments $Arguments -Name 'height' -Minimum 400 -Maximum 3200 -Default 2200
    if (([int64]$width * [int64]$height) -gt $MaximumRenderedVisualPixels) { throw 'width multiplied by height exceeds the rendered-visual pixel budget.' }
    $crop = Get-PropertyValue $Arguments 'crop'
    if ($null -ne $crop) {
        Assert-McpObject -Value $crop -ParameterName 'crop'
        foreach ($name in @('x', 'y', 'width', 'height')) {
            $value = Get-PropertyValue $crop $name
            if ($value -isnot [double] -and $value -isnot [decimal] -and $value -isnot [int] -and $value -isnot [int64]) { throw "crop.$name must be a number." }
        }
        [double]$cropX = Get-PropertyValue $crop 'x'
        [double]$cropY = Get-PropertyValue $crop 'y'
        [double]$cropWidth = Get-PropertyValue $crop 'width'
        [double]$cropHeight = Get-PropertyValue $crop 'height'
        if ($cropX -lt 0 -or $cropY -lt 0 -or $cropWidth -le 0 -or $cropHeight -le 0 -or $cropX + $cropWidth -gt 1 -or $cropY + $cropHeight -gt 1) {
            throw 'crop must be normalized to the reviewed page and remain inside 0..1.'
        }
    }
    $renderDirectory = Get-RenderedVisualDirectory $evidencePack
    $existingRenders = @(Get-ChildItem -LiteralPath $renderDirectory -File -Filter '*.png' -Force)
    if ($existingRenders.Count -ge $MaximumRenderedVisualsPerPaper) { throw "This evidence pack already reached the $MaximumRenderedVisualsPerPaper-rendered-visual limit." }
    [int64]$existingBytes = @($existingRenders | Measure-Object -Property Length -Sum).Sum
    if ($existingBytes -ge $MaximumRenderedVisualBytesPerPaper) { throw 'This evidence pack already reached the rendered-visual byte budget.' }
    $safeStem = ($visualId -replace '[^a-zA-Z0-9._-]', '-')
    $outputPath = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $renderDirectory "$safeStem-page-$pageNumber-$([Guid]::NewGuid().ToString('N')).png") -AllowedRoots @($renderDirectory) -ParameterName 'rendered visual output path'
    $result = Invoke-PdfVisualRenderer -PdfPath $safePdfPath -PageNumber $pageNumber -OutputPath $outputPath -Width $width -Height $height -Crop $crop
    $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath $outputPath -AllowedRoots @($renderDirectory) -ParameterName 'rendered paper visual'
    if ($existingBytes + [int64]$image.bytes -gt $MaximumRenderedVisualBytesPerPaper) {
        Remove-Item -LiteralPath $outputPath -Force
        throw 'Rendered visual would exceed the per-paper byte budget.'
    }
    return [pscustomobject]@{
        visual_id = $visualId
        visual_kind = if ($visualId -like 'table-*') { 'table' } else { 'figure' }
        source_page = $pageNumber
        selection_mode = 'user-confirmed-page-crop'
        crop = $crop
        image_path = $image.path
        image_sha256 = $image.sha256
        image_bytes = $image.bytes
        width = $image.width
        height = $image.height
        renderer = 'windows-data-pdf'
        note = 'Review this rendered candidate before mapping it into a journal-club slide.'
    }
}

function Get-SectionsByTitlePattern {
    param($EvidencePack, [string]$NamePattern)
    return @((Get-PropertyValue $EvidencePack 'sections' @()) | Where-Object {
        (Get-PropertyValue $_ 'title' '') -match $NamePattern
    })
}

function Get-FirstSectionByTitlePattern {
    param($EvidencePack, [string]$NamePattern)
    $sections = @(Get-SectionsByTitlePattern $EvidencePack $NamePattern)
    if ($sections.Count) { return $sections[0] }
    return $null
}

function Get-SectionText {
    param($Section)
    return [string](Get-PropertyValue $Section 'excerpt' '')
}

function ConvertTo-ReadableSlideText {
    param([string]$Text, [int]$ChineseLimit, [int]$OtherLimit)

    $normalized = (($Text -replace '\s+', ' ').Trim())
    if ([string]::IsNullOrWhiteSpace($normalized)) { return '' }
    $limit = Get-ReadableTextLimit -Text $normalized -ChineseLimit $ChineseLimit -OtherLimit $OtherLimit
    if ($normalized.Length -le $limit) { return $normalized }
    # 可读性优先于在单页中塞入整句。完整原文仍保留在 evidence pack，可通过 source id 追溯。
    return ($normalized.Substring(0, [Math]::Max(1, $limit - 1)).TrimEnd() + '…')
}

function Get-VisualDisplayLabel {
    param($Visual, [string]$FallbackId = '')

    $label = ([string](Get-PropertyValue $Visual 'label' '')).Trim()
    if ($label) { return $label }
    return $FallbackId
}

function New-MethodExplanationPoint {
    <#
      系统图说明同时保留两条来源：source_section_id 指向论文 Methods，source_figure_id
      指向图号。这样 PowerPoint 中的口播文字既可回到具体图注/邻近文本，也不会把图像
      像素误当作科学证据。
    #>
    param(
        [string]$Text,
        [string]$Stage,
        [string]$SourceSectionId,
        [string]$SourceFigureId = '',
        [int]$SourcePage = 0,
        [string]$SourceKind = 'methods-text',
        [string]$SourceExcerpt = ''
    )

    return [pscustomobject]@{
        text = $Text
        stage = $Stage
        source_section_id = $SourceSectionId
        source_figure_id = $SourceFigureId
        source_page = $SourcePage
        source_kind = $SourceKind
        source_excerpt = $SourceExcerpt
    }
}

function Get-MethodExplanationEvidence {
    <#
      生成“输入 → 模块 → 输出/验证”顺序的系统图口播点。图注只用于定位图中内容，
      具体科学表述仍优先取自 Methods 原文；当 PDF 提取不到某个阶段时，明确标记为
      方法依据，而不根据截图补写流程细节。
    #>
    param($MethodsSection, $Figure, [string]$Language)

    if ($null -eq $MethodsSection) { return @() }
    $sectionId = [string](Get-PropertyValue $MethodsSection 'id' '')
    if ([string]::IsNullOrWhiteSpace($sectionId)) { return @() }

    $isChinese = Test-ChineseLanguage $Language
    $figureId = [string](Get-PropertyValue $Figure 'id' '')
    $figureLabel = Get-VisualDisplayLabel -Visual $Figure -FallbackId $figureId
    $figurePage = 0
    try { $figurePage = [int](Get-PropertyValue $Figure 'source_page' 0) } catch { }
    $captionContext = ([string](Get-PropertyValue $Figure 'context' '')).Trim()
    $points = @()

    # 图注/邻近文本先交代“这张图为何属于方法页”。它仍显式绑定到 Methods section，
    # 并额外保留 source_figure_id/source_excerpt 供审核器和汇报者追溯。
    if ($figureId -and $captionContext) {
        $captionPrefix = if ($isChinese) { "$figureLabel 的系统图依据是" } else { "$figureLabel system context is" }
        $points += New-MethodExplanationPoint `
            -Text (ConvertTo-ReadableSlideText -Text "$captionPrefix $captionContext" -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters) `
            -Stage 'figure-context' -SourceSectionId $sectionId -SourceFigureId $figureId -SourcePage $figurePage `
            -SourceKind 'figure-caption-or-context' -SourceExcerpt $captionContext
    }

    $stageDefinitions = @(
        [pscustomobject]@{
            id = 'input'
            pattern = '\b(input|data|dataset|sample|instance|query|meme|image|text|feature|observation)\b|输入|数据|数据集|样本|实例|查询|模因|图像|文本|特征|观测'
            zh = '输入与任务设定'
            en = 'Input and task setup'
        },
        [pscustomobject]@{
            id = 'module'
            pattern = '\b(module|agent|component|stage|pipeline|workflow|framework|retrieve|retrieval|derive|derivation|process|processing|encode|encoder|fusion|reasoning)\b|模块|智能体|组件|阶段|流程|框架|检索|推导|处理|编码|融合|推理'
            zh = '关键模块与处理流程'
            en = 'Key modules and processing'
        },
        [pscustomobject]@{
            id = 'output-validation'
            pattern = '\b(output|prediction|inference|classif|result|evaluate|evaluation|validat|metric|benchmark|test)\b|输出|预测|推理|分类|结果|评估|验证|指标|基准|测试'
            zh = '输出与验证方式'
            en = 'Output and validation'
        }
    )
    $usedSentences = @()
    foreach ($stageDefinition in $stageDefinitions) {
        if ($points.Count -ge $MaximumExplanationPointsPerSlide) { break }
        $candidate = @(
            Find-SentenceEvidence @($MethodsSection) $stageDefinition.pattern $MaximumExplanationPointsPerSlide |
                Where-Object { $_.text -and $_.text -notin $usedSentences } |
                Select-Object -First 1
        )
        if ($candidate.Count -eq 0) { continue }
        $sentence = [string]$candidate[0].text
        $usedSentences += $sentence
        $stageLabel = if ($isChinese) { "$($stageDefinition.zh)体现在" } else { "$($stageDefinition.en) is described by" }
        $points += New-MethodExplanationPoint `
            -Text (ConvertTo-ReadableSlideText -Text "$stageLabel $sentence" -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters) `
            -Stage $stageDefinition.id -SourceSectionId $sectionId -SourceFigureId $figureId -SourcePage $figurePage `
            -SourceKind 'methods-text' -SourceExcerpt $sentence
    }

    # 有些论文的 Methods 不直接使用 input/module/output 词汇。为确保方法图至少拥有两条
    # 真实且可追溯的说明，补入未使用的原文句子，但不人为指定不存在的功能模块。
    foreach ($fallback in @(Find-SentenceEvidence @($MethodsSection) '' $MaximumExplanationPointsPerSlide)) {
        if ($points.Count -ge $MaximumExplanationPointsPerSlide) { break }
        $sentence = [string]$fallback.text
        if (-not $sentence -or $sentence -in $usedSentences) { continue }
        $usedSentences += $sentence
        $fallbackLabel = if ($isChinese) { '方法原文进一步说明' } else { 'Methods evidence further states' }
        $points += New-MethodExplanationPoint `
            -Text (ConvertTo-ReadableSlideText -Text "$fallbackLabel $sentence" -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters) `
            -Stage 'methods-evidence' -SourceSectionId $sectionId -SourceFigureId $figureId -SourcePage $figurePage `
            -SourceKind 'methods-text' -SourceExcerpt $sentence
    }
    return @($points | Select-Object -First $MaximumExplanationPointsPerSlide)
}

function Get-ClaimPrimaryPaperEvidence {
    param($Claim)

    $claimId = [string](Get-PropertyValue $Claim 'id' '')
    foreach ($evidence in @(Get-PropertyValue $Claim 'evidence' @())) {
        $excerpt = ([string](Get-PropertyValue $evidence 'excerpt' '')).Trim()
        if (-not $excerpt) { $excerpt = ([string](Get-PropertyValue $Claim 'text' '')).Trim() }
        if (-not $excerpt) { continue }
        return [pscustomobject]@{
            text = $excerpt
            source_section_id = [string](Get-PropertyValue $evidence 'section_id' '')
            source_page = Get-PropertyValue $evidence 'page_number'
            source_claim_id = $claimId
            source_kind = 'claim-evidence'
        }
    }
    $fallback = ([string](Get-PropertyValue $Claim 'text' '')).Trim()
    if (-not $fallback) { return $null }
    return [pscustomobject]@{
        text = $fallback
        source_section_id = ''
        source_page = $null
        source_claim_id = $claimId
        source_kind = 'claim-text-without-excerpt'
    }
}

function Find-SourceBackedSentence {
    param(
        $EvidencePack,
        [string[]]$PreferredSectionIds = @(),
        [string]$Pattern = '',
        [string]$ExcludeText = ''
    )

    $sectionsById = @{}
    foreach ($section in @(Get-PropertyValue $EvidencePack 'sections' @())) {
        $id = [string](Get-PropertyValue $section 'id' '')
        if ($id) { $sectionsById[$id] = $section }
    }
    $orderedSections = @()
    foreach ($id in @($PreferredSectionIds)) {
        if ($id -and $sectionsById.ContainsKey([string]$id) -and $sectionsById[[string]$id] -notin $orderedSections) {
            $orderedSections += $sectionsById[[string]$id]
        }
    }
    foreach ($section in @(Get-PropertyValue $EvidencePack 'sections' @())) {
        $title = [string](Get-PropertyValue $section 'title' '')
        if ($title -match 'result|discussion|conclusion|结果|讨论|结论' -and $section -notin $orderedSections) {
            $orderedSections += $section
        }
    }
    foreach ($section in $orderedSections) {
        foreach ($sentence in @(Split-PaperSentences (Get-SectionText $section))) {
            if (-not $sentence -or ($Pattern -and $sentence -notmatch $Pattern)) { continue }
            if ($ExcludeText -and (($sentence -replace '\s+', ' ').Trim()) -eq (($ExcludeText -replace '\s+', ' ').Trim())) { continue }
            return [pscustomobject]@{
                text = $sentence
                source_section_id = [string](Get-PropertyValue $section 'id' '')
                source_page = Get-PropertyValue $section 'source_page'
                source_kind = 'paper-section-sentence'
            }
        }
    }
    return $null
}

function Find-AuthorExplanationEvidence {
    param($Claim, $EvidencePack, $ComparisonEvidence)

    $preferredSectionIds = @((Get-PropertyValue $Claim 'evidence' @()) | ForEach-Object { [string](Get-PropertyValue $_ 'section_id' '') } | Where-Object { $_ } | Select-Object -Unique)
    # 只摘录作者自己使用的解释性措辞。找不到独立解释时退回同一条已引用的作者结果句，
    # 并保留 source_kind，避免将插件的泛化判断伪装成论文作者的机制解释。
    $pattern = '\b(because|due to|suggest|indicat|demonstrat|show|support|consistent|explain|reflect|imply|attribut|therefore|mechanism)\b|因为|由于|提示|表明|说明|支持|一致|解释|反映|意味着|机制'
    $found = Find-SourceBackedSentence -EvidencePack $EvidencePack -PreferredSectionIds $preferredSectionIds -Pattern $pattern -ExcludeText ([string](Get-PropertyValue $ComparisonEvidence 'text' ''))
    if ($null -ne $found) {
        $found | Add-Member -NotePropertyName 'source_kind' -NotePropertyValue 'author-explanation-sentence' -Force
        return $found
    }
    if ($null -ne $ComparisonEvidence) {
        return [pscustomobject]@{
            text = [string]$ComparisonEvidence.text
            source_section_id = [string]$ComparisonEvidence.source_section_id
            source_page = $ComparisonEvidence.source_page
            source_kind = 'author-result-statement-no-separate-explanation'
        }
    }
    return $null
}

function Find-ResultCaveatEvidence {
    param($Claim, $CounterEvidence, $EvidencePack)

    # Windows PowerShell 会把管道筛选结果保留为单元素数组；统一取第一个对象，避免
    # 反例文本在五分钟汇报的单页结果中被静默丢失。
    $counterItem = @($CounterEvidence | Where-Object { $null -ne $_ } | Select-Object -First 1)
    if ($counterItem.Count -eq 1) {
        $counter = Get-ClaimPrimaryPaperEvidence -Claim $counterItem[0]
        if ($null -ne $counter) {
            $counter | Add-Member -NotePropertyName 'source_kind' -NotePropertyValue 'reported-counter-claim' -Force
            return $counter
        }
    }
    # 合并五分钟结果页时，选择器会把已确认的反例文本挂回主 claim。若对象在跨进程 JSON
    # 还原中变成标量，仍优先保留这段原文，并带回原始 claim id 供审核器核验。
    $counterText = ([string](Get-PropertyValue $Claim 'counter_evidence_text' '')).Trim()
    $counterClaimId = ([string](Get-PropertyValue $Claim 'counter_evidence_claim_id' '')).Trim()
    # 调用方有时会在构造过程中直接修改 selected claim，使两个字段尚未随对象一起带入。
    # 只要已显式传入 CounterEvidence，仍可保留其原文和 id；这不是从图像猜测的回退。
    if (-not $counterText -and $counterItem.Count -eq 1) {
        $counterText = ([string](Get-PropertyValue $counterItem[0] 'text' '')).Trim()
        $counterClaimId = ([string](Get-PropertyValue $counterItem[0] 'id' '')).Trim()
    }
    if ($counterText -and $counterClaimId) {
        return [pscustomobject]@{
            text = $counterText
            source_section_id = ''
            source_page = $null
            source_claim_id = $counterClaimId
            source_kind = 'reported-counter-claim'
        }
    }
    $preferredSectionIds = @((Get-PropertyValue $Claim 'evidence' @()) | ForEach-Object { [string](Get-PropertyValue $_ 'section_id' '') } | Where-Object { $_ } | Select-Object -Unique)
    $pattern = '\b(limit(?:ation|ed|s)?|constrain(?:t|ed|s)?|caveat|lack(?:s|ed)?|small sample|few (?:donors|patients|samples)|in[ -]?vitro|generaliz|may not|cannot|unable|unclear|bias|inferior|worse|failed|no\s+improvement|no\s+(?:statistically\s+)?significant)\b|局限|限制|样本量小|体外|外推|泛化|偏倚|尚不清楚|不能|较差|劣于|失败|未改善|无显著'
    $found = Find-SourceBackedSentence -EvidencePack $EvidencePack -PreferredSectionIds $preferredSectionIds -Pattern $pattern
    if ($null -ne $found) {
        $found | Add-Member -NotePropertyName 'source_kind' -NotePropertyValue 'reported-limitation-or-counterexample' -Force
    }
    return $found
}

function Test-TraceablePaperExcerpt {
    <#
      deck spec 中的短摘录可能因版式限制被省略号截断。审核时只接受它是 evidence pack
      原文的前缀或完整句，不能接受与原文毫无关系的“读图得到的数值”。
    #>
    param([string]$Excerpt, [string[]]$SourceTexts = @())

    $normalizedExcerpt = (($Excerpt -replace '[…]+$', '') -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($normalizedExcerpt)) { return $false }
    foreach ($sourceText in @($SourceTexts)) {
        $normalizedSource = (([string]$sourceText -replace '\s+', ' ').Trim())
        if (-not $normalizedSource) { continue }
        if ($normalizedSource.StartsWith($normalizedExcerpt, [StringComparison]::OrdinalIgnoreCase) -or
            $normalizedExcerpt.StartsWith($normalizedSource, [StringComparison]::OrdinalIgnoreCase) -or
            $normalizedSource.IndexOf($normalizedExcerpt, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }
    return $false
}

function New-ResultAnalysis {
    <#
      结果/消融图表的三行口播分别是：原文比较、作者解释、限制或反例。所有可用的
      图表编号、正文摘录和章节 id 都保存在结构化字段中；生成器只渲染三行文字，绝不
      从 PNG 截图读取数值或推导优劣结论。
    #>
    param($Claim, $Visual, $EvidencePack, $CounterEvidence, [string]$Language)

    $kind = [string](Get-PropertyValue $Claim 'comparison_kind' '')
    if (-not $kind) { $kind = Get-ClaimComparisonKind ([string](Get-PropertyValue $Claim 'text' '')) }
    $isChinese = Test-ChineseLanguage $Language
    $visualId = [string](Get-PropertyValue $Visual 'id' '')
    $visualLabel = Get-VisualDisplayLabel -Visual $Visual -FallbackId $visualId
    $visualRole = [string](Get-PropertyValue $Visual 'journal_club_role' '')
    $comparisonEvidence = Get-ClaimPrimaryPaperEvidence -Claim $Claim
    $comparisonText = if ($comparisonEvidence) { [string]$comparisonEvidence.text } else { [string](Get-PropertyValue $Claim 'text' '') }
    $authorEvidence = Find-AuthorExplanationEvidence -Claim $Claim -EvidencePack $EvidencePack -ComparisonEvidence $comparisonEvidence
    $caveatEvidence = Find-ResultCaveatEvidence -Claim $Claim -CounterEvidence $CounterEvidence -EvidencePack $EvidencePack

    $authorText = if ($authorEvidence) { [string]$authorEvidence.text } elseif ($comparisonText) { $comparisonText } else { '' }
    # `Find-SourceBackedSentence` 的 section excerpt 有时只保留前 1600 个字符；它能找到
    # 后半段的完整句，但审核器只能检查 evidence pack 中已暴露的 excerpt。因此统一把
    # 展示和可追溯字段截成两端都能验证的前缀，完整原文仍保留在解析器输出/论文文件中。
    $authorExcerpt = ConvertTo-ReadableSlideText -Text $authorText -ChineseLimit 180 -OtherLimit 360
    if ($authorEvidence) {
        $authorSectionText = ''
        foreach ($section in @(Get-PropertyValue $EvidencePack 'sections' @())) {
            if ([string](Get-PropertyValue $section 'id' '') -eq [string](Get-PropertyValue $authorEvidence 'source_section_id' '')) {
                $authorSectionText = [string](Get-PropertyValue $section 'excerpt' '')
                break
            }
        }
        if ($authorSectionText -and -not (Test-TraceablePaperExcerpt -Excerpt $authorExcerpt -SourceTexts @($authorSectionText))) {
            # 不把一条审计无法回链的后半段结果句包装成“作者解释”。退回已验证的比较句，
            # 并在输出文案中如实标注为结果原文陈述。
            $authorEvidence = [pscustomobject]@{
                text = $comparisonText
                source_section_id = if ($comparisonEvidence) { [string]$comparisonEvidence.source_section_id } else { '' }
                source_page = if ($comparisonEvidence) { $comparisonEvidence.source_page } else { $null }
                source_kind = 'author-result-statement-no-separate-explanation'
            }
            $authorText = [string]$authorEvidence.text
            $authorExcerpt = ConvertTo-ReadableSlideText -Text $authorText -ChineseLimit 180 -OtherLimit 360
        }
    }
    # 没有独立解释句时，只能说“结果原文指出”，不能把同一条比较结果冒充成作者的
    # 机制解释。这个来源类型也会继续输出，便于审计和汇报者区分两种情况。
    $authorEvidenceKind = if ($authorEvidence) { [string](Get-PropertyValue $authorEvidence 'source_kind' '') } else { 'missing' }
    $hasIndependentAuthorExplanation = $authorEvidenceKind -eq 'author-explanation-sentence'
    if ($isChinese) {
        $comparisonPrefix = if ($visualLabel) { "$visualLabel 原文比较" } else { '论文原文比较' }
        $interpretationPrefix = if ($hasIndependentAuthorExplanation) {
            if ($visualLabel) { "$visualLabel 作者解释" } else { '论文作者解释' }
        } else {
            if ($visualLabel) { "$visualLabel 结果原文指出" } else { '结果原文指出' }
        }
        $caveatPrefix = if ($visualLabel) { "$visualLabel 限制或反例" } else { '限制或反例' }
        $comparisonDisplayText = "$comparisonPrefix 显示 $comparisonText"
        $interpretationDisplayText = "$interpretationPrefix $authorText"
        $caveatJoinText = '是'
    } else {
        $comparisonPrefix = if ($visualLabel) { "$visualLabel reported comparison" } else { 'Reported comparison' }
        $interpretationPrefix = if ($hasIndependentAuthorExplanation) {
            if ($visualLabel) { "$visualLabel author explanation" } else { 'Author explanation' }
        } else {
            if ($visualLabel) { "$visualLabel result text" } else { 'Result text' }
        }
        $caveatPrefix = if ($visualLabel) { "$visualLabel limitation or counterevidence" } else { 'Limitation or counterevidence' }
        $comparisonDisplayText = "$comparisonPrefix shows $comparisonText"
        $interpretationDisplayText = "$interpretationPrefix states $authorText"
        $caveatJoinText = 'is'
    }
    $caveatKind = 'evidence-boundary'
    if ($caveatEvidence) {
        $caveatKind = [string](Get-PropertyValue $caveatEvidence 'source_kind' 'reported-limitation-or-counterexample')
        $caveatText = [string]$caveatEvidence.text
    } else {
        # 没有可回链的反例不能据此写“论文没有局限”。这里只陈述展示边界，提醒汇报者
        # 不要看着截图补写数字、显著性或因果结论。
        $caveatText = if ($isChinese) {
            '当前可回链文本未定位到该图表的反例；不要根据截图推断未报告数值或结论。'
        } else {
            'No figure-specific counterexample was located in the traceable text; do not infer unreported values or conclusions from the screenshot.'
        }
    }
    $comparisonDisplay = ConvertTo-ReadableSlideText -Text $comparisonDisplayText -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters
    $interpretationDisplay = ConvertTo-ReadableSlideText -Text $interpretationDisplayText -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters
    $caveatDisplay = ConvertTo-ReadableSlideText -Text "$caveatPrefix $caveatJoinText $caveatText" -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters

    return [pscustomobject]@{
        comparison_kind = $kind
        visual_id = if ($visualId) { $visualId } else { $null }
        visual_label = if ($visualLabel) { $visualLabel } else { $null }
        visual_role = if ($visualRole) { $visualRole } else { $null }
        visual_caption_excerpt = ConvertTo-ReadableSlideText -Text ([string](Get-PropertyValue $Visual 'context' '')) -ChineseLimit 180 -OtherLimit 360
        comparison = $comparisonDisplay
        comparison_source_claim_id = if ($comparisonEvidence) { [string]$comparisonEvidence.source_claim_id } else { [string](Get-PropertyValue $Claim 'id' '') }
        comparison_source_section_id = if ($comparisonEvidence) { [string]$comparisonEvidence.source_section_id } else { '' }
        comparison_source_excerpt = ConvertTo-ReadableSlideText -Text $comparisonText -ChineseLimit 180 -OtherLimit 360
        interpretation = $interpretationDisplay
        author_explanation_excerpt = $authorExcerpt
        author_explanation_source_section_id = if ($authorEvidence) { [string]$authorEvidence.source_section_id } else { '' }
        author_explanation_source_kind = $authorEvidenceKind
        caveat = $caveatDisplay
        caveat_kind = $caveatKind
        caveat_source_claim_id = if ($caveatEvidence) { [string](Get-PropertyValue $caveatEvidence 'source_claim_id' '') } else { $null }
        caveat_source_section_id = if ($caveatEvidence) { [string](Get-PropertyValue $caveatEvidence 'source_section_id' '') } else { $null }
        caveat_source_excerpt = if ($caveatEvidence) { ConvertTo-ReadableSlideText -Text ([string]$caveatEvidence.text) -ChineseLimit 180 -OtherLimit 360 } else { $null }
    }
}

function Get-ResultSlideTitle {
    param($Claim, [string]$Language)

    $kind = [string](Get-PropertyValue $Claim 'comparison_kind' '')
    if (-not $kind) { $kind = Get-ClaimComparisonKind ([string](Get-PropertyValue $Claim 'text' '')) }
    if (Test-ChineseLanguage $Language) {
        switch ($kind) {
            'reported-better' { return '主要比较显示更优表现' }
            'reported-worse-or-no-benefit' { return '主要比较未显示预期改善' }
            'no-clear-difference' { return '主要比较未显示明确差异' }
            default { return '论文报告的核心实验结果' }
        }
    }
    switch ($kind) {
        'reported-better' { return 'Primary comparison reports better performance' }
        'reported-worse-or-no-benefit' { return 'Primary comparison reports no expected improvement' }
        'no-clear-difference' { return 'Primary comparison reports no clear difference' }
        default { return 'Core experimental result reported by the paper' }
    }
}

function Select-NarrativeResultClaims {
    param($Claims, [int]$Maximum)

    $rank = @{ 'reported-better' = 0; 'reported-worse-or-no-benefit' = 1; 'no-clear-difference' = 2; 'reported-observation' = 3 }
    $ordered = @($Claims | Sort-Object @{ Expression = {
        $kind = [string](Get-PropertyValue $_ 'comparison_kind' '')
        if (-not $rank.ContainsKey($kind)) { return 9 }
        return $rank[$kind]
    } })
    $selected = @($ordered | Select-Object -First $Maximum)
    # 若论文同时给出正面与不利/无差异证据，尽量让短汇报也呈现边界，而非只挑好看的结果。
    $counterEvidence = @($ordered | Where-Object { (Get-PropertyValue $_ 'comparison_kind' '') -in @('reported-worse-or-no-benefit', 'no-clear-difference') } | Select-Object -First 1)
    if ($Maximum -eq 1 -and $counterEvidence.Count -and $selected.Count -and (Get-PropertyValue $selected[0] 'comparison_kind' '') -eq 'reported-better') {
        # 五分钟汇报只有一页结果时，不能静默丢弃论文已报告的不利/无差异证据。将它合并为
        # 同一 claim 的“需呈现边界”标记，设计阶段仍只复用真实来源 id，不虚构第二张图。
        $selected[0] | Add-Member -NotePropertyName 'counter_evidence_claim_id' -NotePropertyValue ([string](Get-PropertyValue $counterEvidence[0] 'id' '')) -Force
        $selected[0] | Add-Member -NotePropertyName 'counter_evidence_text' -NotePropertyValue ([string](Get-PropertyValue $counterEvidence[0] 'text' '')) -Force
        return @($selected)
    }
    if ($Maximum -ge 2 -and $counterEvidence.Count -and @($selected | Where-Object { (Get-PropertyValue $_ 'id' '') -eq (Get-PropertyValue $counterEvidence[0] 'id' '') }).Count -eq 0) {
        if ($selected.Count -ge $Maximum) { $selected[$selected.Count - 1] = $counterEvidence[0] } else { $selected += $counterEvidence[0] }
    }
    return @($selected)
}

function Find-SentenceEvidence {
    param($Sections, [string]$Pattern = '', [int]$Maximum = 3)

    # 只返回论文中能逐句回链的文本；不会为“创新点”或“未来方向”补写未经证实的事实。
    # 不使用 PowerShell 的自动变量 $Matches，避免正则匹配后将其重置为 Hashtable。
    $evidenceItems = @()
    foreach ($section in @($Sections)) {
        if ($null -eq $section) { continue }
        $sentences = @(Split-PaperSentences (Get-SectionText $section))
        if ($Pattern) { $sentences = @($sentences | Where-Object { $_ -match $Pattern }) }
        foreach ($sentence in $sentences) {
            $evidenceItems += [pscustomobject]@{ section = $section; text = $sentence }
            if ($evidenceItems.Count -ge $Maximum) { return @($evidenceItems) }
        }
    }
    return @($evidenceItems)
}

function Get-EvidenceItemSectionIds {
    param($EvidenceItems)
    $ids = @()
    foreach ($item in @($EvidenceItems)) {
        $sectionId = Get-PropertyValue (Get-PropertyValue $item 'section') 'id'
        if ($sectionId -and $sectionId -notin $ids) { $ids += [string]$sectionId }
    }
    return @($ids)
}

function Get-ClaimEvidenceSectionIds {
    param($Claim)
    $ids = @()
    foreach ($evidence in @(Get-PropertyValue $Claim 'evidence' @())) {
        $sectionId = Get-PropertyValue $evidence 'section_id'
        if ($sectionId -and $sectionId -notin $ids) { $ids += [string]$sectionId }
    }
    return @($ids)
}

function Test-FigureIdMention {
    param([string]$Text, [string]$FigureId)

    if ([string]::IsNullOrWhiteSpace($Text) -or [string]::IsNullOrWhiteSpace($FigureId)) { return $false }
    $idMatch = [regex]::Match($FigureId, '^fig-(\d+)([a-zA-Z]?)$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $idMatch.Success) { return $false }

    $number = [regex]::Escape($idMatch.Groups[1].Value)
    $suffix = [regex]::Escape($idMatch.Groups[2].Value)
    # 数字和可选 panel 后都必须是边界：Fig. 1 不能命中 Fig. 10 / Fig. 1A；
    # Fig. 1A 也不能命中 Fig. 1AB。统一从完整“Fig/Figure/图 + 编号”提取，
    # 不再对 label 做普通子串匹配。
    $pattern = "(?i)(?:\b(?:fig(?:ure)?s?\.?)\s*|图\s*)$number$suffix(?![A-Za-z0-9])"
    return [regex]::IsMatch($Text, $pattern)
}

function Get-MentionedFigureIds {
    param([string]$Text, $Figures)

    $mentionedIds = @()
    foreach ($figure in @($Figures)) {
        $figureId = [string](Get-PropertyValue $figure 'id' '')
        if ($figureId -and (Test-FigureIdMention -Text $Text -FigureId $figureId) -and $figureId -notin $mentionedIds) {
            $mentionedIds += $figureId
        }
    }
    return @($mentionedIds)
}

function Test-TableIdMention {
    param([string]$Text, [string]$TableId)

    if ([string]::IsNullOrWhiteSpace($Text) -or [string]::IsNullOrWhiteSpace($TableId)) { return $false }
    $idMatch = [regex]::Match($TableId, '^table-(\d+)([a-zA-Z]?)$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $idMatch.Success) { return $false }
    $number = [regex]::Escape($idMatch.Groups[1].Value)
    $suffix = [regex]::Escape($idMatch.Groups[2].Value)
    $pattern = "(?i)(?:\b(?:table|tab\.)\s*|表\s*)$number$suffix(?![A-Za-z0-9])"
    return [regex]::IsMatch($Text, $pattern)
}

function Find-TableIdForClaim {
    param($Claim, $Tables)

    $claimText = [string](Get-PropertyValue $Claim 'text' '')
    $mentionedIds = @($Tables | Where-Object { Test-TableIdMention -Text $claimText -TableId ([string](Get-PropertyValue $_ 'id' '')) } | ForEach-Object { [string](Get-PropertyValue $_ 'id' '') })
    if ($mentionedIds.Count -eq 1) { return $mentionedIds[0] }
    if ($mentionedIds.Count -gt 1) { return $null }
    $claimPages = @((Get-PropertyValue $Claim 'evidence' @()) | ForEach-Object { Get-PropertyValue $_ 'page_number' } | Where-Object { $null -ne $_ })
    foreach ($claimPage in $claimPages) {
        $pageTables = @($Tables | Where-Object { (Get-PropertyValue $_ 'source_page') -eq $claimPage })
        if ($pageTables.Count -eq 1) { return [string](Get-PropertyValue $pageTables[0] 'id') }
    }
    return $null
}

function Find-FigureIdForClaim {
    param($Claim, $Figures)
    $claimText = [string](Get-PropertyValue $Claim 'text' '')
    $mentionedIds = @(Get-MentionedFigureIds -Text $claimText -Figures $Figures)
    # 一个结论明确提及多个图时不擅自选择第一个；让用户用 figure_asset_selection
    # 明确决定要呈现哪张图或哪个 panel。
    if ($mentionedIds.Count -eq 1) { return $mentionedIds[0] }
    if ($mentionedIds.Count -gt 1) { return $null }
    # 只有当该结论所在页恰好只有一个已识别图号时，才采用页码作为保守辅助匹配；多图页绝不猜测。
    $claimPages = @((Get-PropertyValue $Claim 'evidence' @()) | ForEach-Object { Get-PropertyValue $_ 'page_number' } | Where-Object { $null -ne $_ })
    foreach ($claimPage in $claimPages) {
        $pageFigures = @($Figures | Where-Object { (Get-PropertyValue $_ 'source_page') -eq $claimPage })
        if ($pageFigures.Count -eq 1) { return Get-PropertyValue $pageFigures[0] 'id' }
    }
    return $null
}

function Find-FigureIdForSection {
    param($Section, $Figures)

    if ($null -eq $Section) { return $null }
    $sectionText = Get-SectionText $Section
    $mentionedIds = @(Get-MentionedFigureIds -Text $sectionText -Figures $Figures)
    if ($mentionedIds.Count -eq 1) { return $mentionedIds[0] }
    if ($mentionedIds.Count -gt 1) { return $null }

    # 方法章节的文字没有直接写出图号时，只在该章节定位页恰好只有一个图号的情况下
    # 才使用同页图作为系统结构/流程图候选；多图页不进行猜测。
    $sectionPage = Get-PropertyValue $Section 'source_page'
    if ($null -ne $sectionPage) {
        $pageFigures = @($Figures | Where-Object { (Get-PropertyValue $_ 'source_page') -eq $sectionPage })
        if ($pageFigures.Count -eq 1) { return Get-PropertyValue $pageFigures[0] 'id' }
    }
    return $null
}

function Get-FigureById {
    param($Figures, [string]$FigureId)
    if (-not $FigureId) { return $null }
    foreach ($figure in @($Figures)) {
        if ((Get-PropertyValue $figure 'id') -eq $FigureId) { return $figure }
    }
    return $null
}

function Find-JournalClubVisualIdForSection {
    <#
      方法页只能选择 methods-system。即使 Methods 段落提到实验曲线或案例图，也不能
      因为图号出现过就借用为系统结构图；多张合格方法图同样不擅自挑选第一张。
    #>
    param($Section, $Figures)

    if ($null -eq $Section) { return $null }
    $sectionText = Get-SectionText $Section
    $mentionedIds = @(Get-MentionedFigureIds -Text $sectionText -Figures $Figures)
    $eligible = @($Figures | Where-Object {
        (Get-PropertyValue $_ 'id' '') -in $mentionedIds -and
        (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles @('methods-system'))
    })
    if ($eligible.Count -eq 1) { return [string](Get-PropertyValue $eligible[0] 'id' '') }
    if ($eligible.Count -gt 1) { return $null }

    # 没有明确图号时，只允许“章节所在页恰有一张合格系统图”的辅助匹配。
    $sectionPage = Get-PropertyValue $Section 'source_page'
    if ($null -ne $sectionPage) {
        $pageEligible = @($Figures | Where-Object {
            (Get-PropertyValue $_ 'source_page') -eq $sectionPage -and
            (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles @('methods-system'))
        })
        if ($pageEligible.Count -eq 1) { return [string](Get-PropertyValue $pageEligible[0] 'id' '') }
    }
    return $null
}

function Find-JournalClubVisualIdForClaim {
    <#
      实验数据页接受 main-result 或 ablation。候选选择仍以 claim 中明确图表编号为先，
      只有同页恰好一张合格图表时才辅助匹配，避免把案例页“顺手”带进结果页。
    #>
    param($Claim, $Figures, $Tables)

    $claimText = [string](Get-PropertyValue $Claim 'text' '')
    $allowedRoles = @('main-result', 'ablation')
    $mentionedFigures = @($Figures | Where-Object {
        (Test-FigureIdMention -Text $claimText -FigureId ([string](Get-PropertyValue $_ 'id' ''))) -and
        (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles $allowedRoles)
    })
    $mentionedTables = @($Tables | Where-Object {
        (Test-TableIdMention -Text $claimText -TableId ([string](Get-PropertyValue $_ 'id' ''))) -and
        (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles $allowedRoles)
    })
    # 只要 claim 提及多个 Figure/Table，即使其中只有一张可用，也不把剩余一张当成
    # 默认答案。作者可能是在做并列比较，必须由用户显式决定展示哪个 panel 或表格。
    $allMentionedFigureIds = @(Get-MentionedFigureIds -Text $claimText -Figures $Figures)
    $allMentionedTableIds = @($Tables | Where-Object { Test-TableIdMention -Text $claimText -TableId ([string](Get-PropertyValue $_ 'id' '')) } | ForEach-Object { [string](Get-PropertyValue $_ 'id' '') })
    if (($allMentionedFigureIds.Count + $allMentionedTableIds.Count) -gt 1) { return $null }
    $mentioned = @($mentionedFigures) + @($mentionedTables)
    if ($mentioned.Count -eq 1) { return [string](Get-PropertyValue $mentioned[0] 'id' '') }
    if ($mentioned.Count -gt 1) { return $null }

    $claimPages = @((Get-PropertyValue $Claim 'evidence' @()) | ForEach-Object { Get-PropertyValue $_ 'page_number' } | Where-Object { $null -ne $_ } | Select-Object -Unique)
    foreach ($claimPage in $claimPages) {
        $pageEligible = @(
            $Figures | Where-Object {
                (Get-PropertyValue $_ 'source_page') -eq $claimPage -and
                (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles $allowedRoles)
            }
        ) + @(
            $Tables | Where-Object {
                (Get-PropertyValue $_ 'source_page') -eq $claimPage -and
                (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles $allowedRoles)
            }
        )
        if ($pageEligible.Count -eq 1) { return [string](Get-PropertyValue $pageEligible[0] 'id' '') }
    }
    return $null
}

function Find-PreferredResultVisualId {
    <#
      自动抽取的 claim 有时落在引言、方法或跨页段落，不能仅凭 claim 的局部句子找到
      主实验表。仅在 Results 章节中没有一个合格图表已被选中时，才以“唯一可用且有
      页码的主实验/消融图表”作为保守回退；有多个候选时宁可不插图。
    #>
    param($Claims, $Figures, $Tables)

    $eligibleVisuals = @(
        $Figures | Where-Object { (Get-PropertyValue $_ 'source_page') -and (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles @('main-result', 'ablation')) }
    ) + @(
        $Tables | Where-Object { (Get-PropertyValue $_ 'source_page') -and (Test-JournalClubVisualEligibility -Visual $_ -AllowedRoles @('main-result', 'ablation')) }
    )
    $eligibleVisuals = @($eligibleVisuals | Sort-Object @{ Expression = {
        # 主实验优先于消融；相同角色再按论文顺序，保持摘要→主结果→机制/消融的叙事。
        if ((Get-PropertyValue $_ 'journal_club_role' '') -eq 'main-result') { 0 } else { 1 }
    } }, @{ Expression = { [int](Get-PropertyValue $_ 'source_page' 0) } })
    # 这个回退只在整篇论文中恰好只有一个合格主实验/消融视觉时才成立。
    # 多个候选意味着“结果句 → 图表”关系不明确，不能为了填满第一页结果页而任取
    # 排在前面的表或图；此时保留文字分析，并由用户确认具体 Figure/Table。
    if ($eligibleVisuals.Count -ne 1) { return $null }

    # 只有在全部 result claims 都没有可靠可视化对应时才回退，避免在已有明确 Fig/Table
    # 对应的论文中重复插入同一张图表。
    foreach ($claim in @($Claims)) {
        $claimText = [string](Get-PropertyValue $claim 'text' '')
        $allMentionedFigureIds = @(Get-MentionedFigureIds -Text $claimText -Figures $Figures)
        $allMentionedTableIds = @($Tables | Where-Object { Test-TableIdMention -Text $claimText -TableId ([string](Get-PropertyValue $_ 'id' '')) } | ForEach-Object { [string](Get-PropertyValue $_ 'id' '') })
        # Claim 明确说“Fig. 1 and Fig. 10 …”时，绝不能用回退规则擅自挑主实验图；这和
        # 直接匹配到两个可用图表一样都需要用户选择。
        if (($allMentionedFigureIds.Count + $allMentionedTableIds.Count) -gt 1) { return $null }
        if (Find-JournalClubVisualIdForClaim -Claim $claim -Figures $Figures -Tables $tables) { return $null }
    }
    return [string](Get-PropertyValue $eligibleVisuals[0] 'id' '')
}

function Get-VisualById {
    param($Figures, $Tables, [string]$VisualId)
    $figure = Get-FigureById -Figures $Figures -FigureId $VisualId
    if ($null -ne $figure) { return $figure }
    foreach ($table in @($Tables)) {
        if ([string](Get-PropertyValue $table 'id' '') -eq $VisualId) { return $table }
    }
    return $null
}

function Get-SafeEvidenceAssetDirectory {
    param($EvidencePack)

    $directory = Get-PropertyValue (Get-PropertyValue $EvidencePack 'extraction') 'asset_directory'
    if ([string]::IsNullOrWhiteSpace($directory)) { return $null }
    return Assert-PaperToJournalClubAllowedPath -Path $directory -AllowedRoots @((Get-PaperToJournalClubTemporaryRoot)) -ParameterName 'evidence_pack.extraction.asset_directory'
}

function Get-FigureAssetCandidates {
    param($Figure, [string]$AssetDirectory)
    if ($null -eq $Figure) { return @() }

    # 只透传解析器或用户已提供的图像资产路径，绝不根据图号猜测磁盘文件。
    $candidates = @()
    foreach ($propertyName in @('figure_asset_candidates', 'asset_candidates', 'image_candidates')) {
        foreach ($value in @(Get-PropertyValue $Figure $propertyName @())) {
            foreach ($candidate in @(ConvertTo-NonEmptyStringArray $value)) {
                if ([string]::IsNullOrWhiteSpace($AssetDirectory)) {
                    throw 'Figure asset candidates require an extracted temporary asset directory from this plugin.'
                }
                $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath $candidate -AllowedRoots @($AssetDirectory) -ParameterName 'figure asset candidate'
                if ($image.path -notin $candidates) { $candidates += $image.path }
            }
        }
    }
    foreach ($propertyName in @('suggested_image_path', 'image_path', 'asset_path')) {
        $path = Get-PropertyValue $Figure $propertyName
        if ($path) {
            if ([string]::IsNullOrWhiteSpace($AssetDirectory)) {
                throw 'Figure asset candidates require an extracted temporary asset directory from this plugin.'
            }
            $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath ([string]$path) -AllowedRoots @($AssetDirectory) -ParameterName 'figure asset candidate'
            if ($image.path -notin $candidates) { $candidates += $image.path }
        }
    }
    return @($candidates)
}

function Get-ExplicitFigureAssetPath {
    param($Selections, [string]$FigureId, [string]$AssetDirectory)
    if ($null -eq $Selections -or -not $FigureId) { return $null }
    $candidate = Get-PropertyValue $Selections $FigureId
    if (-not $candidate) { return $null }
    if ([string]::IsNullOrWhiteSpace($AssetDirectory)) {
        throw "Selected figure asset for $FigureId is only allowed when it was extracted into this plugin's temporary asset directory."
    }
    $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath ([string]$candidate) -AllowedRoots @($AssetDirectory) -ParameterName "selected figure asset for $FigureId"
    return $image.path
}

function Resolve-FigureImageBinding {
    param($Figure, $Selections, [string]$AssetDirectory)

    $figureId = [string](Get-PropertyValue $Figure 'id' '')
    # 即使调用方给出了经确认的图片路径，案例图和未分类图也不能被静默插入。
    # 此处重算角色，而不是信任 evidence pack 中可能被修改的布尔字段。
    $classification = Get-JournalClubVisualClassification -VisualKind 'figure' -CaptionOrContext ([string](Get-PropertyValue $Figure 'context' ''))
    if (-not $classification.eligible_for_journal_club_visual) {
        return [pscustomobject]@{
            figure_asset_candidates = @()
            suggested_image_path = $null
            selection_mode = 'excluded-by-journal-club-role'
            source_asset_id = $null
        }
    }
    $candidates = @(Get-FigureAssetCandidates -Figure $Figure -AssetDirectory $AssetDirectory)
    $explicitImagePath = Get-ExplicitFigureAssetPath -Selections $Selections -FigureId $figureId -AssetDirectory $AssetDirectory
    if ($explicitImagePath -and $explicitImagePath -notin $candidates) {
        $candidates = @($explicitImagePath) + $candidates
    }

    $automaticImagePath = Get-PropertyValue $Figure 'automatic_image_path'
    $automaticBinding = Get-PropertyValue $Figure 'automatic_binding'
    $automaticAssetId = $null
    # evidence_pack 也可能由 Agent/用户编辑，不能只信任 automatic_image_path 这个字段。
    # 自动插图必须在设计阶段再次满足“一图号、一张合规候选图”的完整条件。
    $automaticMappingAllowed = ([string](Get-PropertyValue $Figure 'asset_match' '')) -eq 'same-page-single-raster' -and $candidates.Count -eq 1
    $safeAutomaticImagePath = $null
    if ($automaticMappingAllowed -and $automaticImagePath -and -not [string]::IsNullOrWhiteSpace($AssetDirectory)) {
        try {
            $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath ([string]$automaticImagePath) -AllowedRoots @($AssetDirectory) -ParameterName "automatic figure asset for $figureId"
            # automatic_image_path 只有同时出现在候选清单中才可信，避免手工构造 evidence pack
            # 把任意临时目录文件伪装成自动匹配的论文图片。
            if ($image.path -in $candidates) {
                $safeAutomaticImagePath = $image.path
                # 新 evidence pack 会把“Fig 编号 -> parser asset id”作为显式、可审计映射返回。
                # 对手工/旧包，只有完整字段、编号和路径都与当前候选一致时才携带 asset id；
                # 缺失映射不阻断兼容的自动插图，但绝不杜撰 asset id。
                if ($automaticBinding -is [System.Management.Automation.PSCustomObject] -or $automaticBinding -is [Collections.IDictionary]) {
                    $bindingFigureId = [string](Get-PropertyValue $automaticBinding 'figure_id' '')
                    $bindingPath = [string](Get-PropertyValue $automaticBinding 'path' '')
                    $bindingAssetId = [string](Get-PropertyValue $automaticBinding 'asset_id' '')
                    $bindingSourcePage = Get-PropertyValue $automaticBinding 'source_page'
                    $bindingMode = [string](Get-PropertyValue $automaticBinding 'mode' '')
                    $sourcePage = Get-PropertyValue $Figure 'source_page'
                    if ($bindingFigureId -eq $figureId -and $bindingPath -eq $image.path -and
                        -not [string]::IsNullOrWhiteSpace($bindingAssetId) -and $bindingMode -eq 'automatic-figure-number-single-raster' -and
                        $null -ne $bindingSourcePage -and $null -ne $sourcePage -and [int]$bindingSourcePage -eq [int]$sourcePage) {
                        $automaticAssetId = $bindingAssetId
                    }
                }
            }
        } catch { }
    }

    return [pscustomobject]@{
        figure_asset_candidates = @($candidates)
        suggested_image_path = if ($explicitImagePath) { $explicitImagePath } else { $safeAutomaticImagePath }
        selection_mode = if ($explicitImagePath) { 'user-confirmed' } elseif ($safeAutomaticImagePath) { 'automatic-same-page-single-raster' } else { 'none' }
        source_asset_id = if ($explicitImagePath) { $null } else { $automaticAssetId }
    }
}

function Resolve-TableImageBinding {
    param($Table, $Selections, [string]$AssetDirectory)

    $tableId = [string](Get-PropertyValue $Table 'id' '')
    if (-not $tableId) { return [pscustomobject]@{ figure_asset_candidates = @(); suggested_image_path = $null; selection_mode = 'none'; source_asset_id = $null } }
    # 统计表、案例相关表和无法识别用途的表默认不进入图片工作流；它们仍可在 evidence
    # pack 中被引用，避免因为“存在 Table 编号”就被错误裁剪进主实验页。
    $classification = Get-JournalClubVisualClassification -VisualKind 'table' -CaptionOrContext ([string](Get-PropertyValue $Table 'context' ''))
    if (-not $classification.eligible_for_journal_club_visual) {
        return [pscustomobject]@{
            figure_asset_candidates = @()
            suggested_image_path = $null
            selection_mode = 'excluded-by-journal-club-role'
            source_asset_id = $null
        }
    }
    $explicitImagePath = Get-ExplicitFigureAssetPath -Selections $Selections -FigureId $tableId -AssetDirectory $AssetDirectory
    # 页面渲染的表格或矢量图必须明确选择。不能因为同页只有一张 PNG 就自动把整页截图
    # 当作数据表，否则极易将图注、示意图或无关 panel 错放进实验结论页。
    return [pscustomobject]@{
        figure_asset_candidates = if ($explicitImagePath) { @($explicitImagePath) } else { @() }
        suggested_image_path = $explicitImagePath
        selection_mode = if ($explicitImagePath) { 'user-confirmed-page-crop' } else { 'none' }
        source_asset_id = $null
    }
}

function Test-ChineseLanguage {
    param([string]$Language)
    return -not [string]::IsNullOrWhiteSpace($Language) -and $Language -match '^(?i:zh)(-|$)'
}

function Get-JournalClubSectionTitle {
    param([string]$Section, [string]$Language = 'en')
    if (Test-ChineseLanguage $Language) {
        switch ($Section) {
            'background' { return '研究背景与知识缺口' }
            'innovation' { return '创新点与学术贡献' }
            'methods' { return '研究方法与实验设计' }
            'experimental_data' { return '实验数据与核心发现' }
            'limitations' { return '局限性与批判性评价' }
            'future_directions' { return '未来研究方向' }
            default { return '组会汇报内容' }
        }
    }
    switch ($Section) {
        'background' { return 'Research background and knowledge gap' }
        'innovation' { return 'Innovation and contribution' }
        'methods' { return 'Methods and study design' }
        'experimental_data' { return 'Experimental data and core finding' }
        'limitations' { return 'Limitations and critical appraisal' }
        'future_directions' { return 'Future research directions' }
        default { return 'Journal club content' }
    }
}

function New-JournalClubSlide {
    param(
        [string]$Id,
        [string]$Kind,
        [string]$Section,
        [string]$Title,
        [string]$Takeaway,
        [string[]]$Bullets = @(),
        [string[]]$SourceClaimIds = @(),
        [string[]]$SourceSectionIds = @(),
        [string[]]$SourceFigureIds = @(),
        [string]$SuggestedFigureId,
        [string]$SuggestedImagePath,
        [string]$SourceAssetId,
        [string]$SourceVisualId,
        [string]$VisualRoleSource,
        [int]$VisualSourcePage = 0,
        [object]$VisualEvidence = $null,
        [string[]]$FigureAssetCandidates = @(),
        [string]$VisualRole = 'generic',
        [object[]]$ExplanationPoints = @(),
        $ResultAnalysis = $null,
        [string]$NarrativeRole = '',
        [string]$EvidenceStatus = 'source-backed',
        [string]$ContentMode = 'author-direct'
    )

    $safeTakeaway = ConvertTo-ReadableSlideText -Text $Takeaway -ChineseLimit $MaximumTakeawayCharactersChinese -OtherLimit $MaximumTakeawayCharacters
    $safeBullets = @(
        ConvertTo-NonEmptyStringArray $Bullets | ForEach-Object {
            ConvertTo-ReadableSlideText -Text ([string]$_) -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters
        }
    )
    $slide = [ordered]@{
        id = $Id
        kind = $Kind
        section = $Section
        title = $Title
        takeaway = $safeTakeaway
        bullets = $safeBullets
        source_claim_ids = @(ConvertTo-NonEmptyStringArray $SourceClaimIds)
        source_section_ids = @(ConvertTo-NonEmptyStringArray $SourceSectionIds)
        source_figure_ids = @(ConvertTo-NonEmptyStringArray $SourceFigureIds)
        evidence_status = $EvidenceStatus
        content_mode = $ContentMode
    }
    if ($VisualRole -and $VisualRole -ne 'generic') { $slide.visual_role = $VisualRole }
    if ($NarrativeRole) { $slide.narrative_role = $NarrativeRole }
    $safeExplanationPoints = @($ExplanationPoints | Where-Object { $null -ne $_ })
    if ($safeExplanationPoints.Count) { $slide.explanation_points = $safeExplanationPoints }
    if ($null -ne $ResultAnalysis) { $slide.result_analysis = $ResultAnalysis }
    if ($SuggestedFigureId) { $slide.suggested_figure_id = $SuggestedFigureId }
    if ($SuggestedImagePath) { $slide.suggested_image_path = $SuggestedImagePath }
    if ($SourceAssetId) { $slide.source_asset_id = $SourceAssetId }
    if ($SourceVisualId) { $slide.source_visual_id = $SourceVisualId }
    # visual_role 是视觉排版合同；下面三个字段则记录“为什么可以用这张图表”：论文图表 id、
    # PDF 页码和图注/邻近文本证据。即使尚未裁剪图片，结果页也能追溯其推荐图表。
    if ($VisualRoleSource) { $slide.visual_role_source = $VisualRoleSource }
    if ($VisualSourcePage -ge 1) { $slide.visual_source_page = $VisualSourcePage }
    if ($null -ne $VisualEvidence) { $slide.visual_evidence = $VisualEvidence }
    $assetCandidates = @(ConvertTo-NonEmptyStringArray $FigureAssetCandidates)
    if ($assetCandidates.Count) { $slide.figure_asset_candidates = $assetCandidates }
    return [pscustomobject]$slide
}

function New-MissingSourceSlide {
    param([string]$Id, [string]$Section, $FallbackSections, [string]$Language = 'en')

    # 占位页明确标记为待人工核对。审核器会阻止这类必备页进入 PPT 生成阶段。
    $fallbackIds = @()
    foreach ($fallback in @($FallbackSections)) {
        $fallbackId = Get-PropertyValue $fallback 'id'
        if ($fallbackId -and $fallbackId -notin $fallbackIds) { $fallbackIds += [string]$fallbackId }
    }
    $takeaway = if (Test-ChineseLanguage $Language) { '尚未为该必备模块提取到可回链的原文证据。' } else { 'No source-backed content was extracted for this required section.' }
    $bullets = if (Test-ChineseLanguage $Language) {
        @('请核对原论文并补充可追溯来源后再生成 PowerPoint。', '不要用推测性科学结论替换此提示。')
    } else {
        @('Review the original paper and add traceable source evidence before generating PowerPoint.', 'Do not replace this notice with an inferred scientific claim.')
    }
    return New-JournalClubSlide -Id $Id -Kind $Section -Section $Section -Title (Get-JournalClubSectionTitle $Section $Language) `
        -Takeaway $takeaway -Bullets $bullets `
        -SourceSectionIds $fallbackIds -EvidenceStatus 'missing' -ContentMode 'needs-review'
}

function Invoke-DesignDeck {
    param($Arguments)
    $evidencePack = Get-PropertyValue $Arguments 'evidence_pack'
    if (-not $evidencePack -or -not (Get-PropertyValue (Get-PropertyValue $evidencePack 'paper') 'title')) { throw 'A complete evidence_pack is required.' }
    $duration = Assert-IntegerArgument -Arguments $Arguments -Name 'duration_minutes' -Minimum 5 -Maximum 90 -Default 15
    $language = Get-PropertyValue $Arguments 'language' 'zh-CN'
    if ($language -isnot [string] -or [string]::IsNullOrWhiteSpace($language) -or $language.Length -gt 32) { throw 'language must be a non-empty string no longer than 32 characters.' }
    $audience = Get-PropertyValue $Arguments 'audience' 'lab'
    if ($audience -isnot [string] -or $audience -notin @('lab', 'mixed', 'expert')) { throw 'audience must be one of: lab, mixed, expert.' }
    $requiredSections = @(Resolve-RequiredSections (Get-PropertyValue $Arguments 'required_sections'))
    $figureAssetSelection = Get-PropertyValue $Arguments 'figure_asset_selection'
    $extractedAssetDirectory = Get-SafeEvidenceAssetDirectory $evidencePack
    $isChinese = Test-ChineseLanguage $language
    $copy = if ($isChinese) {
        @{
            title_subtitle = '组会汇报 | 可编辑演示文稿草稿'
            journal_club = '文献组会'
            tldr_title = '研究问题与核心结论'
            tldr_review = '请核对摘要并补充有原文依据的总结后再汇报。'
            tldr_notice = '这是一条核对提示，不是科学结论。'
            presenter_direction = '汇报者讨论应围绕已报告局限设计下一步实验。'
            reported_limitation = '论文报告的局限 {0}'
            presenter_proposal = '验证性方案属于汇报者讨论，不应表述为作者结论。'
            takehome_title = '核心结论'
            takehome_guard = '结论必须限定在原始研究的数据和方法边界内。'
            audience_note = '请为 {0} 听众准备讨论问题。'
        }
    } else {
        @{
            title_subtitle = 'Journal Club | Editable presentation draft'
            journal_club = 'Journal club'
            tldr_title = 'Research question and main finding'
            tldr_review = 'Review the abstract and add a source-backed summary before presenting.'
            tldr_notice = 'This is a review notice, not a scientific conclusion.'
            presenter_direction = 'Presenter discussion should address the reported limitation with a next experiment.'
            reported_limitation = 'Reported limitation {0}'
            presenter_proposal = 'Label any validation proposal as presenter discussion rather than an author claim.'
            takehome_title = 'Take-home message'
            takehome_guard = 'Keep conclusions within the limits of the original data and methods.'
            audience_note = 'Prepare discussion for a {0} audience.'
        }
    }
    $claims = @(Get-PropertyValue $evidencePack 'claims' @())
    $figures = @(Get-PropertyValue $evidencePack 'figures' @())
    $tables = @(Get-PropertyValue $evidencePack 'tables' @())
    # deck_spec 可能由缓存、文件导入或调用方编辑而来；每次设计都重新计算角色，不能让
    # 旧版的裸 Figure 或伪造 eligible=true 绕过“只选方法/主实验/消融”的默认策略。
    $figures = @(Set-JournalClubVisualClassification -Visuals $figures -VisualKind 'figure')
    $tables = @(Set-JournalClubVisualClassification -Visuals $tables -VisualKind 'table')

    # 先建立可回链的候选来源。每个正式内容页只使用这些逐句证据，避免模型凭空补全论文结论。
    $abstractSection = Get-FirstSectionByTitlePattern $evidencePack 'abstract|摘要'
    $introductionSection = Get-FirstSectionByTitlePattern $evidencePack 'intro|引言|前言|研究背景'
    $methodsSection = Get-FirstSectionByTitlePattern $evidencePack 'method|material|方法'
    $resultsSection = Get-FirstSectionByTitlePattern $evidencePack 'result|结果'
    $discussionSections = @(Get-SectionsByTitlePattern $evidencePack 'discussion|conclusion|讨论|结论')
    $backgroundSection = if ($introductionSection) { $introductionSection } else { $abstractSection }
    $innovationSections = @()
    $innovationSections += @(Get-SectionsByTitlePattern $evidencePack 'abstract|摘要')
    $innovationSections += @(Get-SectionsByTitlePattern $evidencePack 'intro|引言|前言|研究背景')
    $innovationSections += $discussionSections

    $summaryEvidence = @(Find-SentenceEvidence @($abstractSection) '' 3)
    $backgroundEvidence = @(Find-SentenceEvidence @($backgroundSection) '' 3)
    $innovationPattern = '\b(novel|novelty|first|new (?:method|approach|framework|platform)|we (?:combined|developed|introduced|proposed|designed|built|created|integrated)|our (?:approach|method|framework|platform))\b|创新|首次|新方法|新策略|提出|开发|建立|构建|整合|结合'
    $innovationEvidence = @(Find-SentenceEvidence $innovationSections $innovationPattern 3)
    $methodsEvidence = @(Find-SentenceEvidence @($methodsSection) '' 3)
    $limitationPattern = '\b(limit(?:ation|ed|s)?|constrain(?:t|ed|s)?|caveat|lack(?:s|ed)?|small sample|few (?:donors|patients|samples)|in[ -]?vitro|generaliz(?:e|ability)|may not|cannot|unable|unclear|bias)\b|局限|限制|样本量小|体外|外推|泛化|偏倚|尚不清楚|不能'
    $limitationEvidence = @(Find-SentenceEvidence $discussionSections $limitationPattern 3)
    $futurePattern = '\b(future|next step|further (?:work|stud(?:y|ies)|research)|follow-up|will need|should (?:test|evaluate|validate|examine)|warrant(?:s|ed)?|remain(?:s)? to be (?:tested|determined|established)|prospective)\b|未来|下一步|后续|进一步|有待|需要验证|值得探讨'
    $futureEvidence = @(Find-SentenceEvidence $discussionSections $futurePattern 3)

    $slides = @()
    $titleSlide = New-JournalClubSlide -Id 'title' -Kind 'title' -Section '' -Title $evidencePack.paper.title -Takeaway $copy.journal_club `
        -NarrativeRole 'title' -EvidenceStatus 'not-applicable' -ContentMode 'metadata'
    $titleSlide | Add-Member -NotePropertyName subtitle -NotePropertyValue $copy.title_subtitle
    $slides += $titleSlide

    if ($summaryEvidence.Count) {
        $summaryText = @($summaryEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'tl-dr' -Kind 'summary' -Section 'summary' -Title $copy.tldr_title `
            -Takeaway $summaryText[0] -Bullets $summaryText -SourceSectionIds (Get-EvidenceItemSectionIds $summaryEvidence) -NarrativeRole 'research-question-and-conclusion'
    } else {
        $slides += New-JournalClubSlide -Id 'tl-dr' -Kind 'summary' -Section 'summary' -Title $copy.tldr_title `
            -Takeaway $copy.tldr_review -Bullets @($copy.tldr_notice) -NarrativeRole 'research-question-and-conclusion' -EvidenceStatus 'missing' -ContentMode 'needs-review'
    }

    if ($backgroundEvidence.Count) {
        $backgroundText = @($backgroundEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'background' -Kind 'background' -Section 'background' -Title (Get-JournalClubSectionTitle 'background' $language) `
            -Takeaway $backgroundText[0] -Bullets $backgroundText -SourceSectionIds (Get-EvidenceItemSectionIds $backgroundEvidence) -NarrativeRole 'knowledge-gap'
    } elseif ('background' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'background' -Section 'background' -FallbackSections @($backgroundSection) -Language $language
    }

    if ($innovationEvidence.Count) {
        $innovationText = @($innovationEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'innovation' -Kind 'innovation' -Section 'innovation' -Title (Get-JournalClubSectionTitle 'innovation' $language) `
            -Takeaway $innovationText[0] -Bullets $innovationText -SourceSectionIds (Get-EvidenceItemSectionIds $innovationEvidence) -NarrativeRole 'study-contribution'
    } elseif ('innovation' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'innovation' -Section 'innovation' -FallbackSections $innovationSections -Language $language
    }

    if ($methodsEvidence.Count) {
        $methodsText = @($methodsEvidence | ForEach-Object { $_.text })
        # 方法页仅接收已分类为系统架构/流程图的 Figure；案例图、性能曲线和未分类
        # Figure 即使在 Methods 段中被提到，也只保留文字说明，不自动占用视觉区域。
        $methodsFigureId = Find-JournalClubVisualIdForSection -Section $methodsSection -Figures $figures
        $methodsFigureIds = if ($methodsFigureId) { @([string]$methodsFigureId) } else { @() }
        $methodsFigure = Get-FigureById -Figures $figures -FigureId $methodsFigureId
        # 即使还没有选择裁剪图，也应预先生成由 Methods 与 Figure caption/context 支撑的
        # 讲解点；用户稍后确认图片后，生成器会直接把这些文字排进系统图幻灯片。
        $methodExplanationPoints = @(Get-MethodExplanationEvidence -MethodsSection $methodsSection -Figure $methodsFigure -Language $language)
        $methodsVisualTraceability = Get-JournalClubVisualTraceability -Visual $methodsFigure
        $methodsFigureBinding = Resolve-FigureImageBinding -Figure $methodsFigure -Selections $figureAssetSelection -AssetDirectory $extractedAssetDirectory
        $slides += New-JournalClubSlide -Id 'methods' -Kind 'methods' -Section 'methods' -Title (Get-JournalClubSectionTitle 'methods' $language) `
            -Takeaway $methodsText[0] -Bullets $methodsText -SourceSectionIds (Get-EvidenceItemSectionIds $methodsEvidence) -SourceFigureIds $methodsFigureIds `
            -SuggestedFigureId $methodsFigureId -SuggestedImagePath $methodsFigureBinding.suggested_image_path -SourceAssetId $methodsFigureBinding.source_asset_id -SourceVisualId $methodsFigureId -FigureAssetCandidates $methodsFigureBinding.figure_asset_candidates `
            -VisualRoleSource $methodsVisualTraceability.role -VisualSourcePage $methodsVisualTraceability.source_page -VisualEvidence $methodsVisualTraceability.evidence `
            -VisualRole $(if ($methodsFigureBinding.suggested_image_path) { 'system-architecture' } else { 'methods-explanation' }) -ExplanationPoints $methodExplanationPoints -NarrativeRole 'study-design-and-system'
    } elseif ('methods' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'methods' -Section 'methods' -FallbackSections @($methodsSection) -Language $language
    }

    $maxResults = [Math]::Max(1, [Math]::Min(4, [Math]::Round($duration / 5)))
    $resultClaims = @(Select-NarrativeResultClaims -Claims $claims -Maximum $maxResults)
    $preferredResultVisualId = Find-PreferredResultVisualId -Claims $resultClaims -Figures $figures -Tables $tables
    if ($resultClaims.Count) {
        for ($i = 0; $i -lt $resultClaims.Count; $i++) {
            $claim = $resultClaims[$i]
            $claimId = Get-PropertyValue $claim 'id'
            $claimText = Get-PropertyValue $claim 'text'
            $claimSectionIds = @(Get-ClaimEvidenceSectionIds $claim)
            # 结果页只连接 main-result / ablation。这里返回单一、已确认角色的 id，
            # 不再按“Figure 优先、Table 其次”的宽松规则把案例或数据集统计带入。
            $visualId = Find-JournalClubVisualIdForClaim -Claim $claim -Figures $figures -Tables $tables
            if (-not $visualId -and $i -eq 0 -and $preferredResultVisualId) {
                # 仅为首张结果页补一个唯一、可追溯的主实验/消融候选；不把它伪装成 claim
                # 的直接图号引用，仍保留 claim 自己的原文证据和图表的 caption/page 证据。
                $visualId = $preferredResultVisualId
            }
            $figureId = if ($visualId -like 'fig-*') { $visualId } else { $null }
            $tableId = if ($visualId -like 'table-*') { $visualId } else { $null }
            $visualIds = if ($visualId) { @([string]$visualId) } else { @() }
            $sourceVisual = Get-VisualById -Figures $figures -Tables $tables -VisualId $visualId
            $visualTraceability = Get-JournalClubVisualTraceability -Visual $sourceVisual
            $figureBinding = if ($tableId) {
                Resolve-TableImageBinding -Table $sourceVisual -Selections $figureAssetSelection -AssetDirectory $extractedAssetDirectory
            } else {
                Resolve-FigureImageBinding -Figure $sourceVisual -Selections $figureAssetSelection -AssetDirectory $extractedAssetDirectory
            }
            $counterEvidenceId = [string](Get-PropertyValue $claim 'counter_evidence_claim_id' '')
            $counterEvidenceText = [string](Get-PropertyValue $claim 'counter_evidence_text' '')
            $counterEvidence = $null
            if ($counterEvidenceId) {
                $counterEvidence = @($claims | Where-Object { ([string](Get-PropertyValue $_ 'id' '')) -eq $counterEvidenceId } | Select-Object -First 1)[0]
            }
            # New-ResultAnalysis 只使用 evidence pack 内的文本证据和已分类图表的图注。
            # 它不读取 PNG 像素，因此尚未确认裁剪图时也能安全预生成口播分析。
            $resultAnalysis = New-ResultAnalysis -Claim $claim -Visual $sourceVisual -EvidencePack $evidencePack `
                -CounterEvidence $counterEvidence -Language $language
            if ($counterEvidenceId -and $counterEvidenceText) {
                # 选择器在五分钟模式把已确认的反例合并到主 claim。此处再次以该原文覆盖
                # caveat，确保短汇报不会因内容压缩而丢掉作者报告的负面/无改善证据。
                $counterDisplayPrefix = if (Test-ChineseLanguage $language) { '论文报告的反例是' } else { 'The paper also reports' }
                $resultAnalysis.caveat = ConvertTo-ReadableSlideText -Text "$counterDisplayPrefix $counterEvidenceText" `
                    -ChineseLimit $MaximumBulletCharactersChinese -OtherLimit $MaximumBulletCharacters
                $resultAnalysis.caveat_kind = 'reported-counter-claim'
                $resultAnalysis.caveat_source_claim_id = $counterEvidenceId
                $resultAnalysis.caveat_source_section_id = $null
                $resultAnalysis.caveat_source_excerpt = ConvertTo-ReadableSlideText -Text $counterEvidenceText -ChineseLimit 180 -OtherLimit 360
            }
            $slideBullets = @($resultAnalysis.interpretation, $resultAnalysis.caveat)
            $claimEvidenceStatus = if ($claimSectionIds.Count -and $claimId -and $claimText) { 'source-backed' } else { 'missing' }
            $resultClaimIds = @([string]$claimId)
            if ($counterEvidenceId -and $counterEvidenceId -notin $resultClaimIds) { $resultClaimIds += $counterEvidenceId }
            $slides += New-JournalClubSlide -Id "experimental-data-$($i + 1)" -Kind 'result' -Section 'experimental_data' `
                -Title (Get-ResultSlideTitle -Claim $claim -Language $language) -Takeaway $claimText `
                -Bullets $slideBullets `
                -SourceClaimIds $resultClaimIds -SourceSectionIds $claimSectionIds -SourceFigureIds $visualIds `
                -SuggestedFigureId $visualId -SuggestedImagePath $figureBinding.suggested_image_path -SourceAssetId $figureBinding.source_asset_id -SourceVisualId $visualId -FigureAssetCandidates $figureBinding.figure_asset_candidates `
                -VisualRoleSource $visualTraceability.role -VisualSourcePage $visualTraceability.source_page -VisualEvidence $visualTraceability.evidence `
                -VisualRole $(if ($figureBinding.suggested_image_path -and $tableId) { 'result-table' } elseif ($figureBinding.suggested_image_path) { 'result-figure' } else { 'result-analysis' }) -ResultAnalysis $resultAnalysis -NarrativeRole 'experimental-evidence' `
                -EvidenceStatus $claimEvidenceStatus -ContentMode 'author-direct'
        }
    } elseif ('experimental_data' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'experimental-data' -Section 'experimental_data' -FallbackSections @($resultsSection) -Language $language
    }

    if ($limitationEvidence.Count) {
        $limitationText = @($limitationEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'limitations' -Kind 'limitations' -Section 'limitations' -Title (Get-JournalClubSectionTitle 'limitations' $language) `
            -Takeaway $limitationText[0] -Bullets $limitationText -SourceSectionIds (Get-EvidenceItemSectionIds $limitationEvidence) -NarrativeRole 'critical-appraisal'
    } elseif ('limitations' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'limitations' -Section 'limitations' -FallbackSections $discussionSections -Language $language
    }

    if ($futureEvidence.Count) {
        $futureText = @($futureEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'future-directions' -Kind 'future-directions' -Section 'future_directions' -Title (Get-JournalClubSectionTitle 'future_directions' $language) `
            -Takeaway $futureText[0] -Bullets $futureText -SourceSectionIds (Get-EvidenceItemSectionIds $futureEvidence) -NarrativeRole 'next-step'
    } elseif ($limitationEvidence.Count) {
        # 原文未给出下一步时，只生成明确标为“汇报者讨论”的建议，并回链到作者已述局限。
        $reportedLimitation = $limitationEvidence[0].text
        $slides += New-JournalClubSlide -Id 'future-directions' -Kind 'future-directions' -Section 'future_directions' -Title (Get-JournalClubSectionTitle 'future_directions' $language) `
            -Takeaway $copy.presenter_direction `
            -Bullets @(($copy.reported_limitation -f $reportedLimitation), $copy.presenter_proposal) `
            -SourceSectionIds (Get-EvidenceItemSectionIds $limitationEvidence) -NarrativeRole 'next-step' -EvidenceStatus 'source-backed' -ContentMode 'presenter-discussion'
    } elseif ('future_directions' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'future-directions' -Section 'future_directions' -FallbackSections $discussionSections -Language $language
    }

    if ($claims.Count) {
        $takeawayClaimIds = @($claims | Select-Object -First 3 | ForEach-Object { Get-PropertyValue $_ 'id' })
        $takeawaySectionIds = @()
        foreach ($claim in @($claims | Select-Object -First 3)) {
            foreach ($sectionId in @(Get-ClaimEvidenceSectionIds $claim)) {
                if ($sectionId -notin $takeawaySectionIds) { $takeawaySectionIds += $sectionId }
            }
        }
        $slides += New-JournalClubSlide -Id 'takeaway' -Kind 'takeaway' -Section 'takeaway' -Title $copy.takehome_title `
            -Takeaway (Get-PropertyValue $claims[0] 'text') `
            -Bullets @($copy.takehome_guard, ($copy.audience_note -f $audience)) `
            -SourceClaimIds $takeawayClaimIds -SourceSectionIds $takeawaySectionIds -NarrativeRole 'closing-conclusion'
    }

    $deck = [pscustomobject]@{
        schema_version = '0.3'
        deck_template = 'journal-club'
        required_sections = $requiredSections
        language = $language
        duration_minutes = $duration
        audience = $audience
        paper = $evidencePack.paper
        evidence_pack = $evidencePack
        theme = [pscustomobject]@{ primary = '#114B5F'; accent = '#F45B69'; background = '#FFFFFF'; font = 'Aptos' }
        slides = $slides
    }
    $outputPath = Get-PropertyValue $Arguments 'output_path'
    if ($outputPath) {
        $overwrite = Get-StrictBoolean -Object $Arguments -Name 'overwrite' -Default $false
        $absoluteOutput = Resolve-RequestedOutputPath -Path $outputPath -RequiredExtension '.json' -Overwrite $overwrite
        $deck | ConvertTo-Json -Depth 50 | Set-Content -Encoding UTF8 -LiteralPath $absoluteOutput
    }
    return $deck
}

function Invoke-AuditDeck {
    param($Arguments)
    $deck = Get-PropertyValue $Arguments 'deck_spec'
    if (-not $deck) { throw 'deck_spec is required.' }
    $evidencePack = Get-PropertyValue $deck 'evidence_pack'
    if (-not $evidencePack) { throw 'deck_spec.evidence_pack is required for traceability audit.' }
    $requiredSections = @(Resolve-RequiredSections (Get-PropertyValue $deck 'required_sections'))
    $claims = @(Get-PropertyValue $evidencePack 'claims' @())
    $sourceSections = @(Get-PropertyValue $evidencePack 'sections' @())
    $figures = @(Get-PropertyValue $evidencePack 'figures' @())
    $tables = @(Get-PropertyValue $evidencePack 'tables' @())
    $claimIds = Get-ObjectIds $claims
    $sectionIds = Get-ObjectIds $sourceSections
    $figureIds = Get-ObjectIds (@($figures) + @($tables))
    $claimsById = @{}
    foreach ($claim in $claims) {
        $claimId = Get-PropertyValue $claim 'id'
        if ($claimId) { $claimsById[[string]$claimId] = $claim }
    }
    # Get-ObjectIds 只用于快速判断 id 是否存在，值是 $true。系统结构图的讲解审查
    # 还需要读取章节标题来确认它确实来自方法/材料部分，因此保留一份 id 到原章节
    # 对象的映射，不能把布尔值误当成章节对象。
    $sectionsById = @{}
    foreach ($sourceSection in $sourceSections) {
        $sourceSectionId = Get-PropertyValue $sourceSection 'id'
        if ($sourceSectionId) { $sectionsById[[string]$sourceSectionId] = $sourceSection }
    }
    $findings = @()
    $slides = @(Get-PropertyValue $deck 'slides' @())
    try {
        # 与生成入口共享同一套阅读性边界。这里转成 hard finding 而不是让审核工具直接崩溃，
        # 便于调用方知道该改哪一页。
        Assert-DeckSpecificationLimits -Deck $deck
    } catch {
        $findings += [pscustomobject]@{ slide = $null; severity = 'hard'; category = 'readability'; issue = $_.Exception.Message; correction = 'Use short declarative headings, up to three concise points, and readable text lengths.' }
    }
    $requiredSlideIndexes = @{}
    foreach ($section in $requiredSections) { $requiredSlideIndexes[$section] = @() }

    for ($index = 0; $index -lt $slides.Count; $index++) {
        $slide = $slides[$index]
        $slideNumber = $index + 1
        $slideSection = [string](Get-PropertyValue $slide 'section' '')
        $isRequiredContentSlide = $slideSection -in $requiredSections
        if ($isRequiredContentSlide) { $requiredSlideIndexes[$slideSection] += $index }

        if (-not (Get-PropertyValue $slide 'title' '')) {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'title'; issue = 'Missing title'; correction = 'Set a clear slide title.' }
        }
        $slideKind = [string](Get-PropertyValue $slide 'kind' '')
        $slideTitle = [string](Get-PropertyValue $slide 'title' '')
        if ($slideKind -ne 'title' -and $slideTitle -match '[:：?？]') {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'title-style'; issue = 'Non-title heading uses a colon or question mark.'; correction = 'Use one short declarative heading without a colon or question mark.' }
        }
        if ((Get-PropertyValue $slide 'kind') -ne 'title' -and -not (Get-PropertyValue $slide 'takeaway' '')) {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'narrative'; issue = 'Missing takeaway'; correction = 'Add one talkable takeaway.' }
        }

        $slideClaimIds = @(ConvertTo-NonEmptyStringArray (Get-PropertyValue $slide 'source_claim_ids' @()))
        $slideSectionIds = @(ConvertTo-NonEmptyStringArray (Get-PropertyValue $slide 'source_section_ids' @()))
        $slideFigureIds = @(ConvertTo-NonEmptyStringArray (Get-PropertyValue $slide 'source_figure_ids' @()))
        foreach ($claimId in $slideClaimIds) {
            if (-not $claimIds.ContainsKey($claimId)) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'traceability'; issue = "Unknown claim id: $claimId"; correction = 'Reference an existing evidence-pack claim.' }
                continue
            }
            $claim = $claimsById[$claimId]
            $claimEvidenceItems = @(Get-PropertyValue $claim 'evidence' @())
            if ($claimEvidenceItems.Count -eq 0) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'traceability'; issue = "Claim $claimId has no source evidence."; correction = 'Add a section id and original-paper excerpt to the claim evidence.' }
            } else {
                foreach ($claimEvidence in $claimEvidenceItems) {
                    $evidenceSectionId = Get-PropertyValue $claimEvidence 'section_id'
                    $evidenceExcerpt = Get-PropertyValue $claimEvidence 'excerpt'
                    if (-not $evidenceSectionId -or -not $sectionIds.ContainsKey([string]$evidenceSectionId)) {
                        $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'traceability'; issue = "Claim $claimId refers to an unknown evidence section: $evidenceSectionId"; correction = 'Use an existing evidence-pack section id in claim.evidence.' }
                    }
                    if (-not $evidenceExcerpt) {
                        $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'traceability'; issue = "Claim $claimId has an evidence item without an original-paper excerpt."; correction = 'Store the source sentence or a faithful short excerpt in claim.evidence.' }
                    }
                }
            }
            if ((Get-PropertyValue $claim 'confidence' '') -eq 'needs-review') {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'warning'; category = 'scientific-review'; issue = "Claim $claimId is still marked needs-review."; correction = 'Confirm the original wording, statistics, and causal scope before presentation.' }
            }
        }
        foreach ($sectionId in $slideSectionIds) {
            if (-not $sectionIds.ContainsKey($sectionId)) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'traceability'; issue = "Unknown source section id: $sectionId"; correction = 'Reference an existing evidence-pack section.' }
            }
        }
        foreach ($figureId in $slideFigureIds) {
            if (-not $figureIds.ContainsKey($figureId)) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'traceability'; issue = "Unknown source figure id: $figureId"; correction = 'Reference an existing evidence-pack figure.' }
            }
        }
        $suggestedImagePath = Get-PropertyValue $slide 'suggested_image_path'
        $sourceAssetId = [string](Get-PropertyValue $slide 'source_asset_id' '')
        $figureAssetCandidates = @(ConvertTo-NonEmptyStringArray (Get-PropertyValue $slide 'figure_asset_candidates' @()) | ForEach-Object { [IO.Path]::GetFullPath($_) })
        if ($suggestedImagePath) {
            $absoluteImagePath = [IO.Path]::GetFullPath([string]$suggestedImagePath)
            if ($slideFigureIds.Count -eq 0) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'figure-asset'; issue = 'Suggested image has no source figure id.'; correction = 'Link the image to a figure id from the evidence pack.' }
            }
            if ($figureAssetCandidates.Count -eq 0 -or $absoluteImagePath -notin $figureAssetCandidates) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'figure-asset'; issue = 'Suggested image is not declared as an explicit figure asset candidate.'; correction = 'Use figure_asset_selection or add the selected path to figure_asset_candidates.' }
            }
            if (-not (Test-Path -LiteralPath $absoluteImagePath -PathType Leaf)) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'figure-asset'; issue = "Suggested image file was not found: $absoluteImagePath"; correction = 'Select an existing extracted paper image or remove the image reference.' }
            }
            # source_asset_id 不是手工填写的要求，而是自动按 Fig 编号绑定时的可选溯源信息。
            # 若它出现，必须能在该 Figure 的 automatic_binding 中精确找到，避免 deck spec
            # 用一个正确图片路径却冒充另一张 PDF 资产的编号。
            if ($sourceAssetId) {
                $matchingFigureBindings = @($figures | Where-Object {
                    $binding = Get-PropertyValue $_ 'automatic_binding'
                    $figureId = [string](Get-PropertyValue $_ 'id' '')
                    $binding -and $figureId -in $slideFigureIds -and
                    ([string](Get-PropertyValue $binding 'asset_id' '')) -eq $sourceAssetId -and
                    ([string](Get-PropertyValue $binding 'path' '')) -eq $absoluteImagePath -and
                    ([string](Get-PropertyValue $binding 'figure_id' '')) -eq $figureId
                })
                if ($matchingFigureBindings.Count -ne 1) {
                    $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'figure-asset'; issue = 'source_asset_id is not a verified automatic Figure-to-asset binding.'; correction = 'Use the asset id returned by analyse_paper for the same Figure, or omit source_asset_id for a reviewed manual selection.' }
                }
            }
        } elseif ($sourceAssetId) {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'figure-asset'; issue = 'source_asset_id is present without an inserted image.'; correction = 'Remove source_asset_id or attach its matching approved image.' }
        }

        $hasTraceableSource = ($slideClaimIds.Count + $slideSectionIds.Count + $slideFigureIds.Count) -gt 0
        $evidenceStatus = [string](Get-PropertyValue $slide 'evidence_status' '')
        $contentMode = [string](Get-PropertyValue $slide 'content_mode' '')
        if ($isRequiredContentSlide) {
            if (-not $hasTraceableSource) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'required-section-evidence'; issue = "Required section '$slideSection' has no source reference."; correction = 'Add a claim, paper section, or figure id from the evidence pack.' }
            }
            if ($evidenceStatus -ne 'source-backed') {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'required-section-evidence'; issue = "Required section '$slideSection' is marked $evidenceStatus."; correction = 'Replace the placeholder with source-backed content or revise required_sections.' }
            }
            if ($contentMode -eq 'needs-review' -or -not $contentMode) {
                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'required-section-evidence'; issue = "Required section '$slideSection' is not ready for factual presentation."; correction = 'Use an author-direct source or an explicitly labelled presenter discussion with evidence.' }
            }
            if ($slideSection -eq 'experimental_data') {
                if ($slideClaimIds.Count -eq 0) {
                    $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-data'; issue = 'Experimental-data slide has no source claim id.'; correction = 'Link each data slide to a result claim with an original-paper excerpt.' }
                }
                if ($slideFigureIds.Count -eq 0) {
                    $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'warning'; category = 'experimental-data'; issue = 'Experimental-data slide has no figure reference.'; correction = 'Add the corresponding figure id when the paper contains a relevant figure.' }
                }
                $resultAnalysis = Get-PropertyValue $slide 'result_analysis'
                if ($null -eq $resultAnalysis) {
                    $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-data'; issue = 'Experimental-data slide has no source-grounded comparison analysis.'; correction = 'Add comparison, interpretation, and caveat based on the cited result sentence.' }
                } else {
                    Assert-McpObject -Value $resultAnalysis -ParameterName 'result_analysis'
                    foreach ($analysisName in @('comparison', 'interpretation', 'caveat')) {
                        if ([string]::IsNullOrWhiteSpace([string](Get-PropertyValue $resultAnalysis $analysisName ''))) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-data'; issue = "Experimental-data slide is missing result_analysis.$analysisName."; correction = 'State the source-backed comparison, interpretation, and caveat.' }
                        }
                    }
                    # 主实验与消融页必须把分析绑定到同一张已引用的 Figure/Table。没有视觉
                    # 也允许先生成文字，但一旦指定 visual，就不能用其他图的 caption 或
                    # 截图读数替代论文文本。
                    $analysisVisualId = [string](Get-PropertyValue $resultAnalysis 'visual_id' '')
                    if ($slideFigureIds.Count -gt 0) {
                        if (-not $analysisVisualId -or $analysisVisualId -notin $slideFigureIds) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis is not bound to the slide source figure or table id.'; correction = 'Set result_analysis.visual_id to the cited main-result or ablation visual.' }
                        }
                    }
                    if ($analysisVisualId) {
                        $analysisVisual = Get-VisualById -Figures $figures -Tables $tables -VisualId $analysisVisualId
                        if ($null -eq $analysisVisual) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis references an unknown visual id.'; correction = 'Use a Figure/Table id from the evidence pack.' }
                        } else {
                            $analysisRole = [string](Get-PropertyValue $analysisVisual 'journal_club_role' '')
                            if ($analysisRole -notin @('main-result', 'ablation')) {
                                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = "Result analysis is bound to an ineligible visual role: $analysisRole."; correction = 'Use only a main-result or ablation Figure/Table for the default journal-club experimental narrative.' }
                            }
                            $analysisCaption = [string](Get-PropertyValue $resultAnalysis 'visual_caption_excerpt' '')
                            if ($analysisCaption -and -not (Test-TraceablePaperExcerpt -Excerpt $analysisCaption -SourceTexts @([string](Get-PropertyValue $analysisVisual 'context' '')))) {
                                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis visual caption excerpt is not traceable to the cited Figure/Table context.'; correction = 'Use the caption or nearby text returned in the evidence pack; do not infer it from the screenshot.' }
                            }
                        }
                    }
                    $comparisonClaimId = [string](Get-PropertyValue $resultAnalysis 'comparison_source_claim_id' '')
                    if (-not $comparisonClaimId -or $comparisonClaimId -notin $slideClaimIds -or -not $claimsById.ContainsKey($comparisonClaimId)) {
                        $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis comparison is not linked to a cited source claim.'; correction = 'Set comparison_source_claim_id to one of the slide source_claim_ids.' }
                    } else {
                        $comparisonClaim = $claimsById[$comparisonClaimId]
                        $comparisonSourceTexts = @((Get-PropertyValue $comparisonClaim 'evidence' @()) | ForEach-Object { [string](Get-PropertyValue $_ 'excerpt' '') }) + @([string](Get-PropertyValue $comparisonClaim 'text' ''))
                        $comparisonExcerpt = [string](Get-PropertyValue $resultAnalysis 'comparison_source_excerpt' '')
                        if (-not (Test-TraceablePaperExcerpt -Excerpt $comparisonExcerpt -SourceTexts $comparisonSourceTexts)) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis comparison excerpt is not traceable to the cited paper claim.'; correction = 'Use the original-paper claim evidence; do not infer values or significance from an image.' }
                        }
                    }
                    $authorSectionId = [string](Get-PropertyValue $resultAnalysis 'author_explanation_source_section_id' '')
                    $authorExcerpt = [string](Get-PropertyValue $resultAnalysis 'author_explanation_excerpt' '')
                    if (-not $authorSectionId -or -not $sectionIds.ContainsKey($authorSectionId)) {
                        $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis author explanation has no valid source section id.'; correction = 'Link the author explanation to a results or discussion section in the evidence pack.' }
                    } elseif (-not (Test-TraceablePaperExcerpt -Excerpt $authorExcerpt -SourceTexts @((Get-PropertyValue $sectionsById[$authorSectionId] 'excerpt' '')))) {
                        $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis author explanation excerpt is not traceable to its source section.'; correction = 'Quote or faithfully shorten the paper text; do not invent a mechanism from the visual.' }
                    }
                    $caveatSectionId = [string](Get-PropertyValue $resultAnalysis 'caveat_source_section_id' '')
                    $caveatClaimId = [string](Get-PropertyValue $resultAnalysis 'caveat_source_claim_id' '')
                    $caveatExcerpt = [string](Get-PropertyValue $resultAnalysis 'caveat_source_excerpt' '')
                    $caveatKind = [string](Get-PropertyValue $resultAnalysis 'caveat_kind' '')
                    if ($caveatKind -eq 'evidence-boundary') {
                        if ($caveatSectionId -or $caveatClaimId -or $caveatExcerpt) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Evidence-boundary caveat must not claim an unverified paper source.'; correction = 'Either cite an actual limitation/counterexample or retain an explicitly source-free display boundary.' }
                        }
                    } elseif ($caveatClaimId) {
                        if ($caveatClaimId -notin $slideClaimIds -or -not $claimsById.ContainsKey($caveatClaimId)) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result-analysis caveat claim is not cited by the slide.'; correction = 'Add the reported counterexample claim to source_claim_ids.' }
                        } else {
                            $caveatClaim = $claimsById[$caveatClaimId]
                            $caveatSourceTexts = @((Get-PropertyValue $caveatClaim 'evidence' @()) | ForEach-Object { [string](Get-PropertyValue $_ 'excerpt' '') }) + @([string](Get-PropertyValue $caveatClaim 'text' ''))
                            if (-not (Test-TraceablePaperExcerpt -Excerpt $caveatExcerpt -SourceTexts $caveatSourceTexts)) {
                                $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result-analysis caveat excerpt is not traceable to the cited counterexample.'; correction = 'Use a reported counterexample sentence rather than a screenshot-derived statement.' }
                            }
                        }
                    } elseif ($caveatSectionId) {
                        if (-not $sectionIds.ContainsKey($caveatSectionId) -or -not (Test-TraceablePaperExcerpt -Excerpt $caveatExcerpt -SourceTexts @((Get-PropertyValue $sectionsById[$caveatSectionId] 'excerpt' '')))) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result-analysis caveat excerpt is not traceable to its source section.'; correction = 'Use an author-reported limitation or counterexample from the evidence pack.' }
                        }
                    } else {
                        $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'experimental-analysis'; issue = 'Result analysis caveat has no source or explicit evidence-boundary marker.'; correction = 'Cite a limitation/counterexample or mark the display boundary explicitly.' }
                    }
                }
            }
            # 方法页一旦使用论文系统图，就必须解释图中的输入、模块、输出或验证。不能靠
            # 把 visual_role 手工改成 generic 绕过讲解要求，否则会退化为只有一张图的幻灯片。
            $methodHasVisual = $slideSection -eq 'methods' -and -not [string]::IsNullOrWhiteSpace([string](Get-PropertyValue $slide 'suggested_image_path' ''))
            if ($methodHasVisual) {
                $explanationPoints = @(Get-PropertyValue $slide 'explanation_points' @())
                if ($explanationPoints.Count -lt 2) {
                    $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'methods-explanation'; issue = 'System architecture visual has fewer than two source-backed explanation points.'; correction = 'Explain the system input, key module, output, or validation with two to four paper-backed points.' }
                } else {
                    foreach ($point in $explanationPoints) {
                        # 纯字符串可能来自旧 deck spec；一旦方法图作为证据进入新流程，每条解释
                        # 都必须显式回链到 evidence_pack 的方法章节，不能只用通用话术填满左栏。
                        if ($point -is [string]) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'methods-explanation'; issue = 'System architecture explanation is missing source_section_id.'; correction = 'Use explanation_points objects with a valid methods-section source_section_id.' }
                            continue
                        }
                        $pointSectionId = [string](Get-PropertyValue $point 'source_section_id' '')
                        if (-not $pointSectionId -or -not $sectionIds.ContainsKey($pointSectionId)) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'methods-explanation'; issue = 'System architecture explanation has an unknown source_section_id.'; correction = 'Link each explanation point to an existing methods section in the evidence pack.' }
                            continue
                        }
                        $sourceSection = $sectionsById[$pointSectionId]
                        if ([string](Get-PropertyValue $sourceSection 'title' '') -notmatch 'method|material|方法') {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'methods-explanation'; issue = 'System architecture explanation is not linked to a methods section.'; correction = 'Link system explanation points to the paper methods/materials section.' }
                        }
                        $pointFigureId = [string](Get-PropertyValue $point 'source_figure_id' '')
                        if ($pointFigureId -and $pointFigureId -notin $slideFigureIds) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'methods-explanation'; issue = 'System architecture explanation references a Figure not cited by the methods slide.'; correction = 'Use the same system Figure id in source_figure_ids and explanation_points.' }
                        }
                        $pointExcerpt = [string](Get-PropertyValue $point 'source_excerpt' '')
                        if ($pointExcerpt -and -not (Test-TraceablePaperExcerpt -Excerpt $pointExcerpt -SourceTexts @((Get-PropertyValue $sourceSection 'excerpt' '')) + @([string](Get-PropertyValue $point 'text' '')))) {
                            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'methods-explanation'; issue = 'System architecture explanation excerpt is not traceable to the Methods text.'; correction = 'Use a Methods excerpt or a labelled Figure caption context; do not narrate unverified diagram details.' }
                        }
                    }
                }
            }
            if ($slideSection -eq 'future_directions' -and $contentMode -eq 'presenter-discussion') {
                if ($slideSectionIds.Count -eq 0 -or (Get-PropertyValue $slide 'takeaway' '') -notmatch '(?i)presenter|discussion|汇报者|讨论') {
                    $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'future-directions'; issue = 'Presenter-derived future direction is not clearly labelled and sourced.'; correction = 'Label it as presenter discussion and cite the limitation or result that motivates it.' }
                }
            }
        } elseif ((Get-PropertyValue $slide 'kind') -ne 'title' -and -not $hasTraceableSource -and $evidenceStatus -ne 'not-applicable') {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'warning'; category = 'traceability'; issue = 'Non-title slide has no traceable source.'; correction = 'Add a claim, paper section, or figure reference before final presentation.' }
        }
        if (@(Get-PropertyValue $slide 'bullets' @()).Count -gt $MaximumBulletsPerSlide) {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'readability'; issue = "More than $MaximumBulletsPerSlide bullets"; correction = 'Split the slide or retain only the strongest supporting points.' }
        }
    }

    # 结构门禁：默认六个模块或调用方显式配置的模块，任意一个缺页都会阻断生成。
    foreach ($requiredSection in $requiredSections) {
        if (@($requiredSlideIndexes[$requiredSection]).Count -eq 0) {
            $findings += [pscustomobject]@{ slide = $null; severity = 'hard'; category = 'required-section'; issue = "Missing required section: $requiredSection"; correction = "Add a slide with section='$requiredSection' and source-backed content." }
        }
    }

    # 组会不是目录的随机拼接。只要这些模块存在，就强制研究问题→背景/创新→方法→证据→
    # 批判与下一步→结论的顺序；这样生成的 PPT 更接近可直接讲述的故事线。
    $firstIndexBySection = @{}
    foreach ($sectionName in @('summary', 'background', 'innovation', 'methods', 'experimental_data', 'limitations', 'future_directions', 'takeaway')) {
        $firstMatch = @($slides | Where-Object { [string](Get-PropertyValue $_ 'section' '') -eq $sectionName } | Select-Object -First 1)
        if ($firstMatch.Count) { $firstIndexBySection[$sectionName] = [array]::IndexOf($slides, $firstMatch[0]) }
    }
    foreach ($pair in @(@('background', 'innovation'), @('innovation', 'methods'), @('methods', 'experimental_data'), @('experimental_data', 'limitations'), @('limitations', 'future_directions'))) {
        if ($firstIndexBySection.ContainsKey($pair[0]) -and $firstIndexBySection.ContainsKey($pair[1]) -and $firstIndexBySection[$pair[0]] -gt $firstIndexBySection[$pair[1]]) {
            $findings += [pscustomobject]@{ slide = $null; severity = 'hard'; category = 'narrative-order'; issue = "Section '$($pair[0])' appears after '$($pair[1])'."; correction = 'Restore the research question, study design, evidence, critique, and conclusion sequence.' }
        }
    }
    if ($firstIndexBySection.ContainsKey('takeaway') -and $firstIndexBySection['takeaway'] -ne ($slides.Count - 1)) {
        $findings += [pscustomobject]@{ slide = $firstIndexBySection['takeaway'] + 1; severity = 'hard'; category = 'narrative-order'; issue = 'The closing conclusion is not the final slide.'; correction = 'Move the take-home conclusion to the end of the deck.' }
    }

    $hardFindings = @($findings | Where-Object { $_.severity -eq 'hard' })
    $warningFindings = @($findings | Where-Object { $_.severity -eq 'warning' })
    $requiredFindings = @($hardFindings | Where-Object { $_.category -like 'required-section*' -or $_.category -eq 'experimental-data' -or $_.category -eq 'future-directions' })
    $traceabilityFindings = @($hardFindings | Where-Object { $_.category -eq 'traceability' })
    $sectionCoverage = @()
    foreach ($requiredSection in $requiredSections) {
        $coveredSlides = @($requiredSlideIndexes[$requiredSection] | ForEach-Object { $slides[$_] })
        $sourceBackedSlides = @($coveredSlides | Where-Object { (Get-PropertyValue $_ 'evidence_status' '') -eq 'source-backed' })
        $sectionCoverage += [pscustomobject]@{
            section = $requiredSection
            slide_count = $coveredSlides.Count
            source_backed_slide_count = $sourceBackedSlides.Count
            passed = ($coveredSlides.Count -gt 0 -and $sourceBackedSlides.Count -eq $coveredSlides.Count)
        }
    }
    return [pscustomobject]@{
        pass = ($hardFindings.Count -eq 0)
        slide_count = $slides.Count
        required_sections = $requiredSections
        findings = $findings
        checks = [pscustomobject]@{
            required_sections = if ($requiredFindings.Count) { 'failed' } else { 'passed' }
            traceability = if ($traceabilityFindings.Count) { 'failed' } else { 'passed' }
            editable_content_contract = 'pending-powerpoint-inspection'
            visual_rendering = 'pending-powerpoint-export'
        }
        quality = [pscustomobject]@{
            safe_to_generate = ($hardFindings.Count -eq 0)
            hard_finding_count = $hardFindings.Count
            warning_finding_count = $warningFindings.Count
            required_section_coverage = $sectionCoverage
            experimental_data_slide_count = @($slides | Where-Object { (Get-PropertyValue $_ 'section' '') -eq 'experimental_data' }).Count
        }
    }
}

function Release-ComReference {
    param($Reference)
    if ($null -eq $Reference) { return }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($Reference) | Out-Null } catch { }
}

function Resolve-PreviewDirectory {
    param([string]$RequestedDirectory, [string]$DefaultStem)
    if ([string]::IsNullOrWhiteSpace($RequestedDirectory)) {
        $defaultDirectory = Join-Path (Get-PaperToJournalClubTemporaryRoot) "$DefaultStem-$([Guid]::NewGuid().ToString('N'))"
        return Assert-PaperToJournalClubAllowedPath -Path $defaultDirectory -AllowedRoots @((Get-PaperToJournalClubTemporaryRoot)) -ParameterName 'default preview_directory'
    }
    $absoluteDirectory = Assert-PaperToJournalClubAllowedPath -Path $RequestedDirectory -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'preview_directory'
    if (-not (Test-DirectoryIsEmptyOrMissing $absoluteDirectory)) {
        throw "preview_directory must be a new or empty directory: $absoluteDirectory"
    }
    return $absoluteDirectory
}

function Get-DeckReferencedFigureAssets {
    param($Deck)

    # 这里只用于预检导出是否需要创建目录。最终真正复制哪一张图，仍以生成器报告的
    # inserted_figure_assets 为准，避免把 deck spec 中未成功插入的候选图长期保存下来。
    $assetDirectory = Get-SafeEvidenceAssetDirectory $Deck.evidence_pack
    if ([string]::IsNullOrWhiteSpace($assetDirectory)) { return @() }
    $images = @()
    foreach ($slide in @(Get-PropertyValue $Deck 'slides' @())) {
        $imagePath = Get-PropertyValue $slide 'suggested_image_path'
        if ([string]::IsNullOrWhiteSpace([string]$imagePath)) { continue }
        $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath ([string]$imagePath) -AllowedRoots @($assetDirectory) -ParameterName 'referenced figure asset'
        if (@($images | Where-Object { $_.path.Equals($image.path, [StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) {
            $images += $image
        }
    }
    return @($images)
}

function New-FigureAssetExportPlan {
    param([string]$PptxPath)

    # 长期保留的原图目录只能由确定的 PPTX 路径派生：<PPT名>_assets\images。
    # 不接受用户给出的任意导出目录，避免把这一可选功能变成本地批量写入能力。
    $pptxDirectory = Split-Path -Parent $PptxPath
    $pptxStem = [IO.Path]::GetFileNameWithoutExtension($PptxPath)
    if ([string]::IsNullOrWhiteSpace($pptxDirectory) -or [string]::IsNullOrWhiteSpace($pptxStem)) {
        throw 'Could not derive the figure-asset export directory from output_path.'
    }
    $writeRoots = Get-PaperToJournalClubApprovedWriteRoots
    $assetRoot = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $pptxDirectory "$pptxStem`_assets") -AllowedRoots $writeRoots -ParameterName 'figure asset export directory'
    $imagesDirectory = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $assetRoot 'images') -AllowedRoots $writeRoots -ParameterName 'figure asset export directory'
    # 即使 overwrite=true 也绝不覆盖、合并或删除既有图片资产目录；用户必须显式换 PPTX
    # 名称或自行整理旧资产，避免把长期保存的原图误删。
    if (Test-Path -LiteralPath $assetRoot) {
        throw "Figure asset export root already exists and will not be overwritten: $assetRoot"
    }
    $stagingRoot = Assert-PaperToJournalClubAllowedPath -Path (Join-Path $pptxDirectory ".${pptxStem}_assets.incoming-$([Guid]::NewGuid().ToString('N'))") -AllowedRoots $writeRoots -ParameterName 'temporary figure asset export directory'
    return [pscustomobject]@{
        asset_root = $assetRoot
        images_directory = $imagesDirectory
        staging_root = $stagingRoot
        staging_images_directory = Join-Path $stagingRoot 'images'
    }
}

function Get-ValidatedInsertedFigureAssets {
    param($Deck, $InsertedAssets)

    $assetDirectory = Get-SafeEvidenceAssetDirectory $Deck.evidence_pack
    if ($null -eq $InsertedAssets) { $InsertedAssets = @() }
    # PowerShell 5.1 对单元素 JSON array 与 PSCustomObject 的管道展开存在差异；在进入
    # foreach 前把它固定成对象数组，避免把 source_path 等属性值当成多条插图记录。
    $insertedItems = @($InsertedAssets | ForEach-Object { $_ })
    if ([string]::IsNullOrWhiteSpace($assetDirectory)) {
        if ($insertedItems.Count) { throw 'PowerPoint reported inserted figures but deck_spec has no approved extracted asset directory.' }
        return @()
    }

    # 确认生成器回传的路径确实也是该 deck 曾请求插入的图片，随后再次进行 PNG/JPEG
    # 校验。即使子进程输出被意外污染，也不能借此复制临时目录中的其它文件。
    $requestedPaths = @{}
    foreach ($image in @(Get-DeckReferencedFigureAssets -Deck $Deck)) {
        $requestedPaths[[IO.Path]::GetFullPath($image.path)] = $true
    }
    $byPath = @{}
    foreach ($entry in $insertedItems) {
        $sourcePath = Get-PropertyValue $entry 'source_path'
        if ([string]::IsNullOrWhiteSpace([string]$sourcePath)) { throw 'PowerPoint generator returned an inserted figure without source_path.' }
        $absoluteSourcePath = [IO.Path]::GetFullPath([string]$sourcePath)
        if (-not $requestedPaths.ContainsKey($absoluteSourcePath)) {
            throw "PowerPoint generator reported a figure that was not requested by deck_spec: $absoluteSourcePath"
        }
        $image = Get-PaperToJournalClubApprovedRasterImage -ImagePath $absoluteSourcePath -AllowedRoots @($assetDirectory) -ParameterName 'inserted figure asset'
        if (-not $byPath.ContainsKey($image.path)) {
            $byPath[$image.path] = [pscustomobject]@{
                image = $image
                figure_ids = @()
                slide_ids = @()
                slide_numbers = @()
            }
        }
        $record = $byPath[$image.path]
        foreach ($figureId in @(Get-PropertyValue $entry 'figure_ids' @()) + @(Get-PropertyValue $entry 'figure_id')) {
            $text = ([string]$figureId).Trim()
            if ($text -and $text -notin $record.figure_ids) { $record.figure_ids += $text }
        }
        $slideId = ([string](Get-PropertyValue $entry 'slide_id' '')).Trim()
        if ($slideId -and $slideId -notin $record.slide_ids) { $record.slide_ids += $slideId }
        $slideNumber = Get-PropertyValue $entry 'slide_number'
        if ($null -ne $slideNumber) {
            try {
                $number = [int]$slideNumber
                if ($number -ge 1 -and $number -notin $record.slide_numbers) { $record.slide_numbers += $number }
            } catch { throw 'PowerPoint generator returned an invalid inserted figure slide_number.' }
        }
    }
    return @($byPath.Values)
}

function Export-InsertedFigureAssets {
    param($Deck, $InsertedAssets, $ExportPlan)

    $assets = @(Get-ValidatedInsertedFigureAssets -Deck $Deck -InsertedAssets $InsertedAssets)
    # 用户明确请求导出、但本次没有真正插入论文图时，不创建空的 _assets 目录。
    if ($assets.Count -eq 0) {
        return [pscustomobject]@{
            requested = $true
            figure_assets_directory = $null
            figure_asset_paths = @()
            figure_assets_exported = $false
            exported_count = 0
            total_bytes = [int64]0
            assets = @()
            note = 'No approved paper raster was inserted, so no figure-asset directory was created.'
        }
    }
    if ($null -eq $ExportPlan) {
        throw 'Approved figures were inserted, but no figure-asset export plan was prepared.'
    }
    if ($assets.Count -gt 30) { throw 'The figure-asset export exceeds the 30-image safety limit.' }
    [int64]$totalSourceBytes = @($assets | ForEach-Object { [int64]$_.image.bytes } | Measure-Object -Sum).Sum
    if ($totalSourceBytes -gt 50MB) { throw 'The figure-asset export exceeds the 50 MB safety limit.' }

    $stagingCreated = $false
    try {
        # CreateNew/Directory.Move 让导出目录以同卷 staging 形式发布；中途失败时不会产生
        # 半成品正式目录，也永不覆盖现有 <PPT名>_assets。
        New-Item -ItemType Directory -Path $ExportPlan.staging_root -ErrorAction Stop | Out-Null
        $stagingCreated = $true
        [void](Assert-PaperToJournalClubAllowedPath -Path $ExportPlan.staging_root -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'temporary figure asset export directory')
        New-Item -ItemType Directory -Path $ExportPlan.staging_images_directory -ErrorAction Stop | Out-Null
        [void](Assert-PaperToJournalClubAllowedPath -Path $ExportPlan.staging_images_directory -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'temporary figure asset export directory')

        $exportedAssets = @()
        $index = 0
        foreach ($assetRecord in $assets) {
            $index++
            $sourceImage = $assetRecord.image
            $extension = [IO.Path]::GetExtension($sourceImage.path).ToLowerInvariant()
            $targetName = 'paper-figure-{0:D2}{1}' -f $index, $extension
            $stagingPath = Join-Path $ExportPlan.staging_images_directory $targetName
            $sourceStream = $null
            $targetStream = $null
            try {
                # 再次执行完整源图片校验，避免临时目录中的文件在前一次校验后被同用户
                # 进程替换为链接、错误格式或不同图像。
                $sourceImage = Get-PaperToJournalClubApprovedRasterImage -ImagePath $sourceImage.path -AllowedRoots @((Get-SafeEvidenceAssetDirectory $Deck.evidence_pack)) -ParameterName 'inserted figure asset before export'
                $sourceStream = [IO.File]::Open($sourceImage.path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
                $targetStream = [IO.File]::Open($stagingPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
                $sourceStream.CopyTo($targetStream)
            } finally {
                if ($targetStream) { $targetStream.Dispose() }
                if ($sourceStream) { $sourceStream.Dispose() }
            }
            $copiedImage = Get-PaperToJournalClubApprovedRasterImage -ImagePath $stagingPath -AllowedRoots @($ExportPlan.staging_images_directory) -ParameterName 'exported figure asset'
            $sourceHash = (Get-FileHash -LiteralPath $sourceImage.path -Algorithm SHA256).Hash
            $targetHash = (Get-FileHash -LiteralPath $copiedImage.path -Algorithm SHA256).Hash
            if ($sourceImage.bytes -ne $copiedImage.bytes -or -not $sourceHash.Equals($targetHash, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Exported figure asset integrity verification failed: $stagingPath"
            }
            $exportedAssets += [pscustomobject]@{
                path = Join-Path $ExportPlan.images_directory $targetName
                source_path = $sourceImage.path
                sha256 = $targetHash
                source_sha256 = $sourceHash
                bytes = [int64]$copiedImage.bytes
                figure_ids = @($assetRecord.figure_ids)
                slide_ids = @($assetRecord.slide_ids)
                slide_numbers = @($assetRecord.slide_numbers)
            }
        }

        # Directory.Move 要求最终根不存在；再次检查防止并发进程在预检后创建同名目录。
        [void](Assert-PaperToJournalClubAllowedPath -Path $ExportPlan.asset_root -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'figure asset export directory')
        if (Test-Path -LiteralPath $ExportPlan.asset_root) {
            throw "Figure asset export root appeared during export and will not be overwritten: $($ExportPlan.asset_root)"
        }
        [IO.Directory]::Move($ExportPlan.staging_root, $ExportPlan.asset_root)
        $stagingCreated = $false
        [void](Assert-PaperToJournalClubAllowedPath -Path $ExportPlan.images_directory -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'published figure asset directory')
        return [pscustomobject]@{
            requested = $true
            figure_assets_directory = $ExportPlan.images_directory
            figure_asset_paths = @($exportedAssets | ForEach-Object { $_.path })
            figure_assets_exported = $true
            exported_count = $exportedAssets.Count
            total_bytes = [int64](@($exportedAssets | ForEach-Object { [int64]$_.bytes } | Measure-Object -Sum).Sum)
            assets = @($exportedAssets)
            note = $null
        }
    } catch {
        # 仅删除我们刚创建、且仍位于批准目录内的 staging；正式根若已被其它进程创建，
        # 一律不触碰。PPTX 已保存时会保留，错误会清楚报告图片资产未发布。
        if ($stagingCreated -and (Test-Path -LiteralPath $ExportPlan.staging_root)) {
            try {
                $safeStaging = Assert-PaperToJournalClubAllowedPath -Path $ExportPlan.staging_root -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'temporary figure asset export directory'
                foreach ($item in @(Get-ChildItem -LiteralPath $safeStaging -Force -Recurse)) {
                    Assert-PaperToJournalClubNoReparsePoint -Path $item.FullName -ParameterName 'temporary figure asset export file'
                }
                Remove-Item -LiteralPath $safeStaging -Recurse -Force
            } catch { }
        }
        throw
    }
}

function New-TemporaryDeckSpecFile {
    <#
      生成器必须读取 JSON deck spec，但默认不应把内部工作文件留在用户的 PPTX
      输出目录。这里创建插件专用临时目录，并在生成结束后由调用方清理。
    #>
    $temporaryRoot = Get-PaperToJournalClubTemporaryRoot
    $directory = Join-Path $temporaryRoot "deck-spec-work-$([Guid]::NewGuid().ToString('N'))"
    $directory = Assert-PaperToJournalClubAllowedPath -Path $directory -AllowedRoots @($temporaryRoot) -ParameterName 'temporary deck spec directory'
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
    return [pscustomobject]@{
        directory = $directory
        path = Join-Path $directory 'deck-spec.json'
    }
}

function Remove-TemporaryDeckSpecDirectory {
    param([string]$Directory)

    if ([string]::IsNullOrWhiteSpace($Directory)) { return }
    $temporaryRoot = Get-PaperToJournalClubTemporaryRoot
    # 删除前重新验证位置和 reparse point，避免临时目录被替换后越界删除。
    $target = Assert-PaperToJournalClubAllowedPath -Path $Directory -AllowedRoots @($temporaryRoot) -ParameterName 'temporary deck spec directory'
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}

function Get-PowerPointStatus {
    $registered = Test-PowerPointComRegistration
    $application = $null
    $presentation = $null
    try {
        $application = Get-ActivePowerPointApplication
        if ($application) {
            try { $presentation = $application.ActivePresentation } catch { $presentation = $null }
            return [pscustomobject]@{
                target_application = 'Microsoft PowerPoint'
                operating_system = 'Windows'
                powerpoint_com_registered = $registered
                active_process_detected = $true
                connected_current_window = ($null -ne $presentation)
                connection_scope = if ($presentation) { 'current-window-inspection' } else { 'application-without-active-presentation' }
                active_presentation = if ($presentation) { Get-PowerPointPresentationSummary $presentation } else { $null }
                capabilities = [pscustomobject]@{
                    inspect_current_presentation = ($null -ne $presentation)
                    create_new_editable_presentation = $registered
                    edit_current_presentation = $false
                    export_native_previews = $registered
                }
                generation_scope = 'new-background-presentation'
                note = 'This plugin can inspect the active PowerPoint window through COM. Generation creates a new editable presentation and does not modify the current one.'
            }
        }
        return [pscustomobject]@{
            target_application = 'Microsoft PowerPoint'
            operating_system = 'Windows'
            powerpoint_com_registered = $registered
            active_process_detected = $false
            connected_current_window = $false
            connection_scope = 'no-active-powerpoint-window'
            active_presentation = $null
            capabilities = [pscustomobject]@{
                inspect_current_presentation = $false
                create_new_editable_presentation = $registered
                edit_current_presentation = $false
                export_native_previews = $registered
            }
            generation_scope = 'new-background-presentation'
            note = if ($registered) { 'PowerPoint is installed but no active PowerPoint window is available.' } else { 'Microsoft PowerPoint desktop is not registered on this computer.' }
        }
    } finally {
        Release-ComReference $presentation
        Release-ComReference $application
    }
}

function Invoke-InspectPowerPoint {
    param($Arguments)
    if (-not (Test-PowerPointComRegistration)) { throw 'Microsoft PowerPoint desktop is not registered on this computer.' }
    $filePath = Get-PropertyValue $Arguments 'file_path'
    $exportPreviews = Get-StrictBoolean -Object $Arguments -Name 'export_previews' -Default $false
    if ($filePath) {
        $absolutePath = Get-ExistingPowerPointPath $filePath
        $session = $null
        try {
            $session = Open-PowerPointPresentationReadOnly -Path $absolutePath -WithWindow:$exportPreviews
            $previewDirectory = if ($exportPreviews) { Resolve-PreviewDirectory (Get-PropertyValue $Arguments 'preview_directory') 'file-previews' } else { $null }
            $quality = Invoke-PowerPointQualityAudit -Presentation $session.presentation -PreviewDirectory $previewDirectory -ExportPreviews:$exportPreviews
            return [pscustomobject]@{
                connection_scope = 'file-read-only'
                connected_current_window = $false
                working_copy = $false
                presentation_path = $absolutePath
                quality_audit = $quality
            }
        } finally {
            Close-PowerPointReadOnlySession $session
        }
    }

    $application = $null
    $presentation = $null
    try {
        $application = Get-ActivePowerPointApplication
        if (-not $application) { throw 'No active PowerPoint application is available. Pass file_path to inspect a saved PPTX read-only.' }
        try { $presentation = $application.ActivePresentation } catch { $presentation = $null }
        if (-not $presentation) { throw 'PowerPoint is open but has no active presentation. Pass file_path or create a presentation first.' }
        $previewDirectory = if ($exportPreviews) { Resolve-PreviewDirectory (Get-PropertyValue $Arguments 'preview_directory') 'current-window-previews' } else { $null }
        $quality = Invoke-PowerPointQualityAudit -Presentation $presentation -PreviewDirectory $previewDirectory -ExportPreviews:$exportPreviews
        return [pscustomobject]@{
            connection_scope = 'current-window-inspection'
            connected_current_window = $true
            working_copy = $false
            presentation_path = $presentation.FullName
            quality_audit = $quality
        }
    } finally {
        # 这里只释放 COM 引用，绝不关闭用户正在使用的 PowerPoint 或演示文稿。
        Release-ComReference $presentation
        Release-ComReference $application
    }
}

function Invoke-AuditEditablePptx {
    param($Arguments)
    $filePath = Get-PropertyValue $Arguments 'file_path'
    if (-not $filePath) { throw 'file_path is required for a saved-PPTX audit.' }
    $inspectionArguments = [pscustomobject]@{
        file_path = $filePath
        # PNG 预览用于人工复核，默认不向用户目录额外写入文件。
        export_previews = Get-StrictBoolean -Object $Arguments -Name 'export_previews' -Default $false
        preview_directory = Get-PropertyValue $Arguments 'preview_directory'
    }
    return Invoke-InspectPowerPoint $inspectionArguments
}

function Test-RetryableSavedPptxAuditState {
    param($QualityAudit)

    if ($null -eq $QualityAudit) { return $true }
    $slideCount = 0
    try { $slideCount = [int](Get-PropertyValue (Get-PropertyValue $QualityAudit 'presentation') 'slide_count' 0) } catch { }
    if ($slideCount -lt 1) { return $true }
    $findings = @(Get-PropertyValue $QualityAudit 'findings' @())
    if ($findings.Count -eq 0) { return $false }
    # 只重试没有任何具体 slide/shape 的 Office collection 同步错误；任何具体对象上的版式、
    # 溢出或图片错误都应立即保留给用户修正，不能被重试吞掉。
    $anonymousFailures = @($findings | Where-Object {
        (Get-PropertyValue $_ 'severity') -eq 'hard' -and
        (Get-PropertyValue $_ 'category') -in @('slides', 'geometry', 'editability') -and
        $null -eq (Get-PropertyValue $_ 'slide') -and
        $null -eq (Get-PropertyValue $_ 'shape')
    })
    return $anonymousFailures.Count -gt 0 -and $anonymousFailures.Count -eq $findings.Count
}

function Invoke-SavedPptxAuditWithRetry {
    param([string]$Path, [bool]$ExportPreviews, [string]$PreviewDirectory)

    $lastResult = $null
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        $lastResult = Invoke-AuditEditablePptx ([pscustomobject]@{
            file_path = $Path
            export_previews = $ExportPreviews
            preview_directory = $PreviewDirectory
        })
        $quality = Get-PropertyValue $lastResult 'quality_audit'
        if (-not (Test-RetryableSavedPptxAuditState -QualityAudit $quality) -or $attempt -eq 4) {
            return $lastResult
        }
        Start-Sleep -Milliseconds 750
    }
    return $lastResult
}

function Invoke-GeneratePptx {
    param($Arguments)
    $deck = Get-PropertyValue $Arguments 'deck_spec'
    $outputPath = Get-PropertyValue $Arguments 'output_path'
    if (-not $deck -or -not $outputPath) { throw 'deck_spec and output_path are required.' }
    Assert-DeckSpecificationLimits -Deck $deck
    # 不能仅依赖调用方先执行 audit：直接调用生成工具也必须经过必备模块与证据门禁。
    $audit = Invoke-AuditDeck ([pscustomobject]@{ deck_spec = $deck })
    if (-not $audit.pass) {
        $blockingSummary = @($audit.findings | Where-Object { $_.severity -eq 'hard' } | Select-Object -First 5 | ForEach-Object { $_.issue }) -join ' | '
        throw "Deck failed the mandatory content audit. Resolve hard findings before generating PowerPoint: $blockingSummary"
    }
    $overwrite = Get-StrictBoolean -Object $Arguments -Name 'overwrite' -Default $false
    $absoluteOutput = Resolve-RequestedOutputPath -Path $outputPath -RequiredExtension '.pptx' -Overwrite $overwrite
    # 仅在调用方明确请求时长期保留本次实际插入 PPT 的论文原图；不接受额外目录参数，
    # 目标固定为 <PPT名>_assets\images，避免模型将本机任意位置变成批量图片输出点。
    $exportFigureAssets = Get-StrictBoolean -Object $Arguments -Name 'export_figure_assets' -Default $false
    # 在启动 PowerPoint 前预检最终目录。若同名长期资产已存在，直接失败而不是先生成
    # PPTX 再发现无法安全发布图片副本。
    $requestedFigureAssets = if ($exportFigureAssets) { @(Get-DeckReferencedFigureAssets -Deck $deck) } else { @() }
    $figureAssetExportPlan = if ($requestedFigureAssets.Count -gt 0) { New-FigureAssetExportPlan -PptxPath $absoluteOutput } else { $null }
    $requestedDeckSpecPath = Get-PropertyValue $Arguments 'deck_spec_output_path'
    $temporaryDeckSpec = $null
    $persistentDeckSpecPath = $null
    $generationResult = $null
    $temporaryCleanupWarning = $null
    if ([string]::IsNullOrWhiteSpace($requestedDeckSpecPath)) {
        # 未明确请求导出时，deck spec 只作为生成器的短生命周期工作文件。
        $temporaryDeckSpec = New-TemporaryDeckSpecFile
        $specPath = $temporaryDeckSpec.path
    } else {
        $persistentDeckSpecPath = Resolve-RequestedOutputPath -Path $requestedDeckSpecPath -RequiredExtension '.json' -Overwrite $overwrite
        $specPath = $persistentDeckSpecPath
    }

    try {
        $deck | ConvertTo-Json -Depth 50 | Set-Content -Encoding UTF8 -LiteralPath $specPath
        $keepOpen = Get-StrictBoolean -Object $Arguments -Name 'keep_powerpoint_open' -Default $false
        # 预览图仅在调用方明确要求时导出，默认交付物只有 PPTX。
        $exportPreviews = Get-StrictBoolean -Object $Arguments -Name 'export_previews' -Default $false
        $previewDirectory = if ($exportPreviews) { Resolve-PreviewDirectory (Get-PropertyValue $Arguments 'preview_directory') 'generated-previews' } else { $null }
        $generatorArguments = @('-NoProfile', '-STA', '-ExecutionPolicy', 'RemoteSigned', '-File', $PptGeneratorPath, '-DeckSpecPath', $specPath, '-OutputPath', $absoluteOutput, '-DeferQualityAudit')
        if ($keepOpen) { $generatorArguments += '-KeepOpen' }
        if ($overwrite) { $generatorArguments += '-Overwrite' }
        $generatorOutput = @(& powershell.exe @generatorArguments 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "PowerPoint generation failed: $($generatorOutput -join [Environment]::NewLine)" }
        $jsonLine = @($generatorOutput | ForEach-Object { [string]$_ } | Where-Object { $_.TrimStart().StartsWith('{') } | Select-Object -Last 1)
        if ($jsonLine.Count -ne 1) { throw 'PowerPoint generator completed without its required JSON quality report.' }
        try {
            $nativeResult = $jsonLine[0] | ConvertFrom-Json
        } catch {
            throw "PowerPoint generator returned an invalid JSON quality report. $($_.Exception.Message)"
        }
        $nativeQuality = Get-PropertyValue $nativeResult 'quality_audit'
        if (Get-StrictBoolean -Object $nativeResult -Name 'quality_audit_deferred' -Default $false) {
            $savedFileAudit = Invoke-SavedPptxAuditWithRetry -Path $absoluteOutput -ExportPreviews $exportPreviews -PreviewDirectory $previewDirectory
            $nativeQuality = Get-PropertyValue $savedFileAudit 'quality_audit'
        }
        if (-not $nativeQuality -or -not (Get-StrictBoolean -Object $nativeQuality -Name 'pass' -Default $false)) {
            $findings = if ($nativeQuality) { Get-PropertyValue $nativeQuality 'findings' @() } else { @() }
            throw "PowerPoint quality audit failed for the saved PPTX: $($findings | ConvertTo-Json -Compress)"
        }
        $reportedInsertedFigureAssets = Get-PropertyValue $nativeResult 'inserted_figure_assets' @()
        $figureAssetExport = if ($exportFigureAssets) {
            Export-InsertedFigureAssets -Deck $deck -InsertedAssets $reportedInsertedFigureAssets -ExportPlan $figureAssetExportPlan
        } else {
            [pscustomobject]@{
                requested = $false
                figure_assets_directory = $null
                figure_asset_paths = @()
                figure_assets_exported = $false
                exported_count = 0
                total_bytes = [int64]0
                assets = @()
                note = 'Figure-asset export was not requested.'
            }
        }
        $generationResult = [pscustomobject]@{
            output_path = $absoluteOutput
            deck_spec_path = $persistentDeckSpecPath
            deck_spec_saved = ($null -ne $persistentDeckSpecPath)
            target_application = 'Microsoft PowerPoint'
            connection_scope = 'new-background-presentation'
            editable_contract = 'native-text-and-shapes'
            content_audit = $audit
            quality_audit = $nativeQuality
            preview_directory = if ($exportPreviews) { $previewDirectory } else { $null }
            preview_paths = Get-PropertyValue $nativeQuality 'preview_paths' @()
            figure_assets_directory = $figureAssetExport.figure_assets_directory
            figure_asset_paths = @($figureAssetExport.figure_asset_paths)
            figure_assets_exported = [bool]$figureAssetExport.figure_assets_exported
            figure_asset_export = $figureAssetExport
            completion_message = '感谢使用 Paper to Journal Club 插件，制作者：jiuBABY6。'
            next_step = if ($exportPreviews -and $figureAssetExport.figure_assets_exported) {
                'Open the generated PPTX in PowerPoint and review the returned PNG previews, exported figure assets, and quality audit.'
            } elseif ($exportPreviews) {
                'Open the generated PPTX in PowerPoint and review the returned PNG previews and quality audit.'
            } elseif ($figureAssetExport.figure_assets_exported) {
                'Open the generated PPTX in PowerPoint. The original raster figures inserted into the deck were also copied to the returned figure_assets_directory.'
            } else {
                'Open the generated PPTX in PowerPoint. The quality audit is returned above; PNG previews, exported figure assets, and a deck-spec file are created only when explicitly requested.'
            }
        }
    } finally {
        if ($temporaryDeckSpec) {
            try {
                Remove-TemporaryDeckSpecDirectory -Directory $temporaryDeckSpec.directory
            } catch {
                # PPTX 已成功保存时，临时文件清理失败不应掩盖主结果；把它显式返回，
                # 便于用户或维护者清理专用临时根，同时保留原始错误信息。
                $temporaryCleanupWarning = "Could not remove the internal temporary deck-spec directory '$($temporaryDeckSpec.directory)': $($_.Exception.Message)"
            }
        }
    }
    if ($temporaryCleanupWarning -and $generationResult) {
        $generationResult | Add-Member -NotePropertyName 'temporary_cleanup_warning' -NotePropertyValue $temporaryCleanupWarning
    }
    return $generationResult
}

function Get-Tools {
    return @(
        [pscustomobject]@{ name = 'analyse_paper'; description = 'Parse a paper into sections, candidate claims, page evidence, and figure assets without Node.js.'; inputSchema = [ordered]@{ type = 'object'; required = @('file_path'); properties = [ordered]@{ file_path = [ordered]@{ type = 'string'; description = 'Absolute paper path.' }; asset_output_dir = [ordered]@{ type = 'string'; description = 'Optional directory for parser-extracted paper image assets. Defaults to a user temp directory.' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'cleanup_paper_assets'; description = 'Delete only a confirmed default temporary paper-asset directory created by this plugin.'; inputSchema = [ordered]@{ type = 'object'; required = @('asset_output_dir', 'confirm'); properties = [ordered]@{ asset_output_dir = [ordered]@{ type = 'string' }; confirm = [ordered]@{ type = 'boolean'; const = $true } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'render_paper_visual'; description = 'Render one reviewed figure or table source page from a PDF into a PNG candidate. A crop is optional but must be user-confirmed; rendered table and vector visuals are never inserted automatically.'; inputSchema = [ordered]@{ type = 'object'; required = @('evidence_pack', 'visual_id', 'page_number'); properties = [ordered]@{ evidence_pack = [ordered]@{ type = 'object' }; visual_id = [ordered]@{ type = 'string'; description = 'A figure or table id from the evidence pack.' }; page_number = [ordered]@{ type = 'integer'; minimum = 1; description = 'Must equal source_page when it is unambiguous, or be one of the reviewed source_pages for a repeated table reference.' }; crop = [ordered]@{ type = 'object'; description = 'Optional user-confirmed normalized crop inside the page.'; required = @('x', 'y', 'width', 'height'); properties = [ordered]@{ x = [ordered]@{ type = 'number'; minimum = 0; maximum = 1 }; y = [ordered]@{ type = 'number'; minimum = 0; maximum = 1 }; width = [ordered]@{ type = 'number'; exclusiveMinimum = 0; maximum = 1 }; height = [ordered]@{ type = 'number'; exclusiveMinimum = 0; maximum = 1 } }; additionalProperties = $false }; width = [ordered]@{ type = 'integer'; minimum = 400; maximum = 3200; default = 1600 }; height = [ordered]@{ type = 'integer'; minimum = 400; maximum = 3200; default = 2200 } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'design_journal_club_deck'; description = 'Create an evidence-backed journal-club deck specification with required content sections and a coherent evidence-first narrative.'; inputSchema = [ordered]@{ type = 'object'; required = @('evidence_pack'); properties = [ordered]@{ evidence_pack = [ordered]@{ type = 'object' }; duration_minutes = [ordered]@{ type = 'integer'; minimum = 5; maximum = 90; default = 15 }; language = [ordered]@{ type = 'string'; default = 'zh-CN' }; audience = [ordered]@{ type = 'string'; enum = @('lab', 'mixed', 'expert'); default = 'lab' }; required_sections = [ordered]@{ type = 'array'; minItems = 1; uniqueItems = $true; items = [ordered]@{ type = 'string'; enum = $KnownJournalClubSections }; default = $DefaultRequiredSections; description = 'Defaults to background, innovation, methods, experimental_data, limitations, and future_directions.' }; figure_asset_selection = [ordered]@{ type = 'object'; description = 'Optional reviewed mapping from evidence-pack figure or table id to an extracted image path. Automatic insertion is limited to one recognized figure and one approved raster asset on the same page. Page-rendered table crops always require explicit review and selection.'; additionalProperties = [ordered]@{ type = 'string' } }; output_path = [ordered]@{ type = 'string'; description = 'Optional new .json deck-spec path.' }; overwrite = [ordered]@{ type = 'boolean'; default = $false } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'audit_journal_club_deck'; description = 'Hard-fail missing required sections and audit evidence traceability before PowerPoint generation.'; inputSchema = [ordered]@{ type = 'object'; required = @('deck_spec'); properties = [ordered]@{ deck_spec = [ordered]@{ type = 'object' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'powerpoint_status'; description = 'Inspect Microsoft PowerPoint COM availability and the current window without modifying it.'; inputSchema = [ordered]@{ type = 'object'; properties = [ordered]@{}; additionalProperties = $false } },
        [pscustomobject]@{ name = 'inspect_powerpoint'; description = 'Inspect the active PowerPoint window through COM, or a saved PPTX in read-only mode. It never edits the presentation.'; inputSchema = [ordered]@{ type = 'object'; properties = [ordered]@{ file_path = [ordered]@{ type = 'string'; description = 'Optional saved .pptx path. Omit only to inspect the active current window.' }; export_previews = [ordered]@{ type = 'boolean'; default = $false }; preview_directory = [ordered]@{ type = 'string'; description = 'New or empty folder for native PowerPoint PNG previews.' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'audit_editable_pptx'; description = 'Open a saved PPTX read-only and verify native editable objects. Native PowerPoint PNG previews are opt-in.'; inputSchema = [ordered]@{ type = 'object'; required = @('file_path'); properties = [ordered]@{ file_path = [ordered]@{ type = 'string' }; export_previews = [ordered]@{ type = 'boolean'; default = $false }; preview_directory = [ordered]@{ type = 'string' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'generate_editable_pptx'; description = 'Create a new editable PowerPoint presentation after the required-section audit passes. By default it writes only the PPTX; PNG previews, copied source figure assets, and a persistent deck-spec JSON require explicit opt-in.'; inputSchema = [ordered]@{ type = 'object'; required = @('deck_spec', 'output_path'); properties = [ordered]@{ deck_spec = [ordered]@{ type = 'object' }; output_path = [ordered]@{ type = 'string'; description = 'New .pptx path.' }; overwrite = [ordered]@{ type = 'boolean'; default = $false }; keep_powerpoint_open = [ordered]@{ type = 'boolean'; default = $false }; deck_spec_output_path = [ordered]@{ type = 'string'; description = 'Optional explicit .json path for a persistent deck-spec export. Omit to keep the generator work file temporary.' }; export_previews = [ordered]@{ type = 'boolean'; default = $false }; preview_directory = [ordered]@{ type = 'string'; description = 'New or empty folder for native PowerPoint PNG previews. Requires export_previews=true.' }; export_figure_assets = [ordered]@{ type = 'boolean'; default = $false; description = 'When true, copy only the approved paper raster images actually inserted in the PPTX to <PPT-name>_assets\\images. The derived asset root must not already exist.' } }; additionalProperties = $false } }
    )
}

function Select-McpProtocolVersion {
    param([string]$RequestedVersion)
    if ($RequestedVersion -in $SupportedProtocolVersions) { return $RequestedVersion }
    # 对未知客户端版本使用已测试的最新协议；客户端仍可按 MCP 规范决定是否继续握手。
    return '2025-06-18'
}

function Write-McpResponse {
    param($Id, $Result, $ErrorMessage = $null)
    $response = if ($ErrorMessage) { [ordered]@{ jsonrpc = '2.0'; id = $Id; error = [ordered]@{ code = -32000; message = $ErrorMessage } } } else { [ordered]@{ jsonrpc = '2.0'; id = $Id; result = $Result } }
    [Console]::Out.WriteLine(($response | ConvertTo-Json -Depth 80 -Compress))
}

function Assert-FigureAssetSelectionArgument {
    param($Arguments)
    if (-not (Test-PropertyExists $Arguments 'figure_asset_selection')) { return }
    $selection = Get-PropertyValue $Arguments 'figure_asset_selection'
    Assert-McpObject -Value $selection -ParameterName 'figure_asset_selection'
    $names = @(Get-McpObjectPropertyNames $selection)
    if ($names.Count -gt 30) { throw 'figure_asset_selection may contain at most 30 visual mappings.' }
    foreach ($figureId in $names) {
        if ($figureId -notmatch '^(?:fig|table)-\d+[a-z]?$') { throw "figure_asset_selection key is not a supported figure or table id: $figureId" }
        $path = Get-PropertyValue $selection $figureId
        if ($path -isnot [string] -or [string]::IsNullOrWhiteSpace($path) -or $path.Length -gt 4096) {
            throw "figure_asset_selection.$figureId must be a non-empty image path no longer than 4096 characters."
        }
    }
}

function Assert-RenderPaperVisualArgument {
    param($Arguments)

    Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('evidence_pack', 'visual_id', 'page_number', 'crop', 'width', 'height') -ToolName 'render_paper_visual'
    if (-not (Test-PropertyExists $Arguments 'evidence_pack')) { throw 'evidence_pack is required.' }
    Assert-McpObject -Value (Get-PropertyValue $Arguments 'evidence_pack') -ParameterName 'evidence_pack'
    Assert-StringArgument -Arguments $Arguments -Name 'visual_id' -Required $true -MaximumLength 64
    if (-not (Test-PropertyExists $Arguments 'page_number')) { throw 'page_number is required.' }
    [void](Assert-IntegerArgument -Arguments $Arguments -Name 'page_number' -Minimum 1 -Maximum 10000 -Default 1)
    [void](Assert-IntegerArgument -Arguments $Arguments -Name 'width' -Minimum 400 -Maximum 3200 -Default 1600)
    [void](Assert-IntegerArgument -Arguments $Arguments -Name 'height' -Minimum 400 -Maximum 3200 -Default 2200)
    if (Test-PropertyExists $Arguments 'crop') {
        $crop = Get-PropertyValue $Arguments 'crop'
        Assert-McpObject -Value $crop -ParameterName 'crop'
        if (@(Get-McpObjectPropertyNames $crop | Where-Object { $_ -notin @('x', 'y', 'width', 'height') }).Count) { throw 'crop contains an unsupported property.' }
        foreach ($name in @('x', 'y', 'width', 'height')) {
            if (-not (Test-PropertyExists $crop $name)) { throw "crop.$name is required." }
            $value = Get-PropertyValue $crop $name
            if ($value -isnot [double] -and $value -isnot [decimal] -and $value -isnot [int] -and $value -isnot [int64]) { throw "crop.$name must be a number." }
        }
    }
}

function Assert-McpToolArguments {
    param([string]$Name, $Arguments)

    Assert-McpObject -Value $Arguments -ParameterName 'tools/call arguments'
    $state = [pscustomobject]@{ nodes = 0; characters = 0 }
    Assert-McpValueLimits -Value $Arguments -State ([ref]$state)

    switch ($Name) {
        'analyse_paper' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('file_path', 'asset_output_dir') -ToolName $Name
            Assert-StringArgument -Arguments $Arguments -Name 'file_path' -Required $true
            Assert-StringArgument -Arguments $Arguments -Name 'asset_output_dir'
        }
        'cleanup_paper_assets' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('asset_output_dir', 'confirm') -ToolName $Name
            Assert-StringArgument -Arguments $Arguments -Name 'asset_output_dir' -Required $true
            if (-not (Get-StrictBoolean -Object $Arguments -Name 'confirm' -Default $false)) { throw 'Set confirm=true before deleting extracted temporary paper assets.' }
        }
        'render_paper_visual' {
            Assert-RenderPaperVisualArgument -Arguments $Arguments
        }
        'design_journal_club_deck' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('evidence_pack', 'duration_minutes', 'language', 'audience', 'required_sections', 'figure_asset_selection', 'output_path', 'overwrite') -ToolName $Name
            if (-not (Test-PropertyExists $Arguments 'evidence_pack')) { throw 'evidence_pack is required.' }
            Assert-McpObject -Value (Get-PropertyValue $Arguments 'evidence_pack') -ParameterName 'evidence_pack'
            [void](Assert-IntegerArgument -Arguments $Arguments -Name 'duration_minutes' -Minimum 5 -Maximum 90 -Default 15)
            Assert-StringArgument -Arguments $Arguments -Name 'language' -MaximumLength 32
            if (Test-PropertyExists $Arguments 'audience') {
                Assert-StringArgument -Arguments $Arguments -Name 'audience' -MaximumLength 16
                if ((Get-PropertyValue $Arguments 'audience') -notin @('lab', 'mixed', 'expert')) { throw 'audience must be one of: lab, mixed, expert.' }
            }
            Assert-StringArrayArgument -Arguments $Arguments -Name 'required_sections' -MinimumItems 1 -MaximumItems 6 -MaximumItemLength 64
            if (Test-PropertyExists $Arguments 'required_sections') {
                foreach ($section in @(Get-PropertyValue $Arguments 'required_sections')) {
                    if ($section -notin $KnownJournalClubSections) { throw "Unknown required_sections value: $section" }
                }
            }
            Assert-FigureAssetSelectionArgument -Arguments $Arguments
            Assert-StringArgument -Arguments $Arguments -Name 'output_path'
            [void](Get-StrictBoolean -Object $Arguments -Name 'overwrite' -Default $false)
        }
        'audit_journal_club_deck' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('deck_spec') -ToolName $Name
            if (-not (Test-PropertyExists $Arguments 'deck_spec')) { throw 'deck_spec is required.' }
            Assert-DeckSpecificationLimits -Deck (Get-PropertyValue $Arguments 'deck_spec')
        }
        'powerpoint_status' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @() -ToolName $Name
        }
        'inspect_powerpoint' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('file_path', 'export_previews', 'preview_directory') -ToolName $Name
            Assert-StringArgument -Arguments $Arguments -Name 'file_path'
            [void](Get-StrictBoolean -Object $Arguments -Name 'export_previews' -Default $false)
            Assert-StringArgument -Arguments $Arguments -Name 'preview_directory'
        }
        'audit_editable_pptx' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('file_path', 'export_previews', 'preview_directory') -ToolName $Name
            Assert-StringArgument -Arguments $Arguments -Name 'file_path' -Required $true
            [void](Get-StrictBoolean -Object $Arguments -Name 'export_previews' -Default $false)
            Assert-StringArgument -Arguments $Arguments -Name 'preview_directory'
        }
        'generate_editable_pptx' {
            Assert-OnlyKnownArguments -Arguments $Arguments -AllowedNames @('deck_spec', 'output_path', 'overwrite', 'keep_powerpoint_open', 'deck_spec_output_path', 'export_previews', 'preview_directory', 'export_figure_assets') -ToolName $Name
            if (-not (Test-PropertyExists $Arguments 'deck_spec')) { throw 'deck_spec is required.' }
            Assert-DeckSpecificationLimits -Deck (Get-PropertyValue $Arguments 'deck_spec')
            Assert-StringArgument -Arguments $Arguments -Name 'output_path' -Required $true
            [void](Get-StrictBoolean -Object $Arguments -Name 'overwrite' -Default $false)
            [void](Get-StrictBoolean -Object $Arguments -Name 'keep_powerpoint_open' -Default $false)
            Assert-StringArgument -Arguments $Arguments -Name 'deck_spec_output_path'
            [void](Get-StrictBoolean -Object $Arguments -Name 'export_previews' -Default $false)
            Assert-StringArgument -Arguments $Arguments -Name 'preview_directory'
            [void](Get-StrictBoolean -Object $Arguments -Name 'export_figure_assets' -Default $false)
        }
        default { throw "Unknown tool: $Name" }
    }
}

function Invoke-McpTool {
    param([string]$Name, $Arguments)
    Assert-McpToolArguments -Name $Name -Arguments $Arguments
    switch ($Name) {
        'analyse_paper' { return Invoke-AnalysePaper $Arguments }
        'cleanup_paper_assets' { return Remove-TemporaryPaperAssets $Arguments }
        'render_paper_visual' { return Invoke-RenderPaperVisual $Arguments }
        'design_journal_club_deck' { return Invoke-DesignDeck $Arguments }
        'audit_journal_club_deck' { return Invoke-AuditDeck $Arguments }
        'powerpoint_status' { return Get-PowerPointStatus }
        'inspect_powerpoint' { return Invoke-InspectPowerPoint $Arguments }
        'audit_editable_pptx' { return Invoke-AuditEditablePptx $Arguments }
        'generate_editable_pptx' { return Invoke-GeneratePptx $Arguments }
        default { throw "Unknown tool: $Name" }
    }
}

function Read-McpRequestLineWithLimit {
    param([int]$MaximumCharacters)

    # Console.ReadLine 会在检查长度前把整个攻击请求分配到内存。逐字符读取并在超限后
    # 丢弃到换行符，能把单个 JSON-RPC 请求的常驻内存限制在 MaximumCharacters 以内。
    $builder = New-Object System.Text.StringBuilder
    $discarding = $false
    while ($true) {
        $next = [Console]::In.Read()
        if ($next -eq -1) {
            if ($discarding) { throw "MCP request exceeds the $MaximumCharacters-character safety limit." }
            if ($builder.Length -eq 0) { return $null }
            return $builder.ToString()
        }
        if ($next -eq 10) {
            if ($discarding) { throw "MCP request exceeds the $MaximumCharacters-character safety limit." }
            return $builder.ToString()
        }
        if ($discarding) { continue }
        if ($builder.Length -ge $MaximumCharacters) {
            $discarding = $true
            continue
        }
        [void]$builder.Append([char]$next)
    }
}

if ($Demo) {
    if (-not $DemoInputPath) { throw 'DemoInputPath is required when using -Demo.' }
    $evidence = Invoke-AnalysePaper ([pscustomobject]@{ file_path = $DemoInputPath })
    $deck = Invoke-DesignDeck ([pscustomobject]@{ evidence_pack = $evidence; duration_minutes = 15; language = 'zh-CN'; audience = 'lab'; output_path = $DemoOutputPath })
    [Console]::Out.WriteLine(([pscustomobject]@{ deck_spec = $DemoOutputPath; audit = Invoke-AuditDeck ([pscustomobject]@{ deck_spec = $deck }) } | ConvertTo-Json -Depth 80))
    exit 0
}

while ($true) {
    try {
        $line = Read-McpRequestLineWithLimit -MaximumCharacters $MaximumMcpRequestCharacters
    } catch {
        Write-McpResponse $null $null $_.Exception.Message
        continue
    }
    if ($null -eq $line) { break }
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $request = $null
    try {
        if ($line.Length -gt $MaximumMcpRequestCharacters) {
            throw "MCP request exceeds the $MaximumMcpRequestCharacters-character safety limit."
        }
        $request = $line.TrimStart([char]0xFEFF) | ConvertFrom-Json
        switch ($request.method) {
            'notifications/initialized' { continue }
            'initialize' {
                $requestedProtocol = Get-PropertyValue $request.params 'protocolVersion' '2025-06-18'
                Write-McpResponse $request.id ([ordered]@{
                    protocolVersion = Select-McpProtocolVersion $requestedProtocol
                    capabilities = [ordered]@{ tools = [ordered]@{} }
                serverInfo = [ordered]@{ name = 'paper-to-journal-club'; version = $PluginVersion }
                })
            }
            'tools/list' { Write-McpResponse $request.id ([ordered]@{ tools = Get-Tools }) }
            'tools/call' {
                $parameters = Get-PropertyValue $request 'params'
                Assert-McpObject -Value $parameters -ParameterName 'tools/call params'
                $toolName = Get-PropertyValue $parameters 'name'
                if ($toolName -isnot [string] -or [string]::IsNullOrWhiteSpace($toolName) -or $toolName.Length -gt 128) {
                    throw 'tools/call params.name must be a non-empty tool name string.'
                }
                $toolArguments = if (Test-PropertyExists $parameters 'arguments') { Get-PropertyValue $parameters 'arguments' } else { [pscustomobject]@{} }
                $toolResult = Invoke-McpTool $toolName $toolArguments
                Write-McpResponse $request.id ([ordered]@{ content = @([ordered]@{ type = 'text'; text = ($toolResult | ConvertTo-Json -Depth 80) }) })
            }
            default { Write-McpResponse $request.id $null "Unsupported method: $($request.method)" }
        }
    } catch {
        Write-McpResponse (Get-PropertyValue $request 'id') $null $_.Exception.Message
    }
}
