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
$ParserPath = Join-Path $PluginRoot "assets\paper-parser.exe"
$PptGeneratorPath = Join-Path $ScriptRoot "generate-editable-pptx.ps1"
$PowerPointQualityPath = Join-Path $ScriptRoot "powerpoint-quality.ps1"
$SupportedProtocolVersions = @('2024-11-05', '2025-03-26', '2025-06-18')
$MaximumPaperBytes = 100MB
$ParserTimeoutMilliseconds = 120000

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
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return $property.Value
}

function Get-ValidatedPaperPath {
    param([string]$FilePath)

    if ([string]::IsNullOrWhiteSpace($FilePath)) { throw 'file_path is required.' }
    $absolutePath = [IO.Path]::GetFullPath($FilePath)
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
    $absolutePath = [IO.Path]::GetFullPath($Path)
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
    $absolutePath = [IO.Path]::GetFullPath($Path)
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
    if ($RequestedDirectory) {
        $directory = [IO.Path]::GetFullPath($RequestedDirectory)
        if (-not (Test-DirectoryIsEmptyOrMissing $directory)) {
            throw "asset_output_dir must be a new or empty directory so existing assets are not overwritten: $directory"
        }
    } else {
        # 默认使用用户临时目录，避免向已安装的插件目录写入运行时数据。
        $safeStem = ([IO.Path]::GetFileNameWithoutExtension($PaperPath) -replace '[^a-zA-Z0-9._-]', '-')
        if (-not $safeStem) { $safeStem = 'paper' }
        $directory = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\$safeStem-$([Guid]::NewGuid().ToString('N'))"
    }
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
    return $directory
}

function Remove-TemporaryPaperAssets {
    param($Arguments)
    $directory = Get-PropertyValue $Arguments 'asset_output_dir'
    $confirmed = [bool](Get-PropertyValue $Arguments 'confirm' $false)
    if (-not $confirmed) { throw 'Set confirm=true before deleting extracted temporary paper assets.' }
    if ([string]::IsNullOrWhiteSpace($directory)) { throw 'asset_output_dir is required.' }

    $target = [IO.Path]::GetFullPath($directory).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $temporaryRoot = [IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'paper-to-journal-club')).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $requiredPrefix = $temporaryRoot + [IO.Path]::DirectorySeparatorChar
    if (-not $target.StartsWith($requiredPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to delete a directory outside this plugin's default temporary asset root: $temporaryRoot"
    }
    if (-not (Test-Path -LiteralPath $target)) {
        return [pscustomobject]@{ deleted = $false; asset_output_dir = $target; note = 'Directory was already absent.' }
    }
    # 只删除用户明确确认、且路径已验证在专用临时根目录内的资产目录。
    Remove-Item -LiteralPath $target -Recurse -Force
    return [pscustomobject]@{ deleted = $true; asset_output_dir = $target }
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
        $standardOutputTask = $process.StandardOutput.ReadToEndAsync()
        $standardErrorTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($ParserTimeoutMilliseconds)) {
            try { $process.Kill() } catch { }
            throw "PDF parsing exceeded the $([int]($ParserTimeoutMilliseconds / 1000))-second safety limit."
        }
        $standardOutput = $standardOutputTask.GetAwaiter().GetResult()
        $standardError = $standardErrorTask.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) {
            throw "Bundled PDF parser failed: $standardError"
        }
        if ([string]::IsNullOrWhiteSpace($standardOutput)) { throw 'Bundled PDF parser returned no JSON package.' }
        return $standardOutput | ConvertFrom-Json
    } finally {
        if ($process) { $process.Dispose() }
    }
}

function Find-PageNumberForText {
    param([string]$Text, $Pages)
    if (-not $Text) { return $null }
    $needle = (Normalize-Text $Text -replace '\s+', ' ').Trim()
    if ($needle.Length -gt 180) { $needle = $needle.Substring(0, 180) }
    if (-not $needle) { return $null }
    foreach ($page in @($Pages)) {
        $pageText = (Normalize-Text ([string](Get-PropertyValue $page 'Text' '')) -replace '\s+', ' ').Trim()
        if ($pageText.IndexOf($needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return Get-PropertyValue $page 'PageNumber'
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
            $text = Normalize-Text (Get-PropertyValue $package 'text' '')
            if ($text.Length -ge 120) {
                return [pscustomobject]@{
                    text = $text
                    pages = @(Get-PropertyValue $package 'pages' @())
                    asset_directory = Get-PropertyValue $package 'asset_directory' $assetDirectory
                    assets_truncated = [bool](Get-PropertyValue $package 'assets_truncated' $false)
                    extraction_method = 'paper-parser-package'
                }
            }
        } catch {
            # 发布包已携带解析器；失败时直接报告，避免再启动无法设置超时的 Word COM 回退。
            throw "Bundled PDF parser could not safely extract this paper. $($_.Exception.Message)"
        }
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

function Get-PaperText {
    param([string]$FilePath)
    return (Get-PaperExtraction $FilePath $null).text
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
    $figureMatches = [regex]::Matches($Text, '(?:\b(?:fig(?:ure)?\.?)\s*|图\s*)(\d+)([a-zA-Z])?', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $seen = @{}
    $figures = @()
    foreach ($match in $figureMatches) {
        $id = "fig-$($match.Groups[1].Value)$($match.Groups[2].Value)".ToLowerInvariant()
        if ($seen.ContainsKey($id)) { continue }
        $seen[$id] = $true
        $start = [Math]::Max(0, $match.Index - 120)
        $length = [Math]::Min($Text.Length - $start, 380)
        $context = Normalize-Text $Text.Substring($start, $length)
        $figures += [pscustomobject]@{
            id = $id
            label = $match.Value
            context = $context
            source_page = Find-PageNumberForText $context $Pages
        }
        if ($figures.Count -ge 30) { break }
    }
    return $figures
}

function Get-CandidateClaims {
    param($Sections, $Pages)
    $resultSections = @($Sections | Where-Object { $_.title -match 'result|结果' })
    $sourceSections = if ($resultSections.Count) { $resultSections } else { $Sections }
    $claims = @()
    foreach ($section in $sourceSections) {
        foreach ($sentence in (Split-PaperSentences $section.text)) {
            if ($sentence -notmatch 'significant|increase|decrease|improve|associated|demonstrate|show|suggest|support|显著|增加|减少|提高|降低|改善|相关|表明|显示|提示|支持') { continue }
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
    return [pscustomobject]@{
        schema_version = '0.3'
        source_file = [IO.Path]::GetFullPath($filePath)
        paper = Get-PaperMetadata $text $fallback
        extraction = [pscustomobject]@{
            method = Get-PropertyValue $extraction 'extraction_method'
            asset_directory = Get-PropertyValue $extraction 'asset_directory'
            assets_truncated = [bool](Get-PropertyValue $extraction 'assets_truncated' $false)
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
            [pscustomobject]@{
                id = $_.id
                title = $_.title
                source_page = Find-PageNumberForText $_.text $pages
                excerpt = $_.text.Substring(0, [Math]::Min(1600, $_.text.Length))
            }
        })
        claims = Get-CandidateClaims $sections $pages
        figures = Get-FigureReferences $text $pages
        journal_club_defaults = [pscustomobject]@{ required_sections = $DefaultRequiredSections }
        ambiguities = @('Candidate claims are not publication-ready facts until a reviewer confirms the original paper.')
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

function Get-Bullets {
    param([string]$Text, [int]$Maximum = 3)
    return @(Split-PaperSentences $Text | Select-Object -First $Maximum | ForEach-Object { $_ -replace '\s+', ' ' })
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

function Find-FigureIdForClaim {
    param($Claim, $Figures)
    $claimText = [string](Get-PropertyValue $Claim 'text' '')
    $normalizedClaimText = $claimText -replace '\s+', ''
    foreach ($figure in @($Figures)) {
        $label = [string](Get-PropertyValue $figure 'label' '')
        $normalizedLabel = $label -replace '\s+', ''
        if ($normalizedLabel -and $normalizedClaimText.IndexOf($normalizedLabel, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return Get-PropertyValue $figure 'id'
        }
    }
    # 只有当该结论所在页恰好只有一个已识别图号时，才采用页码作为保守辅助匹配；多图页绝不猜测。
    $claimPages = @((Get-PropertyValue $Claim 'evidence' @()) | ForEach-Object { Get-PropertyValue $_ 'page_number' } | Where-Object { $null -ne $_ })
    foreach ($claimPage in $claimPages) {
        $pageFigures = @($Figures | Where-Object { (Get-PropertyValue $_ 'source_page') -eq $claimPage })
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

function Get-FigureAssetCandidates {
    param($Figure)
    if ($null -eq $Figure) { return @() }

    # 只透传解析器或用户已提供的图像资产路径，绝不根据图号猜测磁盘文件。
    $candidates = @()
    foreach ($propertyName in @('figure_asset_candidates', 'asset_candidates', 'image_candidates')) {
        foreach ($value in @(Get-PropertyValue $Figure $propertyName @())) {
            foreach ($candidate in @(ConvertTo-NonEmptyStringArray $value)) {
                if ($candidate -notin $candidates) { $candidates += $candidate }
            }
        }
    }
    foreach ($propertyName in @('suggested_image_path', 'image_path', 'asset_path')) {
        $path = Get-PropertyValue $Figure $propertyName
        if ($path -and $path -notin $candidates) { $candidates += [string]$path }
    }
    return @($candidates)
}

function Get-ExplicitFigureAssetPath {
    param($Selections, [string]$FigureId)
    if ($null -eq $Selections -or -not $FigureId) { return $null }
    $candidate = Get-PropertyValue $Selections $FigureId
    if (-not $candidate) { return $null }
    $path = [IO.Path]::GetFullPath([string]$candidate)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Selected figure asset for $FigureId was not found: $path"
    }
    $supportedExtensions = @('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.emf', '.wmf', '.svg')
    if ([IO.Path]::GetExtension($path).ToLowerInvariant() -notin $supportedExtensions) {
        throw "Selected figure asset for $FigureId has an unsupported image extension: $path"
    }
    return $path
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
        [string[]]$FigureAssetCandidates = @(),
        [string]$EvidenceStatus = 'source-backed',
        [string]$ContentMode = 'author-direct'
    )

    $slide = [ordered]@{
        id = $Id
        kind = $Kind
        section = $Section
        title = $Title
        takeaway = $Takeaway
        bullets = @(ConvertTo-NonEmptyStringArray $Bullets)
        source_claim_ids = @(ConvertTo-NonEmptyStringArray $SourceClaimIds)
        source_section_ids = @(ConvertTo-NonEmptyStringArray $SourceSectionIds)
        source_figure_ids = @(ConvertTo-NonEmptyStringArray $SourceFigureIds)
        evidence_status = $EvidenceStatus
        content_mode = $ContentMode
    }
    if ($SuggestedFigureId) { $slide.suggested_figure_id = $SuggestedFigureId }
    if ($SuggestedImagePath) { $slide.suggested_image_path = $SuggestedImagePath }
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
    $duration = [int](Get-PropertyValue $Arguments 'duration_minutes' 15)
    if ($duration -lt 5 -or $duration -gt 90) { throw 'duration_minutes must be between 5 and 90.' }
    $language = Get-PropertyValue $Arguments 'language' 'zh-CN'
    $audience = Get-PropertyValue $Arguments 'audience' 'lab'
    if ($audience -notin @('lab', 'mixed', 'expert')) { $audience = 'lab' }
    $requiredSections = @(Resolve-RequiredSections (Get-PropertyValue $Arguments 'required_sections'))
    $figureAssetSelection = Get-PropertyValue $Arguments 'figure_asset_selection'
    $isChinese = Test-ChineseLanguage $language
    $copy = if ($isChinese) {
        @{
            title_subtitle = '组会汇报 | 可编辑演示文稿草稿'
            journal_club = '文献组会'
            tldr_title = '一句话总结：这篇论文回答了什么问题？'
            tldr_review = '请核对摘要并补充有原文依据的总结后再汇报。'
            tldr_notice = '这是一条核对提示，不是科学结论。'
            experimental_title = '实验数据 {0}'
            experimental_check = '汇报前请结合原始图注核对统计学描述和因果措辞。'
            experimental_figure = '将引用的论文图作为原子图像插入，图中文字、箭头和标注应保持可编辑。'
            # 保留 ASCII 标签，使 Windows PowerShell 的跨进程 JSON 传输在非 UTF-8 控制台中仍可审计。
            presenter_direction = 'Presenter discussion / 汇报者讨论：请设计能够回应已报告局限性的下一步实验。'
            reported_limitation = '已报告的局限性：{0}'
            presenter_proposal = '汇报者建议：所有验证性方案均应标注为讨论，而非作者结论。'
            takehome_title = '核心结论'
            takehome_guard = '结论必须限定在原始研究的数据和方法边界内。'
            audience_note = '请为 {0} 听众准备讨论问题。'
        }
    } else {
        @{
            title_subtitle = 'Journal Club | Editable presentation draft'
            journal_club = 'Journal club'
            tldr_title = 'TL;DR: What does this paper answer?'
            tldr_review = 'Review the abstract and add a source-backed summary before presenting.'
            tldr_notice = 'This is a review notice, not a scientific conclusion.'
            experimental_title = 'Experimental data {0}'
            experimental_check = 'Verify statistics and causal wording against the original figure caption before presenting.'
            experimental_figure = 'Use the cited figure as an atomic image and keep labels, arrows, and callouts editable.'
            presenter_direction = 'Presenter discussion: define the next experiment that addresses the reported limitation.'
            reported_limitation = 'Reported limitation: {0}'
            presenter_proposal = 'Presenter proposal: label any suggested validation as discussion, not as an author claim.'
            takehome_title = 'Take-home message'
            takehome_guard = 'Keep conclusions within the limits of the original data and methods.'
            audience_note = 'Prepare discussion for a {0} audience.'
        }
    }
    $claims = @(Get-PropertyValue $evidencePack 'claims' @())
    $figures = @(Get-PropertyValue $evidencePack 'figures' @())

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
        -EvidenceStatus 'not-applicable' -ContentMode 'metadata'
    $titleSlide | Add-Member -NotePropertyName subtitle -NotePropertyValue $copy.title_subtitle
    $slides += $titleSlide

    if ($summaryEvidence.Count) {
        $summaryText = @($summaryEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'tl-dr' -Kind 'summary' -Section 'summary' -Title $copy.tldr_title `
            -Takeaway $summaryText[0] -Bullets $summaryText -SourceSectionIds (Get-EvidenceItemSectionIds $summaryEvidence)
    } else {
        $slides += New-JournalClubSlide -Id 'tl-dr' -Kind 'summary' -Section 'summary' -Title $copy.tldr_title `
            -Takeaway $copy.tldr_review -Bullets @($copy.tldr_notice) -EvidenceStatus 'missing' -ContentMode 'needs-review'
    }

    if ($backgroundEvidence.Count) {
        $backgroundText = @($backgroundEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'background' -Kind 'background' -Section 'background' -Title (Get-JournalClubSectionTitle 'background' $language) `
            -Takeaway $backgroundText[0] -Bullets $backgroundText -SourceSectionIds (Get-EvidenceItemSectionIds $backgroundEvidence)
    } elseif ('background' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'background' -Section 'background' -FallbackSections @($backgroundSection) -Language $language
    }

    if ($innovationEvidence.Count) {
        $innovationText = @($innovationEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'innovation' -Kind 'innovation' -Section 'innovation' -Title (Get-JournalClubSectionTitle 'innovation' $language) `
            -Takeaway $innovationText[0] -Bullets $innovationText -SourceSectionIds (Get-EvidenceItemSectionIds $innovationEvidence)
    } elseif ('innovation' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'innovation' -Section 'innovation' -FallbackSections $innovationSections -Language $language
    }

    if ($methodsEvidence.Count) {
        $methodsText = @($methodsEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'methods' -Kind 'methods' -Section 'methods' -Title (Get-JournalClubSectionTitle 'methods' $language) `
            -Takeaway $methodsText[0] -Bullets $methodsText -SourceSectionIds (Get-EvidenceItemSectionIds $methodsEvidence)
    } elseif ('methods' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'methods' -Section 'methods' -FallbackSections @($methodsSection) -Language $language
    }

    $maxResults = [Math]::Max(1, [Math]::Min(5, [Math]::Round($duration / 5)))
    $resultClaims = @($claims | Select-Object -First $maxResults)
    if ($resultClaims.Count) {
        for ($i = 0; $i -lt $resultClaims.Count; $i++) {
            $claim = $resultClaims[$i]
            $claimId = Get-PropertyValue $claim 'id'
            $claimText = Get-PropertyValue $claim 'text'
            $claimSectionIds = @(Get-ClaimEvidenceSectionIds $claim)
            $figureId = Find-FigureIdForClaim $claim $figures
            $figureIds = if ($figureId) { @([string]$figureId) } else { @() }
            $sourceFigure = Get-FigureById $figures $figureId
            $figureAssetCandidates = @(Get-FigureAssetCandidates $sourceFigure)
            $explicitImagePath = Get-ExplicitFigureAssetPath -Selections $figureAssetSelection -FigureId $figureId
            if ($explicitImagePath -and $explicitImagePath -notin $figureAssetCandidates) {
                $figureAssetCandidates = @($explicitImagePath) + $figureAssetCandidates
            }
            $suggestedImagePath = if ($figureAssetCandidates.Count) { $figureAssetCandidates[0] } else { $null }
            $claimEvidenceStatus = if ($claimSectionIds.Count -and $claimId -and $claimText) { 'source-backed' } else { 'missing' }
            $slides += New-JournalClubSlide -Id "experimental-data-$($i + 1)" -Kind 'result' -Section 'experimental_data' `
                -Title ($copy.experimental_title -f ($i + 1)) -Takeaway $claimText `
                -Bullets @($copy.experimental_check, $copy.experimental_figure) `
                -SourceClaimIds @([string]$claimId) -SourceSectionIds $claimSectionIds -SourceFigureIds $figureIds `
                -SuggestedFigureId $figureId -SuggestedImagePath $suggestedImagePath -FigureAssetCandidates $figureAssetCandidates `
                -EvidenceStatus $claimEvidenceStatus -ContentMode 'author-direct'
        }
    } elseif ('experimental_data' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'experimental-data' -Section 'experimental_data' -FallbackSections @($resultsSection) -Language $language
    }

    if ($limitationEvidence.Count) {
        $limitationText = @($limitationEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'limitations' -Kind 'limitations' -Section 'limitations' -Title (Get-JournalClubSectionTitle 'limitations' $language) `
            -Takeaway $limitationText[0] -Bullets $limitationText -SourceSectionIds (Get-EvidenceItemSectionIds $limitationEvidence)
    } elseif ('limitations' -in $requiredSections) {
        $slides += New-MissingSourceSlide -Id 'limitations' -Section 'limitations' -FallbackSections $discussionSections -Language $language
    }

    if ($futureEvidence.Count) {
        $futureText = @($futureEvidence | ForEach-Object { $_.text })
        $slides += New-JournalClubSlide -Id 'future-directions' -Kind 'future-directions' -Section 'future_directions' -Title (Get-JournalClubSectionTitle 'future_directions' $language) `
            -Takeaway $futureText[0] -Bullets $futureText -SourceSectionIds (Get-EvidenceItemSectionIds $futureEvidence)
    } elseif ($limitationEvidence.Count) {
        # 原文未给出下一步时，只生成明确标为“汇报者讨论”的建议，并回链到作者已述局限。
        $reportedLimitation = $limitationEvidence[0].text
        $slides += New-JournalClubSlide -Id 'future-directions' -Kind 'future-directions' -Section 'future_directions' -Title (Get-JournalClubSectionTitle 'future_directions' $language) `
            -Takeaway $copy.presenter_direction `
            -Bullets @(($copy.reported_limitation -f $reportedLimitation), $copy.presenter_proposal) `
            -SourceSectionIds (Get-EvidenceItemSectionIds $limitationEvidence) -EvidenceStatus 'source-backed' -ContentMode 'presenter-discussion'
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
            -SourceClaimIds $takeawayClaimIds -SourceSectionIds $takeawaySectionIds
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
        $overwrite = [bool](Get-PropertyValue $Arguments 'overwrite' $false)
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
    $claimIds = Get-ObjectIds $claims
    $sectionIds = Get-ObjectIds $sourceSections
    $figureIds = Get-ObjectIds $figures
    $claimsById = @{}
    foreach ($claim in $claims) {
        $claimId = Get-PropertyValue $claim 'id'
        if ($claimId) { $claimsById[[string]$claimId] = $claim }
    }
    $findings = @()
    $slides = @(Get-PropertyValue $deck 'slides' @())
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
        if ((Get-PropertyValue $slide 'kind') -ne 'title' -and -not (Get-PropertyValue $slide 'takeaway' '')) {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'warning'; category = 'narrative'; issue = 'Missing takeaway'; correction = 'Add one talkable takeaway.' }
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
            }
            if ($slideSection -eq 'future_directions' -and $contentMode -eq 'presenter-discussion') {
                if ($slideSectionIds.Count -eq 0 -or (Get-PropertyValue $slide 'takeaway' '') -notmatch '(?i)presenter|discussion|汇报者|讨论') {
                    $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'hard'; category = 'future-directions'; issue = 'Presenter-derived future direction is not clearly labelled and sourced.'; correction = 'Label it as presenter discussion and cite the limitation or result that motivates it.' }
                }
            }
        } elseif ((Get-PropertyValue $slide 'kind') -ne 'title' -and -not $hasTraceableSource -and $evidenceStatus -ne 'not-applicable') {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'warning'; category = 'traceability'; issue = 'Non-title slide has no traceable source.'; correction = 'Add a claim, paper section, or figure reference before final presentation.' }
        }
        if (@(Get-PropertyValue $slide 'bullets' @()).Count -gt 5) {
            $findings += [pscustomobject]@{ slide = $slideNumber; severity = 'warning'; category = 'readability'; issue = 'More than five bullets'; correction = 'Split the slide or retain only supporting points.' }
        }
    }

    # 结构门禁：默认六个模块或调用方显式配置的模块，任意一个缺页都会阻断生成。
    foreach ($requiredSection in $requiredSections) {
        if (@($requiredSlideIndexes[$requiredSection]).Count -eq 0) {
            $findings += [pscustomobject]@{ slide = $null; severity = 'hard'; category = 'required-section'; issue = "Missing required section: $requiredSection"; correction = "Add a slide with section='$requiredSection' and source-backed content." }
        }
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
        return Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club\$DefaultStem-$([Guid]::NewGuid().ToString('N'))"
    }
    $absoluteDirectory = [IO.Path]::GetFullPath($RequestedDirectory)
    if (-not (Test-DirectoryIsEmptyOrMissing $absoluteDirectory)) {
        throw "preview_directory must be a new or empty directory: $absoluteDirectory"
    }
    return $absoluteDirectory
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
    $exportPreviews = [bool](Get-PropertyValue $Arguments 'export_previews' $false)
    if ($filePath) {
        $absolutePath = Get-ExistingPowerPointPath $filePath
        $session = $null
        try {
            $session = Open-PowerPointPresentationReadOnly -Path $absolutePath
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
        export_previews = [bool](Get-PropertyValue $Arguments 'export_previews' $true)
        preview_directory = Get-PropertyValue $Arguments 'preview_directory'
    }
    return Invoke-InspectPowerPoint $inspectionArguments
}

function Invoke-GeneratePptx {
    param($Arguments)
    $deck = Get-PropertyValue $Arguments 'deck_spec'
    $outputPath = Get-PropertyValue $Arguments 'output_path'
    if (-not $deck -or -not $outputPath) { throw 'deck_spec and output_path are required.' }
    # 不能仅依赖调用方先执行 audit：直接调用生成工具也必须经过必备模块与证据门禁。
    $audit = Invoke-AuditDeck ([pscustomobject]@{ deck_spec = $deck })
    if (-not $audit.pass) {
        $blockingSummary = @($audit.findings | Where-Object { $_.severity -eq 'hard' } | Select-Object -First 5 | ForEach-Object { $_.issue }) -join ' | '
        throw "Deck failed the mandatory content audit. Resolve hard findings before generating PowerPoint: $blockingSummary"
    }
    $overwrite = [bool](Get-PropertyValue $Arguments 'overwrite' $false)
    $absoluteOutput = Resolve-RequestedOutputPath -Path $outputPath -RequiredExtension '.pptx' -Overwrite $overwrite
    $specPath = "$absoluteOutput.deck-spec.json"
    if ((Test-Path -LiteralPath $specPath -PathType Leaf) -and -not $overwrite) {
        throw "Deck-spec sidecar already exists. Set overwrite=true only when it may be replaced: $specPath"
    }
    $deck | ConvertTo-Json -Depth 50 | Set-Content -Encoding UTF8 -LiteralPath $specPath
    $keepOpen = [bool](Get-PropertyValue $Arguments 'keep_powerpoint_open' $false)
    $exportPreviews = [bool](Get-PropertyValue $Arguments 'export_previews' $true)
    $previewDirectory = if ($exportPreviews) { Resolve-PreviewDirectory (Get-PropertyValue $Arguments 'preview_directory') 'generated-previews' } else { $null }
    $generatorArguments = @('-NoProfile', '-STA', '-ExecutionPolicy', 'RemoteSigned', '-File', $PptGeneratorPath, '-DeckSpecPath', $specPath, '-OutputPath', $absoluteOutput, '-KeepOpen', $keepOpen, '-Overwrite')
    if ($previewDirectory) { $generatorArguments += @('-PreviewDirectory', $previewDirectory) }
    if (-not $exportPreviews) { $generatorArguments += '-SkipPreviewExport' }
    $generatorOutput = @(& powershell.exe @generatorArguments 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "PowerPoint generation failed: $($generatorOutput -join [Environment]::NewLine)" }
    $jsonLine = @($generatorOutput | ForEach-Object { [string]$_ } | Where-Object { $_.TrimStart().StartsWith('{') } | Select-Object -Last 1)
    if ($jsonLine.Count -ne 1) { throw 'PowerPoint generator completed without its required JSON quality report.' }
    try {
        $nativeResult = $jsonLine[0] | ConvertFrom-Json
    } catch {
        throw "PowerPoint generator returned an invalid JSON quality report. $($_.Exception.Message)"
    }
    return [pscustomobject]@{
        output_path = $absoluteOutput
        deck_spec_path = $specPath
        target_application = 'Microsoft PowerPoint'
        connection_scope = 'new-background-presentation'
        editable_contract = 'native-text-and-shapes'
        content_audit = $audit
        quality_audit = Get-PropertyValue $nativeResult 'quality_audit'
        preview_directory = Get-PropertyValue $nativeResult 'preview_directory'
        preview_paths = Get-PropertyValue (Get-PropertyValue $nativeResult 'quality_audit') 'preview_paths' @()
        next_step = 'Open the generated PPTX in PowerPoint; the returned PNG previews and quality audit are the final verification record.'
    }
}

function Get-Tools {
    return @(
        [pscustomobject]@{ name = 'analyse_paper'; description = 'Parse a paper into sections, candidate claims, page evidence, and figure assets without Node.js.'; inputSchema = [ordered]@{ type = 'object'; required = @('file_path'); properties = [ordered]@{ file_path = [ordered]@{ type = 'string'; description = 'Absolute paper path.' }; asset_output_dir = [ordered]@{ type = 'string'; description = 'Optional directory for parser-extracted paper image assets. Defaults to a user temp directory.' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'cleanup_paper_assets'; description = 'Delete only a confirmed default temporary paper-asset directory created by this plugin.'; inputSchema = [ordered]@{ type = 'object'; required = @('asset_output_dir', 'confirm'); properties = [ordered]@{ asset_output_dir = [ordered]@{ type = 'string' }; confirm = [ordered]@{ type = 'boolean'; const = $true } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'design_journal_club_deck'; description = 'Create an evidence-backed journal-club deck specification with required content sections.'; inputSchema = [ordered]@{ type = 'object'; required = @('evidence_pack'); properties = [ordered]@{ evidence_pack = [ordered]@{ type = 'object' }; duration_minutes = [ordered]@{ type = 'integer'; minimum = 5; maximum = 90; default = 15 }; language = [ordered]@{ type = 'string'; default = 'zh-CN' }; audience = [ordered]@{ type = 'string'; enum = @('lab', 'mixed', 'expert'); default = 'lab' }; required_sections = [ordered]@{ type = 'array'; minItems = 1; uniqueItems = $true; items = [ordered]@{ type = 'string'; enum = $KnownJournalClubSections }; default = $DefaultRequiredSections; description = 'Defaults to background, innovation, methods, experimental_data, limitations, and future_directions.' }; figure_asset_selection = [ordered]@{ type = 'object'; description = 'Optional explicit mapping from evidence-pack figure id to a reviewed extracted image path. Never guessed automatically.'; additionalProperties = [ordered]@{ type = 'string' } }; output_path = [ordered]@{ type = 'string'; description = 'Optional new .json deck-spec path.' }; overwrite = [ordered]@{ type = 'boolean'; default = $false } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'audit_journal_club_deck'; description = 'Hard-fail missing required sections and audit evidence traceability before PowerPoint generation.'; inputSchema = [ordered]@{ type = 'object'; required = @('deck_spec'); properties = [ordered]@{ deck_spec = [ordered]@{ type = 'object' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'powerpoint_status'; description = 'Inspect Microsoft PowerPoint COM availability and the current window without modifying it.'; inputSchema = [ordered]@{ type = 'object'; properties = [ordered]@{}; additionalProperties = $false } },
        [pscustomobject]@{ name = 'inspect_powerpoint'; description = 'Inspect the active PowerPoint window through COM, or a saved PPTX in read-only mode. It never edits the presentation.'; inputSchema = [ordered]@{ type = 'object'; properties = [ordered]@{ file_path = [ordered]@{ type = 'string'; description = 'Optional saved .pptx path. Omit only to inspect the active current window.' }; export_previews = [ordered]@{ type = 'boolean'; default = $false }; preview_directory = [ordered]@{ type = 'string'; description = 'New or empty folder for native PowerPoint PNG previews.' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'audit_editable_pptx'; description = 'Open a saved PPTX read-only, verify native editable objects, and export final native PowerPoint PNG previews.'; inputSchema = [ordered]@{ type = 'object'; required = @('file_path'); properties = [ordered]@{ file_path = [ordered]@{ type = 'string' }; export_previews = [ordered]@{ type = 'boolean'; default = $true }; preview_directory = [ordered]@{ type = 'string' } }; additionalProperties = $false } },
        [pscustomobject]@{ name = 'generate_editable_pptx'; description = 'Create a new editable PowerPoint presentation after the required-section audit passes. It does not modify the active presentation.'; inputSchema = [ordered]@{ type = 'object'; required = @('deck_spec', 'output_path'); properties = [ordered]@{ deck_spec = [ordered]@{ type = 'object' }; output_path = [ordered]@{ type = 'string'; description = 'New .pptx path.' }; overwrite = [ordered]@{ type = 'boolean'; default = $false }; keep_powerpoint_open = [ordered]@{ type = 'boolean'; default = $false }; export_previews = [ordered]@{ type = 'boolean'; default = $true }; preview_directory = [ordered]@{ type = 'string'; description = 'New or empty folder for native PowerPoint PNG previews.' } }; additionalProperties = $false } }
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

function Invoke-McpTool {
    param([string]$Name, $Arguments)
    switch ($Name) {
        'analyse_paper' { return Invoke-AnalysePaper $Arguments }
        'cleanup_paper_assets' { return Remove-TemporaryPaperAssets $Arguments }
        'design_journal_club_deck' { return Invoke-DesignDeck $Arguments }
        'audit_journal_club_deck' { return Invoke-AuditDeck $Arguments }
        'powerpoint_status' { return Get-PowerPointStatus }
        'inspect_powerpoint' { return Invoke-InspectPowerPoint $Arguments }
        'audit_editable_pptx' { return Invoke-AuditEditablePptx $Arguments }
        'generate_editable_pptx' { return Invoke-GeneratePptx $Arguments }
        default { throw "Unknown tool: $Name" }
    }
}

if ($Demo) {
    if (-not $DemoInputPath) { throw 'DemoInputPath is required when using -Demo.' }
    $evidence = Invoke-AnalysePaper ([pscustomobject]@{ file_path = $DemoInputPath })
    $deck = Invoke-DesignDeck ([pscustomobject]@{ evidence_pack = $evidence; duration_minutes = 15; language = 'zh-CN'; audience = 'lab'; output_path = $DemoOutputPath })
    [Console]::Out.WriteLine(([pscustomobject]@{ deck_spec = $DemoOutputPath; audit = Invoke-AuditDeck ([pscustomobject]@{ deck_spec = $deck }) } | ConvertTo-Json -Depth 80))
    exit 0
}

while ($null -ne ($line = [Console]::In.ReadLine())) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $request = $null
    try {
        $request = $line.TrimStart([char]0xFEFF) | ConvertFrom-Json
        switch ($request.method) {
            'notifications/initialized' { continue }
            'initialize' {
                $requestedProtocol = Get-PropertyValue $request.params 'protocolVersion' '2025-06-18'
                Write-McpResponse $request.id ([ordered]@{
                    protocolVersion = Select-McpProtocolVersion $requestedProtocol
                    capabilities = [ordered]@{ tools = [ordered]@{} }
                    serverInfo = [ordered]@{ name = 'paper-to-journal-club'; version = '1.0.0' }
                })
            }
            'tools/list' { Write-McpResponse $request.id ([ordered]@{ tools = Get-Tools }) }
            'tools/call' {
                $toolResult = Invoke-McpTool (Get-PropertyValue $request.params 'name') (Get-PropertyValue $request.params 'arguments' ([pscustomobject]@{}))
                Write-McpResponse $request.id ([ordered]@{ content = @([ordered]@{ type = 'text'; text = ($toolResult | ConvertTo-Json -Depth 80) }) })
            }
            default { Write-McpResponse $request.id $null "Unsupported method: $($request.method)" }
        }
    } catch {
        Write-McpResponse (Get-PropertyValue $request 'id') $null $_.Exception.Message
    }
}
