<#
  PowerPoint COM quality helpers shared by generation and read-only inspection.

  The checks deliberately distinguish deterministic structural faults from visual
  review items. A PPTX is not approved merely because SaveAs succeeded.
#>

# 安全边界：MCP 只能在用户常用资料目录、插件专用临时目录，或由用户在进程环境中
# 明确声明的目录内读写。这样即使上游提示词被注入，也不能把本机任意路径当成输入或输出。
function ConvertTo-PaperToJournalClubAbsolutePath {
    param([string]$Path, [string]$ParameterName = 'path')

    if ([string]::IsNullOrWhiteSpace($Path)) { throw "$ParameterName is required." }
    if ($Path.Length -gt 4096) { throw "$ParameterName exceeds the 4096-character path safety limit." }
    if ($Path.IndexOf([char]0) -ge 0 -or $Path -match "[\r\n]") {
        throw "$ParameterName must not contain null characters or line breaks."
    }

    try {
        $fullPath = [IO.Path]::GetFullPath($Path)
    } catch {
        throw "$ParameterName is not a valid absolute or relative Windows path."
    }

    # 不把磁盘根目录的 C:\ 误裁剪为 C:；其余路径统一去掉尾部斜杠，便于后续前缀比较。
    $pathRoot = [IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.Length -gt $pathRoot.Length) {
        $fullPath = $fullPath.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    }
    return $fullPath
}

function Test-PaperToJournalClubPathInsideRoot {
    param([string]$Path, [string]$Root)

    $candidate = ConvertTo-PaperToJournalClubAbsolutePath -Path $Path
    $trustedRoot = ConvertTo-PaperToJournalClubAbsolutePath -Path $Root
    if ($candidate.Equals($trustedRoot, [StringComparison]::OrdinalIgnoreCase)) { return $true }

    $rootPrefix = if ($trustedRoot.EndsWith([IO.Path]::DirectorySeparatorChar) -or $trustedRoot.EndsWith([IO.Path]::AltDirectorySeparatorChar)) {
        $trustedRoot
    } else {
        $trustedRoot + [IO.Path]::DirectorySeparatorChar
    }
    return $candidate.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)
}

function Get-PaperToJournalClubTemporaryRoot {
    return ConvertTo-PaperToJournalClubAbsolutePath -Path (Join-Path ([IO.Path]::GetTempPath()) 'paper-to-journal-club')
}

function Get-PaperToJournalClubUserDataRoots {
    $roots = @()
    foreach ($folder in @([Environment+SpecialFolder]::DesktopDirectory, [Environment+SpecialFolder]::MyDocuments)) {
        $candidate = [Environment]::GetFolderPath($folder)
        if (-not [string]::IsNullOrWhiteSpace($candidate)) { $roots += $candidate }
    }
    if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        $roots += (Join-Path $env:USERPROFILE 'Downloads')
    }

    # 例如实验室把数据盘挂载为 D:\Research 时，用户可在启动 Codex 前显式配置：
    # PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS=D:\Research;E:\SharedPapers
    # 环境变量不是 MCP 参数，模型无法在一次工具调用中扩大该边界。
    if (-not [string]::IsNullOrWhiteSpace($env:PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS)) {
        $roots += @($env:PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $resolved = @()
    foreach ($root in $roots) {
        try {
            $absoluteRoot = ConvertTo-PaperToJournalClubAbsolutePath -Path $root -ParameterName 'approved root'
            if ($absoluteRoot -notin $resolved) { $resolved += $absoluteRoot }
        } catch {
            # 无效配置项不能退化为放开所有路径，直接忽略并保留其它有效根目录。
        }
    }
    return @($resolved)
}

function Get-PaperToJournalClubApprovedWriteRoots {
    return @((Get-PaperToJournalClubUserDataRoots) + (Get-PaperToJournalClubTemporaryRoot))
}

function Assert-PaperToJournalClubNoReparsePoint {
    param([string]$Path, [string]$ParameterName = 'path')

    # 输出目标尚不存在时，从最近的已存在父目录开始检查。拒绝 junction/symlink，
    # 防止“看似位于允许目录、实际跳转到其它磁盘位置”的路径穿越。
    $cursor = ConvertTo-PaperToJournalClubAbsolutePath -Path $Path -ParameterName $ParameterName
    while (-not (Test-Path -LiteralPath $cursor)) {
        $parent = Split-Path -Parent $cursor
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $cursor) { break }
        $cursor = $parent
    }

    while (-not [string]::IsNullOrWhiteSpace($cursor)) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if (([int]$item.Attributes -band [int][IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "$ParameterName traverses a symbolic link or junction, which is not allowed: $cursor"
            }
        }
        $parent = Split-Path -Parent $cursor
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $cursor) { break }
        $cursor = $parent
    }
}

function Assert-PaperToJournalClubAllowedPath {
    param(
        [string]$Path,
        [string[]]$AllowedRoots,
        [string]$ParameterName = 'path',
        [switch]$AllowRootItself
    )

    $absolutePath = ConvertTo-PaperToJournalClubAbsolutePath -Path $Path -ParameterName $ParameterName
    $matchedRoot = $null
    foreach ($root in @($AllowedRoots)) {
        if (-not [string]::IsNullOrWhiteSpace($root) -and (Test-PaperToJournalClubPathInsideRoot -Path $absolutePath -Root $root)) {
            $matchedRoot = ConvertTo-PaperToJournalClubAbsolutePath -Path $root -ParameterName 'approved root'
            break
        }
    }
    if ($null -eq $matchedRoot) {
        throw "$ParameterName must be inside an approved user data directory. Configure PAPER_TO_JOURNAL_CLUB_ALLOWED_ROOTS before starting Codex for another folder."
    }
    if (-not $AllowRootItself -and $absolutePath.Equals($matchedRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$ParameterName must point to a child of an approved directory, not the directory root itself."
    }
    Assert-PaperToJournalClubNoReparsePoint -Path $absolutePath -ParameterName $ParameterName
    return $absolutePath
}

function Test-PaperToJournalClubImageSignature {
    param([string]$Path, [string]$Extension)

    $stream = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $buffer = New-Object byte[] 16
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($Extension -eq '.png') {
            return $read -ge 8 -and $buffer[0] -eq 137 -and $buffer[1] -eq 80 -and $buffer[2] -eq 78 -and $buffer[3] -eq 71 -and $buffer[4] -eq 13 -and $buffer[5] -eq 10 -and $buffer[6] -eq 26 -and $buffer[7] -eq 10
        }
        # 只允许 JPEG 的 SOI 标记及紧随其后的 marker，扩展名不能伪装为 JPEG。
        return $read -ge 3 -and $buffer[0] -eq 255 -and $buffer[1] -eq 216 -and $buffer[2] -eq 255
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}

function Get-PaperToJournalClubRasterDimensions {
    param([string]$Path, [string]$Extension)

    $stream = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ($Extension -eq '.png') {
            $header = New-Object byte[] 24
            if ($stream.Read($header, 0, $header.Length) -lt $header.Length) { throw 'PNG header is incomplete.' }
            # PNG IHDR 的宽高均为 32 位大端整数。先读取该头部，避免 GDI+ 为图像炸弹分配像素缓冲区。
            [int64]$width = ([int64]$header[16] * 16777216) + ([int64]$header[17] * 65536) + ([int64]$header[18] * 256) + [int64]$header[19]
            [int64]$height = ([int64]$header[20] * 16777216) + ([int64]$header[21] * 65536) + ([int64]$header[22] * 256) + [int64]$header[23]
            return [pscustomobject]@{ width = $width; height = $height }
        }

        # JPEG 的尺寸在 SOF marker 中。扫描段头而不解码像素数据，支持基线与常见渐进 JPEG。
        if ($stream.ReadByte() -ne 255 -or $stream.ReadByte() -ne 216) { throw 'JPEG header is incomplete.' }
        while ($true) {
            $prefix = $stream.ReadByte()
            while ($prefix -eq 255) { $prefix = $stream.ReadByte() }
            if ($prefix -lt 0) { break }
            if ($prefix -eq 216 -or $prefix -eq 217 -or $prefix -eq 1 -or ($prefix -ge 208 -and $prefix -le 215)) { continue }
            $lengthHigh = $stream.ReadByte()
            $lengthLow = $stream.ReadByte()
            if ($lengthHigh -lt 0 -or $lengthLow -lt 0) { break }
            $segmentLength = ($lengthHigh * 256) + $lengthLow
            if ($segmentLength -lt 2) { throw 'JPEG contains an invalid segment length.' }
            if ($prefix -in @(192, 193, 194, 195, 197, 198, 199, 201, 202, 203, 205, 206, 207)) {
                if ($segmentLength -lt 7) { throw 'JPEG SOF segment is incomplete.' }
                [void]$stream.ReadByte() # precision
                $heightHigh = $stream.ReadByte(); $heightLow = $stream.ReadByte()
                $widthHigh = $stream.ReadByte(); $widthLow = $stream.ReadByte()
                if ($heightHigh -lt 0 -or $heightLow -lt 0 -or $widthHigh -lt 0 -or $widthLow -lt 0) { throw 'JPEG dimensions are incomplete.' }
                return [pscustomobject]@{ width = [int64](($widthHigh * 256) + $widthLow); height = [int64](($heightHigh * 256) + $heightLow) }
            }
            [void]$stream.Seek($segmentLength - 2, [IO.SeekOrigin]::Current)
        }
        throw 'JPEG does not contain a supported SOF dimension marker.'
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}

function Get-PaperToJournalClubApprovedRasterImage {
    param([string]$ImagePath, [string[]]$AllowedRoots, [string]$ParameterName = 'image path')

    $absolutePath = Assert-PaperToJournalClubAllowedPath -Path $ImagePath -AllowedRoots $AllowedRoots -ParameterName $ParameterName
    if (-not (Test-Path -LiteralPath $absolutePath -PathType Leaf)) { throw "$ParameterName was not found: $absolutePath" }
    $extension = [IO.Path]::GetExtension($absolutePath).ToLowerInvariant()
    if ($extension -notin @('.png', '.jpg', '.jpeg')) {
        throw "$ParameterName must be a PNG or JPEG raster image; SVG, EMF, WMF, GIF, TIFF and other active/complex formats are blocked."
    }
    $file = Get-Item -LiteralPath $absolutePath -Force
    if ($file.Length -lt 1 -or $file.Length -gt 25MB) {
        throw "$ParameterName must be between 1 byte and 25 MB."
    }
    if (-not (Test-PaperToJournalClubImageSignature -Path $absolutePath -Extension $extension)) {
        throw "$ParameterName extension does not match a valid PNG or JPEG file signature."
    }

    $headerDimensions = Get-PaperToJournalClubRasterDimensions -Path $absolutePath -Extension $extension
    if ($headerDimensions.width -lt 1 -or $headerDimensions.height -lt 1 -or $headerDimensions.width -gt 10000 -or $headerDimensions.height -gt 10000 -or ([int64]$headerDimensions.width * [int64]$headerDimensions.height) -gt 40000000) {
        throw "$ParameterName exceeds the 10000-pixel-per-side or 40-megapixel safety limit."
    }

    $stream = $null
    $image = $null
    try {
        Add-Type -AssemblyName System.Drawing
        $stream = [IO.File]::Open($absolutePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        # validateImageData=true 会在交给 PowerPoint 前验证图像编码；只读取尺寸，不保存或执行内容。
        $image = [Drawing.Image]::FromStream($stream, $false, $true)
        $width = [int]$image.Width
        $height = [int]$image.Height
        if ($width -ne [int]$headerDimensions.width -or $height -ne [int]$headerDimensions.height) {
            throw "$ParameterName dimensions changed during image decoding."
        }
        return [pscustomobject]@{ path = $absolutePath; width = $width; height = $height; bytes = [int64]$file.Length }
    } catch {
        if ($_.Exception.Message -match 'safety limit') { throw }
        throw "$ParameterName is not a readable PNG or JPEG raster image."
    } finally {
        if ($image) { $image.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}

function Set-PaperToJournalClubOfficeAutomationSecurity {
    param($Application, [string]$OfficeApplication, [int]$DisplayAlertsValue)

    # msoAutomationSecurityForceDisable = 3。读取不可信 Office 文件前必须先强制禁用宏；
    # 无法设置时宁可失败，不降级为带宏自动化打开。
    try {
        $Application.AutomationSecurity = 3
    } catch {
        throw "Could not force-disable Office macros for $OfficeApplication automation."
    }
    try { $Application.DisplayAlerts = $DisplayAlertsValue } catch { }
}

function Get-ActivePowerPointApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    } catch {
        return $null
    }
}

function Test-PowerPointComRegistration {
    return Test-Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID'
}

function Get-PowerPointPresentationSummary {
    param($Presentation)
    $slides = @($Presentation.Slides)
    return [pscustomobject]@{
        name = $Presentation.Name
        full_name = $Presentation.FullName
        slide_count = $slides.Count
        width = [double]$Presentation.PageSetup.SlideWidth
        height = [double]$Presentation.PageSetup.SlideHeight
        shapes_per_slide = @($slides | ForEach-Object { $_.Shapes.Count })
        native_picture_count = 0
    }
}

function Test-ShapeBounds {
    param($Shape, [double]$SlideWidth, [double]$SlideHeight, [double]$Tolerance = 1)
    $findings = @()
    if ($Shape.Width -le 0 -or $Shape.Height -le 0) {
        $findings += 'non-positive-size'
    }
    if ($Shape.Left -lt -$Tolerance -or $Shape.Top -lt -$Tolerance -or ($Shape.Left + $Shape.Width) -gt ($SlideWidth + $Tolerance) -or ($Shape.Top + $Shape.Height) -gt ($SlideHeight + $Tolerance)) {
        $findings += 'out-of-slide-bounds'
    }
    return $findings
}

function Test-ShapeTextOverflow {
    param($Shape)
    try {
        if ($Shape.HasTextFrame -ne -1 -or $Shape.TextFrame.HasText -ne -1) { return $false }
        if ($Shape.TextFrame.Overflowing -eq -1) { return $true }
    } catch {
        # Some shape types do not expose TextFrame.Overflowing; use TextFrame2 as a fallback.
    }
    try {
        $boundHeight = [double]$Shape.TextFrame2.TextRange.BoundHeight
        return $boundHeight -gt ([double]$Shape.Height + 1)
    } catch {
        return $false
    }
}

function Export-PowerPointPreviews {
    param($Presentation, [string]$PreviewDirectory)
    # 即使该帮助函数被独立调用，也在创建目录前复核路径与 junction，避免绕过 MCP 层。
    $PreviewDirectory = Assert-PaperToJournalClubAllowedPath -Path $PreviewDirectory -AllowedRoots (Get-PaperToJournalClubApprovedWriteRoots) -ParameterName 'PreviewDirectory'
    New-Item -ItemType Directory -Force -Path $PreviewDirectory | Out-Null
    $paths = @()
    foreach ($slide in @($Presentation.Slides)) {
        $fileName = 'slide-{0:D2}.png' -f [int]$slide.SlideIndex
        $path = Join-Path $PreviewDirectory $fileName
        # Export uses the native PowerPoint renderer, avoiding LibreOffice font differences.
        $slide.Export($path, 'PNG', 1920, 1080)
        if (-not (Test-Path -LiteralPath $path)) { throw "PowerPoint did not export preview: $path" }
        $paths += $path
    }
    return $paths
}

function Invoke-PowerPointQualityAudit {
    param(
        $Presentation,
        [string]$PreviewDirectory = $null,
        [switch]$ExportPreviews,
        [double]$Tolerance = 1
    )
    $summary = Get-PowerPointPresentationSummary $Presentation
    $findings = @()
    $pictureCount = 0
    $atomicPictureCount = 0
    if ($summary.slide_count -lt 1) {
        $findings += [pscustomobject]@{ severity = 'hard'; category = 'slides'; slide = $null; shape = $null; issue = 'empty-presentation'; correction = 'Create at least one slide.' }
    }

    foreach ($slide in @($Presentation.Slides)) {
        if ($slide.Shapes.Count -lt 2) {
            $findings += [pscustomobject]@{ severity = 'hard'; category = 'editability'; slide = $slide.SlideIndex; shape = $null; issue = 'too-few-native-objects'; correction = 'Create independently editable slide objects.' }
        }
        foreach ($shape in @($slide.Shapes)) {
            if ($shape.Type -eq 13) {
                $pictureCount++
                if ([string]$shape.AlternativeText -match 'atomic_raster_unit=true') { $atomicPictureCount++ }
            }
            foreach ($issue in @(Test-ShapeBounds -Shape $shape -SlideWidth $summary.width -SlideHeight $summary.height -Tolerance $Tolerance)) {
                $findings += [pscustomobject]@{ severity = 'hard'; category = 'geometry'; slide = $slide.SlideIndex; shape = $shape.Name; issue = $issue; correction = 'Resize or reposition the named shape within the slide canvas.' }
            }
            if (Test-ShapeTextOverflow -Shape $shape) {
                $findings += [pscustomobject]@{ severity = 'hard'; category = 'text-fit'; slide = $slide.SlideIndex; shape = $shape.Name; issue = 'text-overflow'; correction = 'Reduce text, resize the text box, or lower the font size.' }
            }
        }
    }

    $previewPaths = @()
    if ($ExportPreviews) {
        if (-not $PreviewDirectory) { throw 'PreviewDirectory is required when ExportPreviews is set.' }
        $previewPaths = Export-PowerPointPreviews -Presentation $Presentation -PreviewDirectory $PreviewDirectory
    }
    $hardFailures = @($findings | Where-Object { $_.severity -eq 'hard' })
    # 摘要与 raster_summary 使用同一统计值，避免质量报告出现互相矛盾的图片计数。
    $summary.native_picture_count = $pictureCount
    return [pscustomobject]@{
        pass = ($hardFailures.Count -eq 0)
        presentation = $summary
        findings = $findings
        preview_paths = $previewPaths
        raster_summary = [pscustomobject]@{ picture_count = $pictureCount; atomic_picture_count = $atomicPictureCount }
        renderer = 'Microsoft PowerPoint COM Slide.Export'
    }
}

function Open-PowerPointPresentationReadOnly {
    param([string]$Path)
    $application = $null
    $presentation = $null
    try {
        $application = New-Object -ComObject PowerPoint.Application
        # ppAlertsNone = 1；先收紧宏与交互警告，再打开用户提供的演示文稿。
        Set-PaperToJournalClubOfficeAutomationSecurity -Application $application -OfficeApplication 'Microsoft PowerPoint' -DisplayAlertsValue 1
        $application.Visible = -1
        # Slide.Export can hang when a file-backed presentation is opened without a window.
        # Keep a native window, then minimize it so the audit does not repeatedly steal focus.
        $presentation = $application.Presentations.Open([IO.Path]::GetFullPath($Path), $true, $false, $true)
        try { $application.WindowState = 2 } catch { }
        return [pscustomobject]@{ application = $application; presentation = $presentation }
    } catch {
        if ($presentation) { $presentation.Close() }
        if ($application) { $application.Quit() }
        throw
    }
}

function Close-PowerPointReadOnlySession {
    param($Session)
    if ($null -eq $Session) { return }
    if ($Session.presentation) { $Session.presentation.Close(); [Runtime.InteropServices.Marshal]::ReleaseComObject($Session.presentation) | Out-Null }
    if ($Session.application) { $Session.application.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($Session.application) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
