<#
  PowerPoint COM quality helpers shared by generation and read-only inspection.

  The checks deliberately distinguish deterministic structural faults from visual
  review items. A PPTX is not approved merely because SaveAs succeeded.
#>

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
