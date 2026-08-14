<#
  Restricted PDF page renderer.

  The MCP server validates the page, input path, crop and output path before starting this script.
  Windows.Data.Pdf is built into Windows 10+, so end users do not need Node, Python, Poppler,
  or browser automation. This runtime source deliberately stays ASCII-only so Windows PowerShell
  5.1 remains safe even if an external packaging tool removes a UTF-8 BOM.
#>
param(
    [Parameter(Mandatory = $true)][string]$PdfPath,
    [Parameter(Mandatory = $true)][int]$PageNumber,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [ValidateRange(400, 3200)][int]$Width = 1600,
    [ValidateRange(400, 3200)][int]$Height = 2200,
    [double]$CropX = -1,
    [double]$CropY = -1,
    [double]$CropWidth = -1,
    [double]$CropHeight = -1
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Wait-WinRtOperation {
    param([object]$Operation, [type]$ResultType)

    # Windows PowerShell 5.1 cannot infer the WinRT generic extension method automatically.
    # Close the generic method by reflection and wait for IAsyncOperation<T> explicitly.
    $method = [System.WindowsRuntimeSystemExtensions].GetMethods([System.Reflection.BindingFlags]'Public,Static') |
        Where-Object {
            $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and
            $_.GetGenericArguments().Count -eq 1 -and $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.IsGenericType -and
            $_.GetParameters()[0].ParameterType.GetGenericTypeDefinition().FullName -eq 'Windows.Foundation.IAsyncOperation`1'
        } | Select-Object -First 1
    if ($null -eq $method) { throw 'Windows Runtime IAsyncOperation adapter is unavailable.' }
    $task = $method.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    return $task.GetAwaiter().GetResult()
}

function Wait-WinRtAction {
    param([object]$Action)

    $method = [System.WindowsRuntimeSystemExtensions].GetMethods([System.Reflection.BindingFlags]'Public,Static') |
        Where-Object {
            $_.Name -eq 'AsTask' -and -not $_.IsGenericMethodDefinition -and
            $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.FullName -eq 'Windows.Foundation.IAsyncAction'
        } | Select-Object -First 1
    if ($null -eq $method) { throw 'Windows Runtime IAsyncAction adapter is unavailable.' }
    $task = $method.Invoke($null, @($Action))
    $task.GetAwaiter().GetResult()
}

function Test-RequestedCrop {
    param([double]$X, [double]$Y, [double]$RequestedWidth, [double]$RequestedHeight)

    $values = @($X, $Y, $RequestedWidth, $RequestedHeight)
    $hasCrop = @($values | Where-Object { $_ -ge 0 }).Count -gt 0
    if (-not $hasCrop) { return $false }
    if ($X -lt 0 -or $Y -lt 0 -or $RequestedWidth -le 0 -or $RequestedHeight -le 0 -or
        $X -ge 1 -or $Y -ge 1 -or $RequestedWidth -gt 1 -or $RequestedHeight -gt 1 -or
        ($X + $RequestedWidth) -gt 1 -or ($Y + $RequestedHeight) -gt 1) {
        throw 'Crop coordinates must be normalized values inside the PDF page: x/y >= 0, width/height > 0, and x + width / y + height <= 1.'
    }
    return $true
}

if ($PageNumber -lt 1) { throw 'PageNumber must be at least 1.' }
if (-not (Test-Path -LiteralPath $PdfPath -PathType Leaf)) { throw "PDF was not found: $PdfPath" }
if ([IO.Path]::GetExtension($PdfPath).ToLowerInvariant() -ne '.pdf') { throw 'PdfPath must end with .pdf.' }
if (Test-Path -LiteralPath $OutputPath) { throw "OutputPath already exists: $OutputPath" }

$pixelCount = [int64]$Width * [int64]$Height
if ($pixelCount -gt 8000000) { throw 'Requested page render exceeds the 8,000,000-pixel limit.' }
$useCrop = Test-RequestedCrop -X $CropX -Y $CropY -RequestedWidth $CropWidth -RequestedHeight $CropHeight
$createdOutput = $false
$page = $null
$outputStream = $null

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $storageFileType = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
    $pdfDocumentType = [Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]
    $renderOptionsType = [Windows.Data.Pdf.PdfPageRenderOptions,Windows.Data.Pdf,ContentType=WindowsRuntime]
    $randomStreamType = [Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
    $bitmapEncoderType = [Windows.Graphics.Imaging.BitmapEncoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]

    # Reserve the file with CreateNew before WinRT opens it; never overwrite an existing file.
    $parent = Split-Path -Parent $OutputPath
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw "Output directory was not found: $parent" }
    [IO.File]::Open($OutputPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None).Dispose()
    $createdOutput = $true

    $inputFile = Wait-WinRtOperation ($storageFileType::GetFileFromPathAsync($PdfPath)) $storageFileType
    $pdf = Wait-WinRtOperation ($pdfDocumentType::LoadFromFileAsync($inputFile)) $pdfDocumentType
    try {
        if ($PageNumber -gt [int]$pdf.PageCount) { throw "PageNumber $PageNumber is outside the PDF page range 1-$($pdf.PageCount)." }
        $page = $pdf.GetPage([uint32]($PageNumber - 1))
        $sourceWidth = [double]$page.Size.Width
        $sourceHeight = [double]$page.Size.Height
        if ($sourceWidth -le 0 -or $sourceHeight -le 0) { throw 'PDF page has an invalid render size.' }

        # Width/Height are bounding dimensions, not forced stretch dimensions. Preserve the
        # page or confirmed-crop aspect ratio so tables, plots and system diagrams are not distorted.
        $renderSourceWidth = if ($useCrop) { $sourceWidth * $CropWidth } else { $sourceWidth }
        $renderSourceHeight = if ($useCrop) { $sourceHeight * $CropHeight } else { $sourceHeight }
        $scale = [Math]::Min(([double]$Width / $renderSourceWidth), ([double]$Height / $renderSourceHeight))
        if ($scale -le 0) { throw 'PDF page crop has an invalid render scale.' }
        $renderWidth = [Math]::Max(1, [Math]::Min($Width, [int][Math]::Round($renderSourceWidth * $scale)))
        $renderHeight = [Math]::Max(1, [Math]::Min($Height, [int][Math]::Round($renderSourceHeight * $scale)))
        if (([int64]$renderWidth * [int64]$renderHeight) -gt 8000000) { throw 'Derived page render exceeds the 8,000,000-pixel limit.' }

        $options = [Activator]::CreateInstance($renderOptionsType)
        $options.DestinationWidth = [uint32]$renderWidth
        $options.DestinationHeight = [uint32]$renderHeight
        $options.BitmapEncoderId = $bitmapEncoderType::PngEncoderId
        if ($useCrop) {
            # Windows PowerShell 5.1 can instantiate this WinRT value type through New-Object,
            # while its ContentType=WindowsRuntime type literal is not consistently resolvable.
            $options.SourceRect = New-Object -TypeName 'Windows.Foundation.Rect' -ArgumentList @(
                $sourceWidth * $CropX,
                $sourceHeight * $CropY,
                $renderSourceWidth,
                $renderSourceHeight
            )
        }
        $outputFile = Wait-WinRtOperation ($storageFileType::GetFileFromPathAsync($OutputPath)) $storageFileType
        $outputStream = Wait-WinRtOperation ($outputFile.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) $randomStreamType
        Wait-WinRtAction ($page.RenderToStreamAsync($outputStream, $options))
        [void](Wait-WinRtOperation ($outputStream.FlushAsync()) ([bool]))
    } finally {
        if ($page) { $page.Dispose() }
        # PdfDocument is a WinRT value that does not expose IDisposable on every
        # supported Windows PowerShell host. Dispose it only when that contract exists.
        if ($pdf -and ($pdf -is [System.IDisposable])) { $pdf.Dispose() }
    }

    # Close the WinRT stream before re-opening the rendered PNG for integrity checks.
    # Keeping it open can retain an exclusive file handle on Windows.
    if ($outputStream) {
        $outputStream.Dispose()
        $outputStream = $null
    }
    $file = Get-Item -LiteralPath $OutputPath -Force
    if ($file.Length -lt 8) { throw 'Rendered PNG is unexpectedly empty.' }
    if ($file.Length -gt 25MB) { throw 'Rendered PNG exceeds the 25 MB limit.' }
    $signature = [IO.File]::ReadAllBytes($OutputPath)[0..7]
    if (($signature -join ',') -ne '137,80,78,71,13,10,26,10') { throw 'Renderer did not produce a PNG file.' }
    [Console]::Out.WriteLine((@{
        path = [IO.Path]::GetFullPath($OutputPath)
        page_number = $PageNumber
        width = $renderWidth
        height = $renderHeight
        crop_applied = $useCrop
        bytes = [int64]$file.Length
    } | ConvertTo-Json -Compress))
} catch {
    if ($createdOutput -and (Test-Path -LiteralPath $OutputPath)) {
        try { Remove-Item -LiteralPath $OutputPath -Force } catch { }
    }
    throw
} finally {
    if ($outputStream) { $outputStream.Dispose() }
}
