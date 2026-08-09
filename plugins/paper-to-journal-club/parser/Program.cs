// 自包含 Windows PDF 文本提取器的发布源码。
// 运行时随插件一起发布为单文件 exe，因此终端用户不需要 Node、Python 或 .NET。
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

// MCP 服务以 UTF-8 读取子进程 JSON；显式设置可避免中文论文在不同 Windows 代码页下变成乱码。
Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
Console.InputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

// 解析器面对用户提供的 PDF，必须有资源上限，防止异常文件耗尽本机内存或磁盘。
const long MaximumInputBytes = 100L * 1024 * 1024;
const int MaximumPdfPages = 200;
const int MaximumExtractedAssets = 60;
const int MaximumPngBytes = 20 * 1024 * 1024;

if (args.Length < 2 || args.Length > 3 || !new[] { "extract", "extract-package" }.Contains(args[0], StringComparer.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("Usage: paper-parser.exe extract <paper-path> | extract-package <paper-path> <asset-directory>");
    return 2;
}

var command = args[0].ToLowerInvariant();
if ((command == "extract" && args.Length != 2) || (command == "extract-package" && args.Length != 3))
{
    return InvalidPackageUsage();
}

var filePath = Path.GetFullPath(args[1]);
if (!File.Exists(filePath))
{
    Console.Error.WriteLine($"File not found: {filePath}");
    return 3;
}
if (new FileInfo(filePath).Length > MaximumInputBytes)
{
    Console.Error.WriteLine($"Paper is larger than the {MaximumInputBytes / 1024 / 1024} MB safety limit.");
    return 7;
}

var extension = Path.GetExtension(filePath).ToLowerInvariant();
if (extension is ".txt" or ".md" or ".tex")
{
    var sourceText = File.ReadAllText(filePath);
    if (string.Equals(command, "extract-package", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length != 3) return InvalidPackageUsage();
        WritePackage(new PaperPackage(sourceText, new[] { new PagePackage(1, sourceText, Array.Empty<ImageAsset>()) }), args[2]);
    }
    else Console.Write(sourceText);
    return 0;
}
if (extension != ".pdf")
{
    Console.Error.WriteLine("Supported extensions: PDF, TXT, Markdown, TeX.");
    return 4;
}

try
{
    var package = ExtractPdfPackage(filePath, command, args.Length == 3 ? args[2] : null);
    var text = package.Text;
    if (text.Length < 120)
    {
        Console.Error.WriteLine("The PDF has too little extractable text. It may be scanned or encrypted and needs OCR or a password-enabled parser.");
        return 5;
    }
    if (string.Equals(command, "extract-package", StringComparison.OrdinalIgnoreCase))
    {
        if (args.Length != 3) return InvalidPackageUsage();
        WritePackage(package, args[2]);
    }
    else Console.Write(text);
    return 0;
}
catch (Exception error)
{
    Console.Error.WriteLine($"PDF extraction failed: {error.Message}");
    return 6;
}

static int InvalidPackageUsage()
{
    Console.Error.WriteLine("extract-package requires an asset directory.");
    return 2;
}

static PaperPackage ExtractPdfPackage(string filePath, string command, string? assetDirectory)
{
    var pages = new List<PagePackage>();
    var extractAssets = string.Equals(command, "extract-package", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(assetDirectory);
    var extractedAssetCount = 0;
    var assetsTruncated = false;
    if (extractAssets) Directory.CreateDirectory(assetDirectory!);
    using var document = PdfDocument.Open(filePath);
    foreach (var page in document.GetPages())
    {
        if (page.Number > MaximumPdfPages)
        {
            throw new InvalidDataException($"The PDF has more than the {MaximumPdfPages}-page safety limit.");
        }
        // ContentOrderTextExtractor provides a more readable order than Page.Text for most papers.
        var pageText = ContentOrderTextExtractor.GetText(page);
        if (string.IsNullOrWhiteSpace(pageText)) pageText = page.Text;
        var assets = new List<ImageAsset>();
        if (extractAssets)
        {
            var imageIndex = 0;
            foreach (var image in page.GetImages())
            {
                if (extractedAssetCount >= MaximumExtractedAssets)
                {
                    assetsTruncated = true;
                    break;
                }
                // Only retain a lossless PNG when PdfPig can produce one. Unknown encodings are ignored,
                // rather than being written with a misleading extension or silently corrupted.
                if (!image.TryGetPng(out var png) || png.Length < 4096 || png.Length > MaximumPngBytes) continue;
                imageIndex++;
                var id = $"page-{page.Number:D2}-image-{imageIndex:D2}";
                var targetPath = Path.Combine(assetDirectory!, $"{id}.png");
                File.WriteAllBytes(targetPath, png);
                assets.Add(new ImageAsset(id, page.Number, Path.GetFullPath(targetPath), png.Length));
                extractedAssetCount++;
            }
        }
        pages.Add(new PagePackage(page.Number, Normalize(pageText), assets));
    }
    return new PaperPackage(string.Join("\n\n", pages.Select(page => page.Text).Where(page => !string.IsNullOrWhiteSpace(page))).Trim(), pages, assetsTruncated);
}

static void WritePackage(PaperPackage package, string assetDirectory)
{
    var result = new
    {
        text = package.Text,
        asset_directory = Path.GetFullPath(assetDirectory),
        pages = package.Pages,
        assets_truncated = package.AssetsTruncated
    };
    Console.Write(JsonSerializer.Serialize(result));
}

static string Normalize(string value)
{
    return Regex.Replace(value.Replace("\r", ""), "[ \\t]+", " ")
        .Replace("\n \n", "\n\n")
        .Trim();
}

internal sealed record PaperPackage(string Text, IEnumerable<PagePackage> Pages, bool AssetsTruncated = false);
internal sealed record PagePackage(int PageNumber, string Text, IEnumerable<ImageAsset> Assets);
internal sealed record ImageAsset(string Id, int PageNumber, string Path, int Bytes);
