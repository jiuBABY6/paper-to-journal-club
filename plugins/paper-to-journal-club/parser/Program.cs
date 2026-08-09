// 自包含 Windows 论文解析器的发布源码。
//
// 解析器直接处理用户提供的 PDF，因此除了依赖库的修复版本外，还必须在进程内部设置
// 明确的资源预算。MCP 层的超时只是第二道防线，不能替代这里的字节、文本、页数和磁盘限制。
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

// MCP 服务通过 UTF-8 与子进程交换 JSON；显式设置可避免中文论文在 Windows 代码页下变成乱码。
Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
Console.InputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

// 输入与解析预算。所有数值都会通过 --build-info 暴露，便于发行构建和回归测试验证。
const long MaximumInputBytes = 100L * 1024 * 1024;
const int MaximumPdfPages = 200;
const int MaximumTextCharactersPerPage = 150_000;
const int MaximumExtractedTextCharacters = 1_500_000;
const int MaximumExtractedAssets = 60;
const int MaximumPngBytes = 20 * 1024 * 1024;
const long MaximumTotalAssetBytes = 50L * 1024 * 1024;
const long MinimumFreeDiskBytes = 128L * 1024 * 1024;
const int MaximumPathCharacters = 240;
const int MaximumAssetFileNameCharacters = 32;
const int TextReadBufferCharacters = 8_192;

if (args.Length == 1 && string.Equals(args[0], "--build-info", StringComparison.OrdinalIgnoreCase))
{
    // 发布脚本必须调用此命令，确保随包 EXE 的 PdfPig 版本和源码要求一致。
    WriteBuildInfo();
    return 0;
}

if (args.Length < 2 || args.Length > 3 || !new[] { "extract", "extract-package" }.Contains(args[0], StringComparer.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("Usage: paper-parser.exe extract <paper-path> | extract-package <paper-path> <asset-directory> | --build-info");
    return 2;
}

var command = args[0].ToLowerInvariant();
if ((command == "extract" && args.Length != 2) || (command == "extract-package" && args.Length != 3))
{
    return InvalidPackageUsage();
}

PreparedAssetDirectory? assetDirectory = null;
try
{
    var filePath = GetVerifiedInputPath(args[1]);
    var inputLength = new FileInfo(filePath).Length;
    if (inputLength > MaximumInputBytes)
    {
        throw new ResourceLimitExceededException($"Paper is larger than the {MaximumInputBytes / 1024 / 1024} MB safety limit.");
    }

    var extension = Path.GetExtension(filePath).ToLowerInvariant();
    if (extension != ".pdf" && extension != ".txt" && extension != ".md" && extension != ".tex")
    {
        Console.Error.WriteLine("Supported extensions: PDF, TXT, Markdown, TeX.");
        return 4;
    }

    if (string.Equals(command, "extract-package", StringComparison.OrdinalIgnoreCase))
    {
        assetDirectory = PrepareAssetDirectory(args[2]);
    }

    if (extension is ".txt" or ".md" or ".tex")
    {
        // 不使用 File.ReadAllText：即使原文件大小合法，仍必须逐块限制实际进入内存和 JSON 的字符数。
        var sourceText = ReadTextWithinLimit(filePath);
        if (assetDirectory is not null)
        {
            WritePackage(new PaperPackage(sourceText, new[] { new PagePackage(1, sourceText, Array.Empty<ImageAsset>()) }), assetDirectory.Path);
        }
        else
        {
            Console.Write(sourceText);
        }
        return 0;
    }

    var package = ExtractPdfPackage(filePath, assetDirectory);
    if (package.Text.Length < 120)
    {
        // 本次没有产生可交付的 package，删除仅由解析器创建的临时资产目录和图片。
        assetDirectory?.CleanupAfterFailure();
        Console.Error.WriteLine("The PDF has too little extractable text. It may be scanned or encrypted and needs OCR or a password-enabled parser.");
        return 5;
    }

    if (assetDirectory is not null)
    {
        WritePackage(package, assetDirectory.Path);
    }
    else
    {
        Console.Write(package.Text);
    }
    return 0;
}
catch (FileNotFoundException error)
{
    Console.Error.WriteLine(error.Message);
    return 3;
}
catch (ResourceLimitExceededException error)
{
    // 失败时只删除本次解析器新建或新写入的资产，不碰调用方原有的任何文件。
    assetDirectory?.CleanupAfterFailure();
    Console.Error.WriteLine(error.Message);
    return 7;
}
catch (Exception error)
{
    assetDirectory?.CleanupAfterFailure();
    Console.Error.WriteLine($"PDF extraction failed: {error.Message}");
    return 6;
}

static int InvalidPackageUsage()
{
    Console.Error.WriteLine("extract-package requires an asset directory.");
    return 2;
}

static string GetVerifiedInputPath(string requestedPath)
{
    var filePath = GetSafeAbsolutePath(requestedPath, "Paper path", reservedSuffixCharacters: 0);
    if (!File.Exists(filePath))
    {
        throw new FileNotFoundException($"File not found: {filePath}", filePath);
    }

    var attributes = File.GetAttributes(filePath);
    if ((attributes & FileAttributes.Directory) != 0)
    {
        throw new ResourceLimitExceededException("Paper path must point to a file, not a directory.");
    }
    if ((attributes & FileAttributes.ReparsePoint) != 0)
    {
        // 不跟随符号链接或其他重解析点，避免 CLI 被用于绕过上层路径权限策略。
        throw new ResourceLimitExceededException("Paper path may not be a reparse point.");
    }
    return filePath;
}

static string GetSafeAbsolutePath(string requestedPath, string label, int reservedSuffixCharacters)
{
    if (string.IsNullOrWhiteSpace(requestedPath))
    {
        throw new ResourceLimitExceededException($"{label} is required.");
    }

    try
    {
        if (!Path.IsPathFullyQualified(requestedPath))
        {
            throw new ResourceLimitExceededException($"{label} must be an absolute local path.");
        }

        // UNC、Win32 设备路径和扩展路径可能触发网络访问或绕过普通 Win32 路径语义；解析器只接受本地盘符路径。
        if (requestedPath.StartsWith(@"\\", StringComparison.Ordinal) || requestedPath.StartsWith("//", StringComparison.Ordinal) || requestedPath.StartsWith(@"\\?\", StringComparison.Ordinal) || requestedPath.StartsWith(@"\\.\", StringComparison.Ordinal))
        {
            throw new ResourceLimitExceededException($"{label} may not use UNC or device-path syntax.");
        }

        var fullPath = Path.GetFullPath(requestedPath);
        if (fullPath.Length > MaximumPathCharacters - reservedSuffixCharacters)
        {
            throw new ResourceLimitExceededException($"{label} exceeds the {MaximumPathCharacters}-character safety limit.");
        }
        return fullPath;
    }
    catch (ResourceLimitExceededException)
    {
        throw;
    }
    catch (Exception error) when (error is ArgumentException or NotSupportedException or PathTooLongException)
    {
        throw new ResourceLimitExceededException($"{label} is not a valid local path.", error);
    }
}

static PreparedAssetDirectory PrepareAssetDirectory(string requestedDirectory)
{
    var directoryPath = GetSafeAbsolutePath(requestedDirectory, "Asset output directory", MaximumAssetFileNameCharacters);
    var normalizedDirectory = TrimTrailingDirectorySeparators(directoryPath);
    var normalizedRoot = TrimTrailingDirectorySeparators(Path.GetPathRoot(directoryPath) ?? string.Empty);
    if (string.IsNullOrWhiteSpace(normalizedRoot) || string.Equals(normalizedDirectory, normalizedRoot, StringComparison.OrdinalIgnoreCase))
    {
        throw new ResourceLimitExceededException("Asset output directory may not be a drive root.");
    }

    if (File.Exists(directoryPath))
    {
        throw new ResourceLimitExceededException("Asset output directory must be a directory, not a file.");
    }

    EnsureNoReparsePointInExistingDirectoryChain(directoryPath);
    var existedBefore = Directory.Exists(directoryPath);
    if (existedBefore && Directory.EnumerateFileSystemEntries(directoryPath).Any())
    {
        // 解析器使用固定、可预测的文件名；只接受空目录才能保证不会覆盖其他文件。
        throw new ResourceLimitExceededException("Asset output directory must be new or empty.");
    }

    Directory.CreateDirectory(directoryPath);
    EnsureNoReparsePointInExistingDirectoryChain(directoryPath);
    EnsureFreeDiskSpace(directoryPath, 0);
    return new PreparedAssetDirectory(directoryPath, !existedBefore);
}

static void EnsureNoReparsePointInExistingDirectoryChain(string directoryPath)
{
    var current = new DirectoryInfo(directoryPath);
    while (current is not null)
    {
        if (current.Exists && (current.Attributes & FileAttributes.ReparsePoint) != 0)
        {
            throw new ResourceLimitExceededException($"Asset output directory may not be inside a reparse point: {current.FullName}");
        }
        current = current.Parent;
    }
}

static void EnsureFreeDiskSpace(string directoryPath, long pendingWriteBytes)
{
    try
    {
        var root = Path.GetPathRoot(directoryPath);
        if (string.IsNullOrWhiteSpace(root))
        {
            throw new ResourceLimitExceededException("Asset output directory has no valid drive root.");
        }

        var drive = new DriveInfo(root);
        if (!drive.IsReady)
        {
            throw new ResourceLimitExceededException("Asset output drive is not ready.");
        }

        var requiredBytes = checked(MinimumFreeDiskBytes + pendingWriteBytes);
        if (drive.AvailableFreeSpace < requiredBytes)
        {
            throw new ResourceLimitExceededException($"Asset output drive must keep at least {MinimumFreeDiskBytes / 1024 / 1024} MB free.");
        }
    }
    catch (ResourceLimitExceededException)
    {
        throw;
    }
    catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException)
    {
        throw new ResourceLimitExceededException("Could not verify free space for the asset output drive.", error);
    }
}

static string ReadTextWithinLimit(string filePath)
{
    var builder = new StringBuilder();
    var buffer = new char[TextReadBufferCharacters];
    using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
    using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: TextReadBufferCharacters, leaveOpen: false);

    while (true)
    {
        var count = reader.Read(buffer, 0, buffer.Length);
        if (count == 0) break;
        if (count > MaximumExtractedTextCharacters - builder.Length)
        {
            throw new ResourceLimitExceededException($"Extracted text exceeds the {MaximumExtractedTextCharacters}-character safety limit.");
        }
        builder.Append(buffer, 0, count);
    }
    return builder.ToString();
}

static PaperPackage ExtractPdfPackage(string filePath, PreparedAssetDirectory? assetDirectory)
{
    var pages = new List<PagePackage>();
    var documentText = new StringBuilder();
    var extractedAssetCount = 0;
    long extractedAssetBytes = 0;
    var assetsTruncated = false;

    using var document = PdfDocument.Open(filePath);
    foreach (var page in document.GetPages())
    {
        if (pages.Count >= MaximumPdfPages)
        {
            throw new ResourceLimitExceededException($"The PDF has more than the {MaximumPdfPages}-page safety limit.");
        }

        // ContentOrderTextExtractor 通常比 Page.Text 更适合论文正文；两种结果都要先过单页上限。
        var pageText = ContentOrderTextExtractor.GetText(page);
        if (string.IsNullOrWhiteSpace(pageText)) pageText = page.Text;
        pageText = NormalizePageTextWithinLimit(pageText ?? string.Empty, page.Number);
        AppendPageTextWithinLimit(documentText, pageText, page.Number);

        var assets = new List<ImageAsset>();
        if (assetDirectory is not null && !assetsTruncated)
        {
            foreach (var image in page.GetImages())
            {
                if (extractedAssetCount >= MaximumExtractedAssets)
                {
                    assetsTruncated = true;
                    break;
                }

                // 仅导出 PdfPig 可可靠转换的 PNG，避免把未知编码伪装成图片写入磁盘。
                if (!image.TryGetPng(out var png)) continue;
                if (png.Length < 4_096 || png.Length > MaximumPngBytes)
                {
                    assetsTruncated = true;
                    continue;
                }

                var nextTotalAssetBytes = checked(extractedAssetBytes + (long)png.Length);
                if (nextTotalAssetBytes > MaximumTotalAssetBytes)
                {
                    assetsTruncated = true;
                    break;
                }

                var id = $"page-{page.Number:D2}-image-{extractedAssetCount + 1:D2}";
                var targetPath = WriteAssetPng(assetDirectory, id, png);
                assets.Add(new ImageAsset(id, page.Number, targetPath, png.Length));
                extractedAssetCount++;
                extractedAssetBytes = nextTotalAssetBytes;
            }
        }

        pages.Add(new PagePackage(page.Number, pageText, assets));
    }

    return new PaperPackage(documentText.ToString(), pages, assetsTruncated);
}

static string NormalizePageTextWithinLimit(string value, int pageNumber)
{
    // 在 Regex 处理之前拒绝异常长的单页文本，避免恶意内容造成额外的大型临时字符串。
    if (value.Length > MaximumTextCharactersPerPage)
    {
        throw new ResourceLimitExceededException($"Page {pageNumber} text exceeds the {MaximumTextCharactersPerPage}-character safety limit.");
    }

    var normalized = Normalize(value);
    if (normalized.Length > MaximumTextCharactersPerPage)
    {
        throw new ResourceLimitExceededException($"Page {pageNumber} normalized text exceeds the {MaximumTextCharactersPerPage}-character safety limit.");
    }
    return normalized;
}

static void AppendPageTextWithinLimit(StringBuilder documentText, string pageText, int pageNumber)
{
    if (string.IsNullOrWhiteSpace(pageText)) return;
    var separatorLength = documentText.Length == 0 ? 0 : 2;
    if (pageText.Length > MaximumExtractedTextCharacters - documentText.Length - separatorLength)
    {
        throw new ResourceLimitExceededException($"Extracted text exceeds the {MaximumExtractedTextCharacters}-character safety limit at page {pageNumber}.");
    }

    if (separatorLength > 0) documentText.Append("\n\n");
    documentText.Append(pageText);
}

static string WriteAssetPng(PreparedAssetDirectory assetDirectory, string id, byte[] png)
{
    EnsureNoReparsePointInExistingDirectoryChain(assetDirectory.Path);
    EnsureFreeDiskSpace(assetDirectory.Path, png.Length);

    var targetPath = Path.GetFullPath(Path.Combine(assetDirectory.Path, $"{id}.png"));
    var directoryPrefix = TrimTrailingDirectorySeparators(assetDirectory.Path) + Path.DirectorySeparatorChar;
    if (!targetPath.StartsWith(directoryPrefix, StringComparison.OrdinalIgnoreCase) || targetPath.Length > MaximumPathCharacters)
    {
        throw new ResourceLimitExceededException("Generated asset path is outside the permitted output directory or too long.");
    }
    if (File.Exists(targetPath) || Directory.Exists(targetPath))
    {
        throw new ResourceLimitExceededException($"Refusing to overwrite an existing asset: {targetPath}");
    }

    // CreateNew 在同名文件被其他进程抢先创建时失败，避免 Time-of-check/Time-of-use 覆盖。
    using (var output = new FileStream(targetPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
    {
        output.Write(png, 0, png.Length);
    }
    assetDirectory.RegisterWrittenAsset(targetPath);
    return targetPath;
}

static void WritePackage(PaperPackage package, string assetDirectory)
{
    var result = new
    {
        text = package.Text,
        asset_directory = assetDirectory,
        pages = package.Pages,
        assets_truncated = package.AssetsTruncated
    };
    Console.Write(JsonSerializer.Serialize(result));
}

static void WriteBuildInfo()
{
    var result = new
    {
        format_version = 1,
        pdfpig_version = typeof(PdfDocument).Assembly.GetName().Version?.ToString() ?? "unknown",
        limits = new
        {
            maximum_input_bytes = MaximumInputBytes,
            maximum_pdf_pages = MaximumPdfPages,
            maximum_text_characters_per_page = MaximumTextCharactersPerPage,
            maximum_extracted_text_characters = MaximumExtractedTextCharacters,
            maximum_extracted_assets = MaximumExtractedAssets,
            maximum_png_bytes = MaximumPngBytes,
            maximum_total_asset_bytes = MaximumTotalAssetBytes,
            minimum_free_disk_bytes = MinimumFreeDiskBytes,
            maximum_path_characters = MaximumPathCharacters
        }
    };
    Console.Write(JsonSerializer.Serialize(result));
}

static string Normalize(string value)
{
    return Regex.Replace(value.Replace("\r", ""), "[ \\t]+", " ")
        .Replace("\n \n", "\n\n")
        .Trim();
}

static string TrimTrailingDirectorySeparators(string path)
{
    return path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
}

// InvalidDataException 在当前 .NET 目标框架中是 sealed，不能作为领域异常的基类。
// IOException 仍能准确表达不安全/无效的本地输入或输出状态，且保留内层异常与退出码 7 的
// 专用处理语义。
internal sealed class ResourceLimitExceededException : IOException
{
    public ResourceLimitExceededException(string message) : base(message) { }
    public ResourceLimitExceededException(string message, Exception innerException) : base(message, innerException) { }
}

internal sealed class PreparedAssetDirectory
{
    private readonly List<string> writtenAssets = new();

    public PreparedAssetDirectory(string path, bool wasCreatedByParser)
    {
        Path = path;
        WasCreatedByParser = wasCreatedByParser;
    }

    public string Path { get; }
    public bool WasCreatedByParser { get; }

    public void RegisterWrittenAsset(string path)
    {
        writtenAssets.Add(path);
    }

    public void CleanupAfterFailure()
    {
        // 仅处理已经通过 CreateNew 写入且由当前实例登记的路径；清理失败不掩盖原始解析错误。
        foreach (var assetPath in writtenAssets.AsEnumerable().Reverse())
        {
            try
            {
                if (File.Exists(assetPath)) File.Delete(assetPath);
            }
            catch
            {
                // 无法删除时由调用方查看目录；绝不扩大删除范围。
            }
        }

        if (!WasCreatedByParser) return;
        try
        {
            if (Directory.Exists(Path) && !Directory.EnumerateFileSystemEntries(Path).Any())
            {
                Directory.Delete(Path, recursive: false);
            }
        }
        catch
        {
            // 不因清理临时空目录失败而覆盖原始错误。
        }
    }
}

internal sealed record PaperPackage(string Text, IEnumerable<PagePackage> Pages, bool AssetsTruncated = false);
internal sealed record PagePackage(int PageNumber, string Text, IEnumerable<ImageAsset> Assets);
internal sealed record ImageAsset(string Id, int PageNumber, string Path, long Bytes);
