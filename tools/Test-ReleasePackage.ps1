<#
  对已经构建好的本地 Marketplace 发行目录做离线验证。

  本脚本不会启动 PowerPoint、不会联网，也不会注册 Codex，因此可直接在 CI 和发布机使用。
  它验证 Marketplace 布局、MCP 入口、随包解析器、脚本语法、禁止的开发产物和 SHA-256 清单。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$MarketplaceRoot,

    # 仅供构建中尚未生成 SHA256SUMS.txt 的临时检查；正式发行不得使用。
    [switch]$SkipChecksumCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PluginName = 'paper-to-journal-club'
$MarketplaceName = 'paper-to-journal-club-tools'
$ParserBuildInfoTimeoutMilliseconds = 15000
$MaximumCapturedProcessOutputBytes = 65536
$ProcessOutputDrainTimeoutMilliseconds = 5000
$ExpectedParserLimits = [ordered]@{
    maximum_input_bytes = 104857600
    maximum_pdf_pages = 200
    maximum_text_characters_per_page = 150000
    maximum_extracted_text_characters = 1500000
    maximum_extracted_assets = 60
    maximum_png_bytes = 20971520
    maximum_total_asset_bytes = 52428800
    minimum_free_disk_bytes = 134217728
    maximum_path_characters = 240
}

function Get-RequiredProperty {
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Context
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        throw "$Context 缺少必填字段 '$Name'。"
    }
    return $property.Value
}

function Assert-ExactPropertyNames {
    <#
      发行包的 JSON 清单属于可执行入口的一部分。除了检查关键字段外，也拒绝额外字段，
      这样恶意或误打包的配置不能偷偷声明第二个插件、第二个 MCP 服务或额外启动参数。
    #>
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string[]]$AllowedNames,
        [Parameter(Mandatory)][string]$Context
    )

    $actualNames = @($Object.PSObject.Properties | ForEach-Object { $_.Name })
    $unexpectedNames = @($actualNames | Where-Object { $AllowedNames -notcontains $_ })
    $missingNames = @($AllowedNames | Where-Object { $actualNames -notcontains $_ })
    if ($unexpectedNames.Count -gt 0 -or $missingNames.Count -gt 0) {
        $details = @()
        if ($unexpectedNames.Count -gt 0) {
            $details += "未允许字段：$($unexpectedNames -join ', ')"
        }
        if ($missingNames.Count -gt 0) {
            $details += "缺少字段：$($missingNames -join ', ')"
        }
        throw "$Context 字段不在允许清单内：$($details -join '；')"
    }
}

function ConvertTo-CanonicalRelativePath {
    <# 将清单与实际文件统一为 ZIP 使用的正斜杠相对路径，并拒绝路径穿越与 Windows 特殊路径。 #>
    param([Parameter(Mandatory)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or $Path.IndexOf([char]0) -ge 0) {
        throw 'SHA-256 清单包含空路径或 NUL 字符。'
    }
    $normalized = $Path.Replace('\', '/').Trim()
    if ($normalized.StartsWith('/') -or $normalized.EndsWith('/') -or $normalized.Contains(':')) {
        throw "SHA-256 清单含不安全路径：$Path"
    }
    $segments = @($normalized -split '/')
    if ($segments.Count -eq 0 -or @($segments | Where-Object { $_ -in @('', '.', '..') }).Count -gt 0) {
        throw "SHA-256 清单含不安全路径：$Path"
    }
    return $normalized
}

function Get-RelativePath {
    <# Windows PowerShell 5.1 的 .NET Framework 没有 System.IO.Path.GetRelativePath。 #>
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$TargetPath
    )

    $baseUri = [System.Uri]::new((Join-Path $BasePath ''))
    $targetUri = [System.Uri]::new($TargetPath)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString()).Replace('\', '/')
}

function Test-PortableExecutable {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $false
    }
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        if ($stream.Length -lt 2) {
            return $false
        }
        return $stream.ReadByte() -eq 0x4D -and $stream.ReadByte() -eq 0x5A
    } finally {
        $stream.Dispose()
    }
}

function Initialize-BoundedProcessRunner {
    <#
      同时排空 stdout 和 stderr，避免子进程交替写两个管道时，调用方串行 ReadToEnd()
      造成死锁。读取器只保留有限字节用于诊断，但会继续排空其余输出，避免不可信
      二进制以海量输出耗尽验证进程内存。

      此帮助类仅依赖 Windows PowerShell 自带的 .NET Framework，不需要 .NET SDK；同一
      PowerShell 会话重复运行验证脚本时复用已加载的类型，避免 Add-Type 重复定义报错。
    #>
    if ($null -ne ('PaperToJournalClub.ReleaseValidation.BoundedProcessRunner' -as [type])) {
        return
    }

    Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace PaperToJournalClub.ReleaseValidation
{
    public sealed class BoundedProcessResult
    {
        public int ExitCode { get; set; }
        public bool TimedOut { get; set; }
        public bool TerminationRequested { get; set; }
        public string StandardOutput { get; set; }
        public string StandardError { get; set; }
        public long StandardOutputTotalBytes { get; set; }
        public long StandardErrorTotalBytes { get; set; }
        public bool StandardOutputTruncated { get; set; }
        public bool StandardErrorTruncated { get; set; }
    }

    internal sealed class BoundedReadResult
    {
        public byte[] CapturedBytes;
        public long TotalBytes;
        public bool WasTruncated;
    }

    public static class BoundedProcessRunner
    {
        private const int BufferSize = 8192;

        public static BoundedProcessResult Run(
            string fileName,
            string arguments,
            int timeoutMilliseconds,
            int maximumCapturedBytes,
            int drainTimeoutMilliseconds)
        {
            if (String.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("子进程路径不能为空。", "fileName");
            }
            if (timeoutMilliseconds <= 0 || maximumCapturedBytes <= 0 || drainTimeoutMilliseconds <= 0)
            {
                throw new ArgumentOutOfRangeException("进程超时和输出上限必须为正数。");
            }

            var startInfo = new ProcessStartInfo();
            startInfo.FileName = fileName;
            startInfo.Arguments = arguments ?? String.Empty;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.StandardOutputEncoding = new UTF8Encoding(false);
            startInfo.StandardErrorEncoding = new UTF8Encoding(false);

            using (var process = new Process())
            {
                process.StartInfo = startInfo;
                try
                {
                    if (!process.Start())
                    {
                        throw new InvalidOperationException("Process.Start 返回 false。");
                    }
                }
                catch (Exception error)
                {
                    throw new InvalidOperationException("无法启动子进程：" + fileName + "。", error);
                }

                Stream standardOutputStream = process.StandardOutput.BaseStream;
                Stream standardErrorStream = process.StandardError.BaseStream;
                Task<BoundedReadResult> standardOutputTask = ReadBoundedAsync(standardOutputStream, maximumCapturedBytes);
                Task<BoundedReadResult> standardErrorTask = ReadBoundedAsync(standardErrorStream, maximumCapturedBytes);
                bool timedOut = !process.WaitForExit(timeoutMilliseconds);
                bool terminationRequested = false;

                if (timedOut)
                {
                    try
                    {
                        if (!process.HasExited)
                        {
                            process.Kill();
                            terminationRequested = true;
                        }
                    }
                    catch (InvalidOperationException)
                    {
                        // 进程恰好在超时边界自行退出，无需再终止。
                    }

                    if (!process.WaitForExit(drainTimeoutMilliseconds))
                    {
                        CloseQuietly(standardOutputStream);
                        CloseQuietly(standardErrorStream);
                        throw new TimeoutException(
                            "子进程在 " + timeoutMilliseconds + " ms 内未退出；已请求终止，但在额外 "
                            + drainTimeoutMilliseconds + " ms 内仍未结束。");
                    }
                }

                try
                {
                    if (!Task.WaitAll(new Task[] { standardOutputTask, standardErrorTask }, drainTimeoutMilliseconds))
                    {
                        CloseQuietly(standardOutputStream);
                        CloseQuietly(standardErrorStream);
                        throw new TimeoutException(
                            "子进程已结束，但 stdout/stderr 在 " + drainTimeoutMilliseconds
                            + " ms 内未能排空；已关闭输出管道。");
                    }
                }
                catch (AggregateException error)
                {
                    throw new InvalidOperationException("读取子进程 stdout/stderr 时失败。", error.Flatten());
                }

                BoundedReadResult standardOutput = standardOutputTask.Result;
                BoundedReadResult standardError = standardErrorTask.Result;
                return new BoundedProcessResult
                {
                    ExitCode = process.ExitCode,
                    TimedOut = timedOut,
                    TerminationRequested = terminationRequested,
                    StandardOutput = Encoding.UTF8.GetString(standardOutput.CapturedBytes),
                    StandardError = Encoding.UTF8.GetString(standardError.CapturedBytes),
                    StandardOutputTotalBytes = standardOutput.TotalBytes,
                    StandardErrorTotalBytes = standardError.TotalBytes,
                    StandardOutputTruncated = standardOutput.WasTruncated,
                    StandardErrorTruncated = standardError.WasTruncated
                };
            }
        }

        private static async Task<BoundedReadResult> ReadBoundedAsync(Stream stream, int maximumCapturedBytes)
        {
            var buffer = new byte[BufferSize];
            var captured = new MemoryStream(Math.Min(BufferSize, maximumCapturedBytes));
            long totalBytes = 0;
            bool wasTruncated = false;

            try
            {
                while (true)
                {
                    int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                    if (bytesRead == 0)
                    {
                        break;
                    }

                    totalBytes += bytesRead;
                    int remaining = maximumCapturedBytes - (int)captured.Length;
                    int bytesToCapture = Math.Min(remaining, bytesRead);
                    if (bytesToCapture > 0)
                    {
                        captured.Write(buffer, 0, bytesToCapture);
                    }
                    if (bytesToCapture < bytesRead)
                    {
                        wasTruncated = true;
                    }
                }

                return new BoundedReadResult
                {
                    CapturedBytes = captured.ToArray(),
                    TotalBytes = totalBytes,
                    WasTruncated = wasTruncated
                };
            }
            finally
            {
                captured.Dispose();
            }
        }

        private static void CloseQuietly(Stream stream)
        {
            try
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
            catch
            {
                // 终止路径中只需尽力关闭管道，原始超时错误更有诊断价值。
            }
        }
    }
}
'@
}

function Invoke-BoundedProcess {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][int]$TimeoutMilliseconds,
        [Parameter(Mandatory)][int]$MaximumCapturedBytes,
        [Parameter(Mandatory)][int]$DrainTimeoutMilliseconds
    )

    Initialize-BoundedProcessRunner
    return [PaperToJournalClub.ReleaseValidation.BoundedProcessRunner]::Run(
        $FilePath,
        $Arguments,
        $TimeoutMilliseconds,
        $MaximumCapturedBytes,
        $DrainTimeoutMilliseconds
    )
}

function Get-ProcessOutputDiagnostic {
    <# 将失败输出限制在小段文字，防止验证失败本身在终端生成大量日志。 #>
    param([Parameter(Mandatory)]$Result)

    $details = @(
        "stdout=$($Result.StandardOutputTotalBytes) bytes",
        "stderr=$($Result.StandardErrorTotalBytes) bytes"
    )
    if ($Result.StandardOutputTruncated) { $details += 'stdout 已截断' }
    if ($Result.StandardErrorTruncated) { $details += 'stderr 已截断' }
    $standardError = [string]$Result.StandardError
    if (-not [string]::IsNullOrWhiteSpace($standardError)) {
        if ($standardError.Length -gt 2048) {
            $standardError = $standardError.Substring(0, 2048) + '…'
        }
        $details += "stderr 内容：$standardError"
    }
    return ($details -join '；')
}

function Test-ParserBuildInfo {
    <#
      MZ 头只能证明文件“看起来像 EXE”。正式发行还必须由解析器自身报告已修复的 PdfPig
      版本与资源预算，防止源码已经升级但 assets/paper-parser.exe 仍是旧二进制。
    #>
    param([Parameter(Mandatory)][string]$Path)

    $processResult = Invoke-BoundedProcess `
        -FilePath $Path `
        -Arguments '--build-info' `
        -TimeoutMilliseconds $ParserBuildInfoTimeoutMilliseconds `
        -MaximumCapturedBytes $MaximumCapturedProcessOutputBytes `
        -DrainTimeoutMilliseconds $ProcessOutputDrainTimeoutMilliseconds
    if ($processResult.TimedOut) {
        throw "paper-parser.exe --build-info 在 $ParserBuildInfoTimeoutMilliseconds ms 内未完成，已请求终止：$(Get-ProcessOutputDiagnostic -Result $processResult)"
    }
    if ($processResult.StandardOutputTruncated -or $processResult.StandardErrorTruncated) {
        throw "paper-parser.exe --build-info 输出超过 $MaximumCapturedProcessOutputBytes bytes 的安全上限：$(Get-ProcessOutputDiagnostic -Result $processResult)"
    }
    if ($processResult.ExitCode -ne 0) {
        throw "paper-parser.exe --build-info 失败（退出码 $($processResult.ExitCode)）：$(Get-ProcessOutputDiagnostic -Result $processResult)"
    }
    try {
        $buildInfo = $processResult.StandardOutput | ConvertFrom-Json
    } catch {
        throw "paper-parser.exe --build-info 未返回有效 JSON：$(Get-ProcessOutputDiagnostic -Result $processResult)"
    }
    if ([int]$buildInfo.format_version -ne 1 -or [string]$buildInfo.pdfpig_version -notmatch '^0\.1\.15(\.|$)') {
        throw "paper-parser.exe 不是受支持的安全构建（PdfPig=$($buildInfo.pdfpig_version)）。"
    }
    foreach ($limitName in $ExpectedParserLimits.Keys) {
        $property = $buildInfo.limits.PSObject.Properties[$limitName]
        if ($null -eq $property -or [int64]$property.Value -ne [int64]$ExpectedParserLimits[$limitName]) {
            throw "paper-parser.exe 的资源预算不符合发行要求：$limitName"
        }
    }
}

function Test-ParserDependencyLock {
    <#
      发行包保留解析器源码时，也必须保留由受控 SDK 实际生成的完整依赖锁。这里不重算
      依赖图：只验证锁文件与 csproj 的关键约束，并由构建脚本的 --locked-mode 最终消费它。
    #>
    param(
        [Parameter(Mandatory)][string]$ProjectPath,
        [Parameter(Mandatory)][string]$LockPath
    )

    if (-not (Test-Path -LiteralPath $ProjectPath -PathType Leaf) -or -not (Test-Path -LiteralPath $LockPath -PathType Leaf)) {
        throw '发行包缺少解析器项目文件或 packages.lock.json。'
    }
    $projectText = Get-Content -LiteralPath $ProjectPath -Raw -Encoding UTF8
    foreach ($requiredFragment in @(
        '<RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>',
        '<RestoreLockedMode>true</RestoreLockedMode>',
        'Include="PdfPig" Version="[0.1.15]"'
    )) {
        if (-not $projectText.Contains($requiredFragment)) {
            throw "解析器项目缺少依赖锁定约束：$requiredFragment"
        }
    }
    try {
        $lock = Get-Content -LiteralPath $LockPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "packages.lock.json 不是有效 JSON：$($_.Exception.Message)"
    }
    if ($null -eq $lock -or [int]$lock.version -ne 1 -or $null -eq $lock.dependencies) {
        throw 'packages.lock.json 缺少 NuGet lock file version 1 或 dependencies。'
    }

    # RuntimeIdentifier 会使 NuGet 把键写为 net8.0/win-x64；接受该合法变体，但不接受其它目标。
    $pdfPigEntries = @()
    foreach ($targetProperty in @($lock.dependencies.PSObject.Properties | Where-Object { $_.Name -like 'net8.0*' })) {
        $pdfPigProperty = $targetProperty.Value.PSObject.Properties['PdfPig']
        if ($null -ne $pdfPigProperty -and $null -ne $pdfPigProperty.Value) {
            $pdfPigEntries += [pscustomobject]@{ target = $targetProperty.Name; value = $pdfPigProperty.Value }
        }
    }
    if ($pdfPigEntries.Count -eq 0) {
        throw 'packages.lock.json 未在 net8.0（含 RID）目标中锁定 PdfPig。'
    }
    foreach ($entry in $pdfPigEntries) {
        $pdfPig = $entry.value
        # requested 的文本序列化在 NuGet SDK 间可以不同；精确请求范围由 csproj
        # 固定，这里验证直接依赖、非空 requested 与锁定后的精确 resolved 版本。
        if ([string]$pdfPig.type -cne 'Direct' -or [string]::IsNullOrWhiteSpace([string]$pdfPig.requested) -or [string]$pdfPig.resolved -cne '0.1.15') {
            throw "packages.lock.json 的 PdfPig 条目不符合精确直接依赖要求：$($entry.target)"
        }
        # NuGet lock file 使用原始 Base64 SHA-512，而非带算法前缀的 SRI 字符串。
        $hashMatch = [regex]::Match([string]$pdfPig.contentHash, '^(?<base64>[A-Za-z0-9+/]+={0,2})$')
        if (-not $hashMatch.Success) {
            throw "packages.lock.json 的 PdfPig contentHash 无效：$($entry.target)"
        }
        try {
            if ([Convert]::FromBase64String($hashMatch.Groups['base64'].Value).Length -ne 64) {
                throw 'unexpected SHA-512 byte count'
            }
        } catch {
            throw "packages.lock.json 的 PdfPig contentHash 不是 SHA-512：$($entry.target)"
        }
    }
}

function Test-ChecksumManifest {
    param([Parameter(Mandatory)][string]$Root)

    $checksumPath = Join-Path $Root 'SHA256SUMS.txt'
    if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) {
        throw "未找到 SHA-256 清单：$checksumPath"
    }
    $lines = @(Get-Content -LiteralPath $checksumPath -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) {
        throw 'SHA-256 清单为空。'
    }

    # 双向集合校验：清单中的每个文件必须存在，解压目录中除清单自身外的每个文件也必须被清单覆盖。
    $manifestEntries = @{}
    foreach ($line in $lines) {
        $match = [System.Text.RegularExpressions.Regex]::Match($line, '^(?<hash>[A-Fa-f0-9]{64})  \*(?<path>.+)$')
        if (-not $match.Success) {
            throw "SHA-256 清单格式错误：$line"
        }
        $relativePath = ConvertTo-CanonicalRelativePath -Path $match.Groups['path'].Value
        if ($manifestEntries.ContainsKey($relativePath)) {
            throw "SHA-256 清单包含重复路径：$relativePath"
        }
        $manifestEntries[$relativePath] = $match.Groups['hash'].Value.ToUpperInvariant()
    }

    $actualEntries = @{}
    foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
        if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "发行包不得包含重解析点文件：$($file.FullName)"
        }
        $relativePath = Get-RelativePath -BasePath $Root -TargetPath $file.FullName
        if ($relativePath -eq 'SHA256SUMS.txt') {
            continue
        }
        $relativePath = ConvertTo-CanonicalRelativePath -Path $relativePath
        if ($actualEntries.ContainsKey($relativePath)) {
            throw "发行包包含规范化后重复的文件路径：$relativePath"
        }
        $actualEntries[$relativePath] = $file.FullName
    }

    $missingEntries = @($manifestEntries.Keys | Where-Object { -not $actualEntries.ContainsKey($_) })
    $unexpectedEntries = @($actualEntries.Keys | Where-Object { -not $manifestEntries.ContainsKey($_) })
    if ($missingEntries.Count -gt 0 -or $unexpectedEntries.Count -gt 0) {
        $details = @()
        if ($missingEntries.Count -gt 0) {
            $details += "清单列出但实际缺失：$($missingEntries -join ', ')"
        }
        if ($unexpectedEntries.Count -gt 0) {
            $details += "实际存在但未列入清单：$($unexpectedEntries -join ', ')"
        }
        throw "SHA-256 清单双向集合校验失败：$($details -join '；')"
    }

    foreach ($relativePath in $manifestEntries.Keys) {
        $actual = (Get-FileHash -LiteralPath $actualEntries[$relativePath] -Algorithm SHA256).Hash.ToUpperInvariant()
        if ($actual -ne $manifestEntries[$relativePath]) {
            throw "SHA-256 不匹配：$relativePath"
        }
    }
}

function Test-PowerShellSyntax {
    param([Parameter(Mandatory)][string]$Root)

    $syntaxErrors = @()
    $utf8WithoutBom = [System.Text.UTF8Encoding]::new($false)
    foreach ($scriptFile in Get-ChildItem -LiteralPath $Root -Recurse -Filter '*.ps1' -File) {
        $sourceText = [System.IO.File]::ReadAllText($scriptFile.FullName, $utf8WithoutBom)
        $tokens = $null
        $parseErrors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($sourceText, [ref]$tokens, [ref]$parseErrors)
        foreach ($parseError in $parseErrors) {
            $syntaxErrors += "{0}:{1}:{2} {3}" -f $scriptFile.FullName, $parseError.Extent.StartLineNumber, $parseError.Extent.StartColumnNumber, $parseError.Message
        }
    }
    if ($syntaxErrors.Count -gt 0) {
        throw "PowerShell AST 校验失败：`n$($syntaxErrors -join [Environment]::NewLine)"
    }
}

$root = (Resolve-Path -LiteralPath $MarketplaceRoot -ErrorAction Stop).Path
$marketplacePath = Join-Path $root '.agents\plugins\marketplace.json'
if (-not (Test-Path -LiteralPath $marketplacePath -PathType Leaf)) {
    throw "未找到 Marketplace 清单：$marketplacePath"
}

$marketplace = Get-Content -LiteralPath $marketplacePath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-ExactPropertyNames -Object $marketplace -AllowedNames @('name', 'interface', 'plugins') -Context 'Marketplace 清单'
if ((Get-RequiredProperty $marketplace 'name' 'Marketplace 清单') -ne $MarketplaceName) {
    throw "Marketplace 名称必须为 '$MarketplaceName'。"
}
$interface = Get-RequiredProperty $marketplace 'interface' 'Marketplace 清单'
Assert-ExactPropertyNames -Object $interface -AllowedNames @('displayName') -Context 'Marketplace interface'
if ([string]::IsNullOrWhiteSpace([string](Get-RequiredProperty $interface 'displayName' 'Marketplace interface'))) {
    throw 'Marketplace interface.displayName 不能为空。'
}

$allEntries = @($marketplace.plugins)
if ($allEntries.Count -ne 1) {
    throw "Marketplace 只能包含一个 '$PluginName' 条目，实际为 $($allEntries.Count) 个。"
}
$entry = @($allEntries | Where-Object { $_.name -eq $PluginName })
if ($entry.Count -ne 1) {
    throw "Marketplace 必须恰好包含一个 '$PluginName' 条目。"
}
$entry = $entry[0]
Assert-ExactPropertyNames -Object $entry -AllowedNames @('name', 'source', 'policy', 'category') -Context 'Marketplace 插件条目'
$source = Get-RequiredProperty $entry 'source' 'Marketplace 插件条目'
$policy = Get-RequiredProperty $entry 'policy' 'Marketplace 插件条目'
Assert-ExactPropertyNames -Object $source -AllowedNames @('source', 'path') -Context 'Marketplace source'
Assert-ExactPropertyNames -Object $policy -AllowedNames @('installation', 'authentication') -Context 'Marketplace policy'
if ((Get-RequiredProperty $source 'source' 'Marketplace source') -ne 'local') {
    throw 'Marketplace source.source 必须为 local。'
}
if ((Get-RequiredProperty $source 'path' 'Marketplace source') -ne "./plugins/$PluginName") {
    throw "Marketplace source.path 必须为 './plugins/$PluginName'。"
}
if ((Get-RequiredProperty $policy 'installation' 'Marketplace policy') -ne 'AVAILABLE') {
    throw 'Marketplace policy.installation 必须为 AVAILABLE。'
}
if ((Get-RequiredProperty $policy 'authentication' 'Marketplace policy') -ne 'ON_INSTALL') {
    throw 'Marketplace policy.authentication 必须为 ON_INSTALL。'
}
[void](Get-RequiredProperty $entry 'category' 'Marketplace 插件条目')

$pluginRoot = Join-Path $root "plugins\$PluginName"
$manifestPath = Join-Path $pluginRoot '.codex-plugin\plugin.json'
$mcpPath = Join-Path $pluginRoot '.mcp.json'
$serverPath = Join-Path $pluginRoot 'scripts\paper-to-journal-club-server.ps1'
$parserPath = Join-Path $pluginRoot 'assets\paper-parser.exe'
$parserProjectPath = Join-Path $pluginRoot 'parser\PaperParser.csproj'
$parserPackageLockPath = Join-Path $pluginRoot 'parser\packages.lock.json'
$requiredPaths = @(
    $pluginRoot,
    $manifestPath,
    $mcpPath,
    $serverPath,
    $parserPath,
    (Join-Path $pluginRoot 'skills'),
    (Join-Path $pluginRoot 'THIRD_PARTY_NOTICES.md'),
    $parserProjectPath,
    (Join-Path $pluginRoot 'parser\NuGet.Config'),
    $parserPackageLockPath
)
foreach ($requiredPath in $requiredPaths) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "发行包缺少所需文件：$requiredPath"
    }
}

foreach ($forbiddenRelativePath in @('parser\bin', 'parser\obj', 'release', 'examples', 'tests', '.git', 'assets\parser-publish')) {
    $forbiddenPath = Join-Path $pluginRoot $forbiddenRelativePath
    if (Test-Path -LiteralPath $forbiddenPath) {
        throw "发行包包含禁止的开发产物：$forbiddenPath"
    }
}

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-ExactPropertyNames -Object $manifest -AllowedNames @('name', 'version', 'description', 'author', 'license', 'keywords', 'skills', 'mcpServers', 'interface') -Context '插件清单'
if ((Get-RequiredProperty $manifest 'name' '插件清单') -ne $PluginName) {
    throw "插件清单名称必须为 '$PluginName'。"
}
$version = [string](Get-RequiredProperty $manifest 'version' '插件清单')
if ($version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$') {
    throw "插件版本不是严格 SemVer：$version"
}
[void](Get-RequiredProperty $manifest 'description' '插件清单')
$author = Get-RequiredProperty $manifest 'author' '插件清单'
Assert-ExactPropertyNames -Object $author -AllowedNames @('name') -Context '插件作者'
[void](Get-RequiredProperty $author 'name' '插件作者')
$pluginInterface = Get-RequiredProperty $manifest 'interface' '插件清单'
Assert-ExactPropertyNames -Object $pluginInterface -AllowedNames @('displayName', 'shortDescription', 'longDescription', 'developerName', 'category', 'capabilities', 'defaultPrompt') -Context '插件 interface'
foreach ($propertyName in @('displayName', 'shortDescription', 'longDescription', 'developerName', 'category', 'capabilities', 'defaultPrompt')) {
    [void](Get-RequiredProperty $pluginInterface $propertyName '插件 interface')
}
if ($manifest.license -ne 'MIT' -or $manifest.skills -ne './skills/' -or $manifest.mcpServers -ne './.mcp.json' -or @($manifest.keywords).Count -eq 0 -or @($pluginInterface.capabilities).Count -eq 0 -or @($pluginInterface.defaultPrompt).Count -eq 0) {
    throw '插件清单不在允许的本地技能/MCP 配置范围内。'
}

$mcp = Get-Content -LiteralPath $mcpPath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-ExactPropertyNames -Object $mcp -AllowedNames @('mcpServers') -Context 'MCP 配置'
Assert-ExactPropertyNames -Object $mcp.mcpServers -AllowedNames @($PluginName) -Context 'MCP server 清单'
$server = $mcp.mcpServers.$PluginName
if ($null -eq $server) {
    throw "未找到 MCP server '$PluginName'。"
}
Assert-ExactPropertyNames -Object $server -AllowedNames @('command', 'args', 'cwd') -Context 'MCP server'
if ($server.command -ne 'powershell.exe') {
    throw 'MCP server 必须由 powershell.exe 启动。'
}
$expectedMcpArguments = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', './scripts/paper-to-journal-club-server.ps1')
$actualMcpArguments = @($server.args | ForEach-Object { [string]$_ })
if ($actualMcpArguments.Count -ne $expectedMcpArguments.Count) {
    throw 'MCP server 参数必须精确为 -NoProfile、RemoteSigned 和本地 paper-to-journal-club-server.ps1 入口。'
}
for ($index = 0; $index -lt $expectedMcpArguments.Count; $index++) {
    if ($actualMcpArguments[$index] -cne $expectedMcpArguments[$index]) {
        throw 'MCP server 参数顺序或内容不在允许清单内。'
    }
}
if ($server.cwd -ne '.') {
    throw 'MCP server.cwd 必须为插件根目录 (.)。'
}
if (-not (Test-PortableExecutable -Path $parserPath)) {
    throw '随包 paper-parser.exe 缺失或不是有效的 Windows 可执行文件。'
}
Test-ParserBuildInfo -Path $parserPath
Test-ParserDependencyLock -ProjectPath $parserProjectPath -LockPath $parserPackageLockPath

Test-PowerShellSyntax -Root $root
if (-not $SkipChecksumCheck) {
    Test-ChecksumManifest -Root $root
}

[pscustomobject]@{
    pass = $true
    marketplace = $MarketplaceName
    plugin = $PluginName
    version = $version
    parser = $parserPath
    node_required = $false
} | ConvertTo-Json -Depth 4
