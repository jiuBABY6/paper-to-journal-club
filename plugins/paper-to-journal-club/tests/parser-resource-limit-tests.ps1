<#
  PDF 解析器资源预算回归测试。

  本测试先静态检查源码和依赖锁定；若给出新构建的 EXE，再检查 --build-info 以及两个
  无需真实 PDF 的运行时拒绝路径。发布构建必须传入 -RequireRuntime，确保不能把旧的
  0.1.13 二进制误打进发行包。开发机没有 .NET SDK 时，默认仅报告源码检查通过并明确提示
  现有 EXE 仍需由发布机构建。
#>
[CmdletBinding()]
param(
    # Windows PowerShell 在 param 默认值求值时尚未初始化 $PSScriptRoot，故在脚本体内再赋默认值。
    [string]$ParserPath,
    [switch]$RequireRuntime,

    # 仅正式构建和发行打包传入。源码回归允许尚未生成锁文件的开发分支继续做静态检查。
    [switch]$RequireLock
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$DefaultProcessTimeoutMilliseconds = 15000
$MaximumCapturedProcessOutputBytes = 65536
$ProcessOutputDrainTimeoutMilliseconds = 5000
$pluginRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$ParserPath = if ([string]::IsNullOrWhiteSpace($ParserPath)) { Join-Path $pluginRoot 'assets\paper-parser.exe' } else { $ParserPath }
$projectPath = Join-Path $pluginRoot 'parser\PaperParser.csproj'
$programPath = Join-Path $pluginRoot 'parser\Program.cs'
$nugetConfigPath = Join-Path $pluginRoot 'parser\NuGet.Config'
$packageLockPath = Join-Path $pluginRoot 'parser\packages.lock.json'
$expectedLimits = [ordered]@{
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

function Assert-Condition {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}

function Test-ParserPackageLock {
    <#
      只接受 NuGet restore 实际生成的锁文件的最小可审阅结构。这里不凭空写入任何
      contentHash：发布构建还会用 --locked-mode 消费该文件，生成工具则会把它和下载的
      nupkg SHA-512 逐字节比对。
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Required
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $message = "未找到 NuGet 锁文件：$Path。开发/源码检查可以继续；正式构建和发行打包必须先用 tools/Generate-ParserPackageLock.ps1 生成并提交真实 packages.lock.json。"
        if ($Required) {
            throw $message
        }
        Write-Warning $message
        return $false
    }

    try {
        $lock = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "NuGet 锁文件不是有效 JSON：$Path。$($_.Exception.Message)"
    }

    Assert-Condition ($null -ne $lock -and [int]$lock.version -eq 1) 'packages.lock.json 必须使用 NuGet lock file format version 1.'
    $dependenciesProperty = $lock.PSObject.Properties['dependencies']
    Assert-Condition ($null -ne $dependenciesProperty -and $null -ne $dependenciesProperty.Value) 'packages.lock.json 缺少 dependencies。'
    # 带 RuntimeIdentifier 的 restore 可能写成 net8.0/win-x64，不能把正确的锁文件
    # 误判为无效；但每个 net8.0 目标中出现的 PdfPig 都必须是同一精确直接依赖。
    $pdfPigEntries = @()
    foreach ($frameworkProperty in @($dependenciesProperty.Value.PSObject.Properties | Where-Object { $_.Name -like 'net8.0*' })) {
        $pdfPigProperty = $frameworkProperty.Value.PSObject.Properties['PdfPig']
        if ($null -ne $pdfPigProperty -and $null -ne $pdfPigProperty.Value) {
            $pdfPigEntries += [pscustomobject]@{ target = $frameworkProperty.Name; value = $pdfPigProperty.Value }
        }
    }
    Assert-Condition ($pdfPigEntries.Count -gt 0) 'packages.lock.json 必须在 net8.0（含 RID）目标中锁定 PdfPig 直接依赖。'

    foreach ($entry in $pdfPigEntries) {
        $pdfPig = $entry.value
        Assert-Condition ([string]$pdfPig.type -ceq 'Direct') "PdfPig 必须是锁文件中的直接依赖：$($entry.target)"
        # NuGet 在不同 SDK 上会把同一 PackageReference 的 requested 字段序列化为
        # 不同但等价的格式（例如 0.1.15 或 [0.1.15]）。精确请求约束由 csproj 的
        # PackageReference 负责；锁文件在这里必须锁住同一个 resolved 版本和内容哈希。
        Assert-Condition (-not [string]::IsNullOrWhiteSpace([string]$pdfPig.requested)) "PdfPig 锁文件必须保留 requested 版本信息：$($entry.target)"
        Assert-Condition ([string]$pdfPig.resolved -ceq '0.1.15') "PdfPig 锁文件解析版本必须精确为 0.1.15：$($entry.target)"
        $contentHash = [string]$pdfPig.contentHash
        $hashMatch = [regex]::Match($contentHash, '^sha512-(?<base64>[A-Za-z0-9+/]+={0,2})$')
        Assert-Condition $hashMatch.Success "PdfPig 锁文件 contentHash 不是 NuGet sha512 格式：$($entry.target)"
        try {
            $decodedContentHash = [Convert]::FromBase64String($hashMatch.Groups['base64'].Value)
        } catch {
            throw "PdfPig 锁文件 contentHash 不是有效 Base64 SHA-512 值：$($entry.target)"
        }
        Assert-Condition ($decodedContentHash.Length -eq 64) "PdfPig 锁文件 contentHash 必须是 SHA-512（64 字节）：$($entry.target)"
    }
    return $true
}

function Initialize-BoundedProcessRunner {
    <#
      解析器会同时使用 stdout 与 stderr。本帮助类并发排空两条管道，同时只保留有限
      字节用于断言和诊断，避免串行 ReadToEnd() 的死锁，以及异常二进制输出导致测试
      进程占满内存。它只使用 Windows PowerShell 自带运行时，无需安装 .NET SDK。
    #>
    if ($null -ne ('PaperToJournalClub.ParserTests.BoundedProcessRunner' -as [type])) {
        return
    }

    Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace PaperToJournalClub.ParserTests
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
                        // 进程刚好在超时边界退出，无需再终止。
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
                // 终止路径只需尽力关闭管道，保留原始超时错误方便定位。
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
    return [PaperToJournalClub.ParserTests.BoundedProcessRunner]::Run(
        $FilePath,
        $Arguments,
        $TimeoutMilliseconds,
        $MaximumCapturedBytes,
        $DrainTimeoutMilliseconds
    )
}

function Invoke-ParserProcess {
    param(
        [Parameter(Mandatory)][string]$Executable,
        [Parameter(Mandatory)][string]$Arguments,
        [int]$TimeoutMilliseconds = $DefaultProcessTimeoutMilliseconds,
        [int]$MaximumCapturedBytes = $MaximumCapturedProcessOutputBytes
    )

    $result = Invoke-BoundedProcess `
        -FilePath $Executable `
        -Arguments $Arguments `
        -TimeoutMilliseconds $TimeoutMilliseconds `
        -MaximumCapturedBytes $MaximumCapturedBytes `
        -DrainTimeoutMilliseconds $ProcessOutputDrainTimeoutMilliseconds
    return [pscustomobject]@{
        exit_code = $result.ExitCode
        timed_out = $result.TimedOut
        termination_requested = $result.TerminationRequested
        standard_output = $result.StandardOutput
        standard_error = $result.StandardError
        standard_output_total_bytes = $result.StandardOutputTotalBytes
        standard_error_total_bytes = $result.StandardErrorTotalBytes
        standard_output_truncated = $result.StandardOutputTruncated
        standard_error_truncated = $result.StandardErrorTruncated
    }
}

function Get-ParserProcessDiagnostic {
    <# 保留退出码、输出量和少量 stderr，既便于定位又不会把异常输出原样铺满日志。 #>
    param([Parameter(Mandatory)]$Result)

    $details = @(
        "stdout=$($Result.standard_output_total_bytes) bytes",
        "stderr=$($Result.standard_error_total_bytes) bytes"
    )
    if ($Result.standard_output_truncated) { $details += 'stdout 已截断' }
    if ($Result.standard_error_truncated) { $details += 'stderr 已截断' }
    $standardError = [string]$Result.standard_error
    if (-not [string]::IsNullOrWhiteSpace($standardError)) {
        if ($standardError.Length -gt 2048) {
            $standardError = $standardError.Substring(0, 2048) + '…'
        }
        $details += "stderr 内容：$standardError"
    }
    return ($details -join '；')
}

function ConvertTo-EncodedPowerShellCommand {
    <# 用 -EncodedCommand 构造临时子进程，无需在磁盘留下未签名脚本，也不依赖 .NET SDK。 #>
    param([Parameter(Mandatory)][string]$Command)

    return [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Command))
}

function Test-BoundedProcessRunnerWithoutDotNet {
    <#
      回归覆盖两个 P2 风险：
      1. 子进程交替写 stdout/stderr 且超过捕获上限时，读取器仍能并发排空而不死锁；
      2. 子进程一直运行时，硬超时会终止进程而不是无限等待。
    #>
    $powershellPath = Join-Path $PSHOME 'powershell.exe'
    Assert-Condition (Test-Path -LiteralPath $powershellPath -PathType Leaf) 'Windows PowerShell executable is required for process-runner regression tests.'

    $largeOutputCommand = @'
$chunk = [string]::new([char]120, 8192)
for ($index = 0; $index -lt 10; $index++) {
    [Console]::Out.Write($chunk)
    [Console]::Error.Write($chunk)
}
'@
    $largeOutputArguments = '-NoProfile -NonInteractive -EncodedCommand {0}' -f (ConvertTo-EncodedPowerShellCommand -Command $largeOutputCommand)
    $largeOutputResult = Invoke-ParserProcess `
        -Executable $powershellPath `
        -Arguments $largeOutputArguments `
        -TimeoutMilliseconds 10000 `
        -MaximumCapturedBytes 4096
    Assert-Condition (-not $largeOutputResult.timed_out) 'Concurrent stdout/stderr drain must complete before the hard timeout.'
    Assert-Condition ($largeOutputResult.exit_code -eq 0) 'Large stdout/stderr test process must exit successfully.'
    Assert-Condition $largeOutputResult.standard_output_truncated 'stdout must be bounded when child output exceeds the capture limit.'
    Assert-Condition $largeOutputResult.standard_error_truncated 'stderr must be bounded when child output exceeds the capture limit.'
    Assert-Condition ($largeOutputResult.standard_output_total_bytes -ge 81920) 'stdout must be fully drained after capture is truncated.'
    Assert-Condition ($largeOutputResult.standard_error_total_bytes -ge 81920) 'stderr must be fully drained after capture is truncated.'
    Assert-Condition ($largeOutputResult.standard_output.Length -le 4096) 'Captured stdout must not exceed its byte budget for ASCII output.'
    Assert-Condition ($largeOutputResult.standard_error.Length -le 4096) 'Captured stderr must not exceed its byte budget for ASCII output.'

    $sleepArguments = '-NoProfile -NonInteractive -EncodedCommand {0}' -f (ConvertTo-EncodedPowerShellCommand -Command 'Start-Sleep -Seconds 20')
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $timeoutResult = Invoke-ParserProcess `
        -Executable $powershellPath `
        -Arguments $sleepArguments `
        -TimeoutMilliseconds 700 `
        -MaximumCapturedBytes 4096
    $stopwatch.Stop()
    Assert-Condition $timeoutResult.timed_out 'A long-running child process must report a timeout.'
    Assert-Condition $timeoutResult.termination_requested 'A timed-out child process must receive a termination request.'
    Assert-Condition ($stopwatch.ElapsedMilliseconds -lt 7000) 'The timeout path must return within the configured timeout plus bounded cleanup time.'
}

foreach ($requiredPath in @($projectPath, $programPath, $nugetConfigPath)) {
    Assert-Condition (Test-Path -LiteralPath $requiredPath -PathType Leaf) "Missing parser source file: $requiredPath"
}

# 依赖必须精确固定到安全修复版本，不能因浮动版本在没有审阅的情况下升级或降级。
$projectText = [IO.File]::ReadAllText($projectPath, $utf8)
Assert-Condition ($projectText -match '<PackageReference\s+Include="PdfPig"\s+Version="\[0\.1\.15\]"\s*/>') 'PdfPig must be pinned exactly to [0.1.15].'
Assert-Condition ($projectText -match '<NuGetAudit>true</NuGetAudit>') 'NuGet audit must be enabled for parser builds.'
Assert-Condition ($projectText -match 'NU1901;NU1902;NU1903;NU1904') 'NuGet vulnerability warnings must fail the parser build.'
Assert-Condition ($projectText -match '<RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>') 'Parser build must opt in to a committed NuGet lock file.'
Assert-Condition ($projectText -match '<RestoreLockedMode>true</RestoreLockedMode>') 'Parser release builds must use NuGet locked mode.'

# 发布构建与正式打包将 RequireLock 置为 true；普通源码审查仍能明确看到未生成锁的提示，
# 但不会伪造一个不可审计的 JSON 文件来掩盖该前置条件。
[void](Test-ParserPackageLock -Path $packageLockPath -Required:$RequireLock)

$nugetConfigText = [IO.File]::ReadAllText($nugetConfigPath, $utf8)
Assert-Condition ($nugetConfigText -match '<clear\s*/>') 'NuGet.Config must clear user-level package sources.'
Assert-Condition ($nugetConfigText -match 'https://api\.nuget\.org/v3/index\.json') 'NuGet.Config must use the official NuGet v3 source.'

$programText = [IO.File]::ReadAllText($programPath, $utf8)
foreach ($requiredToken in @('--build-info', 'MaximumExtractedTextCharacters', 'MaximumTotalAssetBytes', 'MinimumFreeDiskBytes', 'MaximumPathCharacters', 'FileMode.CreateNew', 'ResourceLimitExceededException')) {
    Assert-Condition ($programText.Contains($requiredToken)) "Program.cs must contain parser resource control: $requiredToken"
}

# 本回归在检查真实 EXE 前执行，因此即便开发机尚未安装 .NET SDK，也会验证进程读取的
# 并发排空、输出上限和硬超时语义。
Test-BoundedProcessRunnerWithoutDotNet

$resolvedParserPath = [IO.Path]::GetFullPath($ParserPath)
if (-not (Test-Path -LiteralPath $resolvedParserPath -PathType Leaf)) {
    if ($RequireRuntime) { throw "Parser executable was not found: $resolvedParserPath" }
    Write-Warning "仅完成源码检查：未找到解析器 EXE。正式 Release 必须运行 build-paper-parser.ps1 重新构建。"
    Write-Output 'PASS: parser-resource-limit-tests.ps1 (source only)'
    exit 0
}

$buildInfoProcess = Invoke-ParserProcess -Executable $resolvedParserPath -Arguments '--build-info'
if ($buildInfoProcess.timed_out -or $buildInfoProcess.standard_output_truncated -or $buildInfoProcess.standard_error_truncated -or $buildInfoProcess.exit_code -ne 0) {
    if ($buildInfoProcess.timed_out) {
        $failureReason = "timed out after $DefaultProcessTimeoutMilliseconds ms and received a termination request"
    } elseif ($buildInfoProcess.standard_output_truncated -or $buildInfoProcess.standard_error_truncated) {
        $failureReason = "exceeded the $MaximumCapturedProcessOutputBytes-byte diagnostic output limit"
    } else {
        $failureReason = "exited with code $($buildInfoProcess.exit_code)"
    }
    $message = "Parser EXE does not expose secure --build-info ($failureReason): $(Get-ParserProcessDiagnostic -Result $buildInfoProcess)"
    if ($RequireRuntime) { throw $message }
    Write-Warning "$message`n当前随包 EXE 仍是旧构建，不能作为已修复发行物；请在有 .NET SDK 的发布机执行 build-paper-parser.ps1。"
    Write-Output 'PASS: parser-resource-limit-tests.ps1 (source only; binary rebuild required)'
    exit 0
}

try {
    $buildInfo = $buildInfoProcess.standard_output | ConvertFrom-Json
} catch {
    throw "Parser --build-info did not return valid JSON: $(Get-ParserProcessDiagnostic -Result $buildInfoProcess)"
}

Assert-Condition ([int]$buildInfo.format_version -eq 1) 'Parser build-info format_version must be 1.'
Assert-Condition ([string]$buildInfo.pdfpig_version -match '^0\.1\.15(\.|$)') "Parser EXE must contain PdfPig 0.1.15, actual: $($buildInfo.pdfpig_version)"
foreach ($limitName in $expectedLimits.Keys) {
    $actualProperty = $buildInfo.limits.PSObject.Properties[$limitName]
    Assert-Condition ($null -ne $actualProperty) "Parser build-info is missing limit: $limitName"
    Assert-Condition ([int64]$actualProperty.Value -eq [int64]$expectedLimits[$limitName]) "Parser limit $limitName has an unexpected value."
}

# 运行时测试只生成临时文本文件和空资产目录，不需要额外下载真实 PDF 样本。
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club-parser-test-$([Guid]::NewGuid().ToString('N'))"
try {
    New-Item -ItemType Directory -Force -Path $temporaryRoot | Out-Null
    $oversizedTextPath = Join-Path $temporaryRoot 'oversized.txt'
    [IO.File]::WriteAllText($oversizedTextPath, ('x' * ($expectedLimits.maximum_extracted_text_characters + 1)), $utf8)
    $oversizedResult = Invoke-ParserProcess -Executable $resolvedParserPath -Arguments ('extract "{0}"' -f $oversizedTextPath)
    Assert-Condition (-not $oversizedResult.timed_out) "Oversized text test must not time out: $(Get-ParserProcessDiagnostic -Result $oversizedResult)"
    Assert-Condition (-not $oversizedResult.standard_output_truncated -and -not $oversizedResult.standard_error_truncated) "Oversized text test output must stay within the diagnostic limit: $(Get-ParserProcessDiagnostic -Result $oversizedResult)"
    Assert-Condition ($oversizedResult.exit_code -eq 7) 'Oversized text input must fail with the resource-limit exit code.'
    Assert-Condition ($oversizedResult.standard_error -match 'Extracted text exceeds') 'Oversized text input must report the text resource limit.'

    $shortTextPath = Join-Path $temporaryRoot 'short.txt'
    [IO.File]::WriteAllText($shortTextPath, ('safe content ' * 20), $utf8)
    $nonEmptyAssetDirectory = Join-Path $temporaryRoot 'non-empty-assets'
    New-Item -ItemType Directory -Force -Path $nonEmptyAssetDirectory | Out-Null
    [IO.File]::WriteAllText((Join-Path $nonEmptyAssetDirectory 'existing.txt'), 'do not overwrite', $utf8)
    $assetResult = Invoke-ParserProcess -Executable $resolvedParserPath -Arguments ('extract-package "{0}" "{1}"' -f $shortTextPath, $nonEmptyAssetDirectory)
    Assert-Condition (-not $assetResult.timed_out) "Asset directory test must not time out: $(Get-ParserProcessDiagnostic -Result $assetResult)"
    Assert-Condition (-not $assetResult.standard_output_truncated -and -not $assetResult.standard_error_truncated) "Asset directory test output must stay within the diagnostic limit: $(Get-ParserProcessDiagnostic -Result $assetResult)"
    Assert-Condition ($assetResult.exit_code -eq 7) 'A non-empty asset directory must fail with the resource-limit exit code.'
    Assert-Condition ($assetResult.standard_error -match 'new or empty') 'A non-empty asset directory must not be accepted.'
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        # 仅删除本测试创建的 GUID 临时目录。
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}

Write-Output 'PASS: parser-resource-limit-tests.ps1'
