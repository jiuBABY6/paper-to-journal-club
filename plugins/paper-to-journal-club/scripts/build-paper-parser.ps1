<#
  发布阶段执行一次：生成 assets/paper-parser.exe。

  该 EXE 为 win-x64 self-contained single file，最终用户无需安装 Node 或 .NET。
  本脚本不仅编译，还会启动刚生成的 EXE 检查 PdfPig 版本和资源预算；这样源码升级后，
  旧二进制不能被误打进 GitHub Release。
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$pluginRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $pluginRoot 'parser\PaperParser.csproj'
$nugetConfig = Join-Path $pluginRoot 'parser\NuGet.Config'
$packageLock = Join-Path $pluginRoot 'parser\packages.lock.json'
$runtimeTest = Join-Path $pluginRoot 'tests\parser-resource-limit-tests.ps1'
$output = Join-Path $pluginRoot 'assets\paper-parser.exe'
$dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
$buildRoot = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club-parser-build-$([Guid]::NewGuid().ToString('N'))"
$publishDirectory = Join-Path $buildRoot 'publish'
$intermediateDirectory = (Join-Path $buildRoot 'obj') + [IO.Path]::DirectorySeparatorChar
$binaryDirectory = (Join-Path $buildRoot 'bin') + [IO.Path]::DirectorySeparatorChar
$replacementId = [Guid]::NewGuid().ToString('N')
$temporaryOutput = "$output.$replacementId.tmp"
# 使用同目录、唯一的备份路径实现可恢复的替换；避免 PowerShell 调用 .NET File.Replace
# 时将空备份路径误绑定为空字符串，也兼容 Windows PowerShell 5.1。
$temporaryBackup = "$output.$replacementId.backup"

function Assert-RequiredFile {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Label)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label does not exist: $Path"
    }
}

function Test-WindowsExecutable {
    param([Parameter(Mandatory)][string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    try {
        return (($stream.Length -ge 2) -and ($stream.ReadByte() -eq 0x4D) -and ($stream.ReadByte() -eq 0x5A))
    } finally {
        $stream.Dispose()
    }
}

Assert-RequiredFile -Path $project -Label 'Parser project'
Assert-RequiredFile -Path $nugetConfig -Label 'NuGet configuration'
# packages.lock.json 只能由受控 SDK 的生成脚本写入并提交；发布构建绝不临时重算依赖图。
Assert-RequiredFile -Path $packageLock -Label 'Committed NuGet dependency lock'
Assert-RequiredFile -Path $runtimeTest -Label 'Parser resource-limit test'

if ($dotnetCommand) {
    $dotnetPath = $dotnetCommand.Source
} else {
    throw 'dotnet SDK is required only for release builders. End users receive the compiled paper-parser.exe.'
}

# CI 与发布者生成锁文件都固定在同一 SDK，避免 SDK 自身的还原行为改变锁定依赖。
$dotnetVersion = (& $dotnetPath --version).Trim()
if ($LASTEXITCODE -ne 0 -or $dotnetVersion -ne '8.0.418') {
    throw "Parser release builds require .NET SDK 8.0.418, actual: $dotnetVersion"
}

# 构建缓存全部隔离在带 GUID 的系统临时目录内，避免污染仓库或复用开发机的 NuGet HTTP 缓存。
$env:DOTNET_CLI_HOME = Join-Path $buildRoot 'dotnet-home'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
$env:NUGET_PACKAGES = Join-Path $buildRoot 'nuget-packages'
$env:NUGET_HTTP_CACHE_PATH = Join-Path $buildRoot 'nuget-http-cache'
$env:NUGET_PLUGINS_CACHE_PATH = Join-Path $buildRoot 'nuget-plugins-cache'

try {
    New-Item -ItemType Directory -Force -Path $env:DOTNET_CLI_HOME, $env:NUGET_PACKAGES, $env:NUGET_HTTP_CACHE_PATH, $env:NUGET_PLUGINS_CACHE_PATH, $publishDirectory | Out-Null

    # 先独立 restore。--locked-mode 禁止在发布时重算依赖图；若锁文件与项目不一致则直接失败。
    # --disable-parallel 让审计日志和失败点更容易复现，且避免并发下载放大网络资源占用。
    # dotnet restore 使用 --no-cache 清空 HTTP 缓存影响；不能沿用 NuGet 其他命令的
    # 参数拼写，否则会被底层 MSBuild 作为未知开关拒绝。
    & $dotnetPath restore $project --configfile $nugetConfig --runtime win-x64 --locked-mode --disable-parallel --no-cache "-p:BaseIntermediateOutputPath=$intermediateDirectory" "-p:BaseOutputPath=$binaryDirectory"
    if ($LASTEXITCODE -ne 0) { throw 'paper-parser dependency restore failed.' }

    & $dotnetPath publish $project -c Release -r win-x64 --self-contained true --no-restore -p:PublishSingleFile=true "-p:BaseIntermediateOutputPath=$intermediateDirectory" "-p:BaseOutputPath=$binaryDirectory" -o $publishDirectory
    if ($LASTEXITCODE -ne 0) { throw 'paper-parser build failed.' }

    $builtExecutable = Join-Path $publishDirectory 'PaperParser.exe'
    Assert-RequiredFile -Path $builtExecutable -Label 'Built parser executable'
    if (-not (Test-WindowsExecutable -Path $builtExecutable)) {
        throw "paper-parser build completed without a valid Windows executable: $builtExecutable"
    }

    # 这是对真正即将复制的 EXE 的黑盒测试，不接受仅修改源码却遗留旧二进制的情况。
    & $runtimeTest -ParserPath $builtExecutable -RequireRuntime -RequireLock
    if ($LASTEXITCODE -ne 0) { throw 'Built parser failed resource-limit verification.' }

    # 先复制到同目录临时文件，验证完成后再替换旧 EXE。先把旧文件移动到唯一备份，
    # 新文件移动失败时立即恢复旧文件；成功后才删除备份。
    [IO.File]::Copy($builtExecutable, $temporaryOutput, $false)
    if (Test-Path -LiteralPath $output -PathType Leaf) {
        Move-Item -LiteralPath $output -Destination $temporaryBackup
        try {
            Move-Item -LiteralPath $temporaryOutput -Destination $output
        } catch {
            $replacementError = $_
            if ((Test-Path -LiteralPath $temporaryBackup -PathType Leaf) -and -not (Test-Path -LiteralPath $output -PathType Leaf)) {
                try {
                    Move-Item -LiteralPath $temporaryBackup -Destination $output
                } catch {
                    Write-Warning "Unable to restore the previous parser executable from ${temporaryBackup}: $($_.Exception.Message)"
                }
            }
            throw $replacementError
        }
        Remove-Item -LiteralPath $temporaryBackup -Force
    } else {
        Move-Item -LiteralPath $temporaryOutput -Destination $output
    }
    Write-Output "Built and verified $output"
}
finally {
    # 仅删除本脚本创建的带 GUID 临时目录和未完成的新 EXE，不影响源码或已发布解析器。
    # 如果替换/恢复均失败，保留备份文件以免丢失旧的解析器；构建会失败，发布不会继续。
    if (Test-Path -LiteralPath $temporaryOutput) {
        Remove-Item -LiteralPath $temporaryOutput -Force
    }
    if (Test-Path -LiteralPath $buildRoot) {
        Remove-Item -LiteralPath $buildRoot -Recurse -Force
    }
}
