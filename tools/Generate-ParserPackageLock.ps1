<#
  在受控发布环境中生成 parser/packages.lock.json。

  此脚本只供插件维护者或只读 GitHub Actions 使用。它不会提交、推送或发布任何内容；
  生成后的锁文件必须由维护者审阅并手工提交，之后正式构建才会以 --locked-mode 消费它。
#>
[CmdletBinding()]
param(
    # 已有锁文件时必须显式允许重新计算，避免依赖图在不知情时改变。
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$RepositoryRoot = Split-Path -Parent $PSScriptRoot
$PluginRoot = Join-Path $RepositoryRoot 'plugins\paper-to-journal-club'
$ProjectPath = Join-Path $PluginRoot 'parser\PaperParser.csproj'
$NuGetConfigPath = Join-Path $PluginRoot 'parser\NuGet.Config'
$LockPath = Join-Path $PluginRoot 'parser\packages.lock.json'
$ParserTestPath = Join-Path $PluginRoot 'tests\parser-resource-limit-tests.ps1'
$RequiredSdkVersion = '8.0.418'
$dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("paper-to-journal-club-lock-$([Guid]::NewGuid().ToString('N'))")

function Assert-RequiredFile {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Label)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label does not exist: $Path"
    }
}

foreach ($required in @(
    @{ path = $ProjectPath; label = 'Parser project' },
    @{ path = $NuGetConfigPath; label = 'NuGet configuration' },
    @{ path = $ParserTestPath; label = 'Parser lock validation test' }
)) {
    Assert-RequiredFile -Path $required.path -Label $required.label
}

if ($null -eq $dotnetCommand) {
    throw "Generating packages.lock.json requires .NET SDK $RequiredSdkVersion. End users never need this SDK."
}
$dotnetPath = $dotnetCommand.Source
$actualSdkVersion = (& $dotnetPath --version).Trim()
if ($LASTEXITCODE -ne 0 -or $actualSdkVersion -ne $RequiredSdkVersion) {
    throw "Generating packages.lock.json requires .NET SDK $RequiredSdkVersion, actual: $actualSdkVersion"
}

if ((Test-Path -LiteralPath $LockPath -PathType Leaf) -and -not $Force) {
    throw "A dependency lock already exists: $LockPath. Review it first; use -Force only for an intentional dependency update."
}

# 缓存隔离到 GUID 临时目录，避免用户级 NuGet 源、缓存或凭据影响本次依赖解析。
$env:DOTNET_CLI_HOME = Join-Path $temporaryRoot 'dotnet-home'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
$env:NUGET_PACKAGES = Join-Path $temporaryRoot 'nuget-packages'
$env:NUGET_HTTP_CACHE_PATH = Join-Path $temporaryRoot 'nuget-http-cache'
$env:NUGET_PLUGINS_CACHE_PATH = Join-Path $temporaryRoot 'nuget-plugins-cache'
$intermediateDirectory = (Join-Path $temporaryRoot 'obj') + [IO.Path]::DirectorySeparatorChar
$binaryDirectory = (Join-Path $temporaryRoot 'bin') + [IO.Path]::DirectorySeparatorChar

try {
    New-Item -ItemType Directory -Force -Path $env:DOTNET_CLI_HOME, $env:NUGET_PACKAGES, $env:NUGET_HTTP_CACHE_PATH, $env:NUGET_PLUGINS_CACHE_PATH | Out-Null

    # 生成阶段是唯一允许 --force-evaluate 的地方；显式关闭 locked mode 以重建真实锁文件。
    # 用参数数组代替反引号续行，避免 Windows PowerShell 将 --configfile 误解析为一元运算符。
    $restoreArguments = @(
        'restore',
        $ProjectPath,
        '--configfile',
        $NuGetConfigPath,
        '--runtime',
        'win-x64',
        '--use-lock-file',
        '--force-evaluate',
        '--disable-parallel',
        # dotnet restore 的 CLI 参数为 --no-cache；旧的 NuGet 参数写法会被 MSBuild
        # 误认为未知开关而导致 CI 失败。
        '--no-cache',
        '-p:RestoreLockedMode=false',
        "-p:BaseIntermediateOutputPath=$intermediateDirectory",
        "-p:BaseOutputPath=$binaryDirectory"
    )
    & $dotnetPath @restoreArguments
    if ($LASTEXITCODE -ne 0) {
        throw 'NuGet restore failed while generating packages.lock.json.'
    }

    Assert-RequiredFile -Path $LockPath -Label 'Generated NuGet dependency lock'
    # 复用发行前测试，确保锁包含 PdfPig 精确版本和 NuGet SHA-512 contentHash。
    & $ParserTestPath -RequireLock
    if ($LASTEXITCODE -ne 0) {
        throw 'Generated packages.lock.json did not pass parser dependency validation.'
    }

    [pscustomobject]@{
        pass = $true
        lock_path = $LockPath
        sha256 = (Get-FileHash -LiteralPath $LockPath -Algorithm SHA256).Hash.ToUpperInvariant()
        dotnet_sdk = $actualSdkVersion
        next_step = 'Review packages.lock.json, then commit it before running the release build.'
    } | ConvertTo-Json -Depth 4
} finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        # 只删除本脚本创建的 GUID 临时缓存，不删除项目或锁文件。
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}
