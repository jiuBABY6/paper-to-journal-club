<#
  发布阶段执行一次：生成 assets/paper-parser.exe。
  该 exe 为 win-x64 self-contained single file，最终用户无需安装 Node 或 .NET。

  编译缓存统一写入系统临时目录，避免将 SDK、NuGet 缓存和发布中间文件提交到 GitHub 仓库。
#>
$ErrorActionPreference = 'Stop'
$pluginRoot = Split-Path -Parent $PSScriptRoot
$project = Join-Path $pluginRoot 'parser\PaperParser.csproj'
$nugetConfig = Join-Path $pluginRoot 'parser\NuGet.Config'
$output = Join-Path $pluginRoot 'assets\paper-parser.exe'
$dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
$buildRoot = Join-Path ([IO.Path]::GetTempPath()) "paper-to-journal-club-parser-build-$([Guid]::NewGuid().ToString('N'))"
$publishDirectory = Join-Path $buildRoot 'publish'

if ($dotnetCommand) {
    $dotnetPath = $dotnetCommand.Source
} else {
    throw 'dotnet SDK is required only for release builders. End users receive the compiled paper-parser.exe.'
}

$env:DOTNET_CLI_HOME = Join-Path $buildRoot 'dotnet-home'
$env:NUGET_PACKAGES = Join-Path $buildRoot 'nuget-packages'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'

try {
    New-Item -ItemType Directory -Force -Path $env:DOTNET_CLI_HOME, $env:NUGET_PACKAGES, $publishDirectory | Out-Null
    & $dotnetPath publish $project -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true --configfile $nugetConfig -o $publishDirectory
    if ($LASTEXITCODE -ne 0) { throw 'paper-parser build failed.' }

    $builtExecutable = Join-Path $publishDirectory 'PaperParser.exe'
    if (-not (Test-Path -LiteralPath $builtExecutable -PathType Leaf)) {
        throw "paper-parser build completed without the expected executable: $builtExecutable"
    }
    Copy-Item -LiteralPath $builtExecutable -Destination $output -Force
    Write-Output "Built $output"
}
finally {
    # 仅删除本脚本创建的带 GUID 临时目录，不影响用户源码和已发布的解析器。
    if (Test-Path -LiteralPath $buildRoot) {
        Remove-Item -LiteralPath $buildRoot -Recurse -Force
    }
}
