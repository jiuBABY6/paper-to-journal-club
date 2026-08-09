<#
  发布安装器可调用的前置条件检查。
  普通用户不需要 Node.js、npm、Python 或 .NET；开发者构建 parser.exe 才需要 .NET SDK。
#>
$ErrorActionPreference = 'Stop'
$pluginRoot = Split-Path -Parent $PSScriptRoot
$parser = Join-Path $pluginRoot 'assets\paper-parser.exe'
$powerPointInstalled = Test-Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID'
$wordInstalled = Test-Path 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID'

[pscustomobject]@{
    supported = $powerPointInstalled
    operating_system = 'Windows'
    target_application = 'Microsoft PowerPoint'
    powerpoint_com_registered = $powerPointInstalled
    bundled_pdf_parser = Test-Path -LiteralPath $parser
    word_pdf_fallback = $wordInstalled
    node_required = $false
    python_required = $false
    dotnet_required_for_end_user = $false
    message = if ($powerPointInstalled) { 'Ready for PowerPoint generation.' } else { 'Microsoft PowerPoint desktop is required.' }
} | ConvertTo-Json
