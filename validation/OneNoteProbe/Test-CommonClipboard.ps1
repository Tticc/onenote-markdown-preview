$ErrorActionPreference = 'Stop'
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Missing build output' }
$exe = Join-Path $output 'CommonClipboardTests.exe'
$sources = @('CommonClipboardTests.cs','ClipboardRoundtrip.cs','RestorePolicy.cs','HtmlPayloadTests.cs') | ForEach-Object { Join-Path $PSScriptRoot $_ }
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /platform:x64 /main:CommonClipboardTests "/out:$exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll $sources
if ($LASTEXITCODE -ne 0) { throw 'Common clipboard test build failed' }
& $exe
if ($LASTEXITCODE -ne 0) { throw 'Common clipboard tests failed' }
