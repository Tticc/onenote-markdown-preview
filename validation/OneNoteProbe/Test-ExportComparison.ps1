$ErrorActionPreference = 'Stop'
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Missing build output' }
$interop = "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll"
$exe = Join-Path $output 'ExportComparison.exe'
$sources = @('ExportComparison.cs','OutsideDiff.cs','SelectionScreen.cs','FixtureText.cs') | ForEach-Object { Join-Path $PSScriptRoot $_ }
$sources += Join-Path $PSScriptRoot '..\..\src\Core\ParagraphPolicy.cs'
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /platform:x64 "/out:$exe" "/reference:$interop" $sources
if ($LASTEXITCODE -ne 0) { throw 'Comparison build failed' }
& $exe
if ($LASTEXITCODE -eq 1) { throw 'Comparison failed' }
