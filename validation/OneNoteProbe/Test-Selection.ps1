$ErrorActionPreference = 'Stop'
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$interop = "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll"
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Run offline build first' }
$exe = Join-Path $output 'SelectionProbe.exe'
& $compiler /nologo /platform:x64 "/out:$exe" "/reference:$interop" (Join-Path $PSScriptRoot 'SelectionProbe.cs') (Join-Path $PSScriptRoot 'SelectionScreen.cs') (Join-Path $PSScriptRoot 'SelectionEvidence.cs') (Join-Path $PSScriptRoot 'FixtureText.cs')
if ($LASTEXITCODE -ne 0) { throw 'Selection probe compile failed' }
& $exe
if ($LASTEXITCODE -eq 1) { throw 'Selection probe failed' }
