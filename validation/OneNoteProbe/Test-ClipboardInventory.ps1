$ErrorActionPreference = 'Stop'
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Run offline build first' }
$exe = Join-Path $output 'ClipboardInventory.exe'
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /platform:x64 "/out:$exe" (Join-Path $PSScriptRoot 'ClipboardInventory.cs')
if ($LASTEXITCODE -ne 0) { throw 'Clipboard inventory compile failed' }
& $exe
if ($LASTEXITCODE -eq 1) { throw 'Clipboard inventory failed' }
