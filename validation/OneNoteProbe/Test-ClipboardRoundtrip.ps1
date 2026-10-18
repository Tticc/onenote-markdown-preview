param([int]$DelaySeconds = 0, [switch]$SyntheticDisposable, [switch]$SimulateNewerCopy, [switch]$RestoreAfter500Ms)
$ErrorActionPreference = 'Stop'
if ($DelaySeconds -lt 0 -or $DelaySeconds -gt 60) { throw 'Delay out of range' }
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Run offline build first' }
$exe = Join-Path $output 'ClipboardRoundtrip.exe'
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /platform:x64 "/out:$exe" /reference:System.Windows.Forms.dll (Join-Path $PSScriptRoot 'ClipboardRoundtrip.cs')
if ($LASTEXITCODE -ne 0) { throw 'Clipboard roundtrip compile failed' }
if ($DelaySeconds -gt 0) { Start-Sleep -Seconds $DelaySeconds }
if ($SimulateNewerCopy -and -not $SyntheticDisposable) { throw 'Newer-copy experiment requires synthetic disposable authorization' }
if($RestoreAfter500Ms -and ($SyntheticDisposable -or $SimulateNewerCopy)){throw '500ms sample requires existing plain disposable clipboard; no reset'}
if ($RestoreAfter500Ms) { & $exe --delay-500 } elseif ($SimulateNewerCopy) { & $exe --synthetic-newer-copy } elseif ($SyntheticDisposable) { & $exe --synthetic-disposable } else { & $exe }
if ($LASTEXITCODE -eq 1) { throw 'Disposable clipboard test failed; inspect output before further use' }
