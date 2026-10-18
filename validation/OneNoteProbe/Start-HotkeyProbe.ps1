$ErrorActionPreference = 'Stop'
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Missing build output' }
if (@(Get-Process -Name HotkeyProbe -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Existing observational probe running' }
$exe = Join-Path $output 'HotkeyProbe.exe'
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /platform:x64 "/out:$exe" /reference:System.Windows.Forms.dll (Join-Path $PSScriptRoot 'HotkeyProbe.cs')
if ($LASTEXITCODE -ne 0) { throw 'Keyboard observation build failed' }
$log=Join-Path $output 'hotkey-probe.log'
$child=Start-Process -FilePath $exe -ArgumentList ('"{0}"' -f $log) -WindowStyle Hidden -PassThru
"Observational probe started PID=$($child.Id); never intercepts; expires in 90 seconds"
