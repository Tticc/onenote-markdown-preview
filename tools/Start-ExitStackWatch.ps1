$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
if(-not(Test-Path -LiteralPath (Join-Path $root 'artifacts\diagnostics'))){throw 'Prepare debugger self-test first'}
$script=Join-Path $PSScriptRoot 'Watch-ExitStacks.ps1'
$child=Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $script)) -WindowStyle Hidden -PassThru
"Exit-stack observation running PID=$($child.Id) for up to 10 minutes; preview registration is removed automatically at end"
