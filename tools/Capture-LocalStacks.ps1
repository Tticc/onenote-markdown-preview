param(
    [Parameter(Mandatory=$true)][int]$TargetProcessId,
    [string]$DebuggerPath='C:\Users\18883\AppData\Local\Temp\opencode\debuggers-extracted\Windows Kits\10\Debuggers\x64\cdb.exe',
    [switch]$SelfTest
)
$ErrorActionPreference='Stop'
if(-not(Test-Path -LiteralPath $DebuggerPath)){throw 'Debugger not prepared'}
$signature=Get-AuthenticodeSignature -LiteralPath $DebuggerPath
if($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notlike '*Microsoft Corporation*'){throw 'Debugger signature check failed'}
$target=Get-Process -Id $TargetProcessId -ErrorAction Stop
if(-not $SelfTest -and $target.ProcessName -ne 'ONENOTE'){throw 'This script captures only OneNote; self-test uses a separate explicitly supplied fixture'}
$root=Split-Path -Parent $PSScriptRoot
$output=Join-Path $root 'artifacts\diagnostics'
if(-not(Test-Path -LiteralPath $output)){New-Item -ItemType Directory -Path $output -Force|Out-Null}
$stamp=Get-Date -Format 'yyyyMMdd-HHmmss'
$log=Join-Path $output "stacks-$TargetProcessId-$stamp.txt"
# Non-invasive, no target suspension (-pv), local-only symbols. k prints frames,
# not arguments/locals; no .dump/db/du/heap commands or network symbol server.
$arguments=@('-pv','-p',"$TargetProcessId",'-y',('"{0}"' -f $output),'-logo',('"{0}"' -f $log),'-c','"~*k;qd"')
$debugger=Start-Process -FilePath $DebuggerPath -ArgumentList $arguments -PassThru -WindowStyle Hidden
$handle=$debugger.Handle
try{
    if(-not $debugger.WaitForExit(30000)){
        $debugger.Kill()
        throw 'Stack capture timeout; stopped only the non-invasive debugger, not target'
    }
    if($debugger.ExitCode -ne 0){throw "Stack debugger exit=$($debugger.ExitCode); inspect local log"}
    "Local stack capture finished: $log"
    'No dump created, no target killed, no note/clipboard operations. Non-suspended stacks may be incomplete; exit code alone is not diagnostic success.'
}finally{$debugger.Dispose()}
