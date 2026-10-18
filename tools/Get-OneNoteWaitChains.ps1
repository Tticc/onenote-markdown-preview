$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$output=Join-Path $root 'artifacts\diagnostics'
if(-not(Test-Path -LiteralPath $root)){throw 'Workspace missing'}
if(-not(Test-Path -LiteralPath $output)){New-Item -ItemType Directory -Path $output -Force|Out-Null}
$exe=Join-Path $output 'WaitChainProbe.exe'
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /platform:x64 "/out:$exe" (Join-Path $PSScriptRoot 'WaitChainProbe.cs')
if($LASTEXITCODE -ne 0){throw 'WCT probe compile failed'}
foreach($hostProcess in @(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue)){
    $child=Start-Process -FilePath $exe -ArgumentList $hostProcess.Id -NoNewWindow -PassThru
    $handle=$child.Handle
    try{
        if(-not $child.WaitForExit(15000)){$child.Kill();throw 'WCT probe timeout; only diagnostic process stopped'}
        if($child.ExitCode -ne 0){throw 'WCT diagnostic failed; host untouched'}
    }finally{$child.Dispose()}
}
