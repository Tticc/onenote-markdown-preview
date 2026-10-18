param([switch]$Roundtrip)
$ErrorActionPreference = 'Stop'
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Missing build output' }
$exe = Join-Path $output 'OleInventory.exe'
$sources=@('OleInventory.cs','ClipboardRoundtrip.cs','RestorePolicy.cs','SnapshotLimits.cs')|ForEach-Object{Join-Path $PSScriptRoot $_}
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /platform:x64 /main:OleInventory /reference:System.Windows.Forms.dll "/out:$exe" $sources
if ($LASTEXITCODE -ne 0) { throw 'OLE inventory build failed' }
if($Roundtrip){$child=Start-Process -FilePath $exe -ArgumentList '--roundtrip' -NoNewWindow -PassThru}else{$child = Start-Process -FilePath $exe -NoNewWindow -PassThru}
$handle = $child.Handle
try {
    if (-not $child.WaitForExit(5000)) { $child.Kill(); throw 'OLE metadata call exceeded 5 seconds; only inventory process stopped' }
    if ($child.ExitCode -ne 0) { throw "OLE inventory failed: $($child.ExitCode)" }
} finally { $child.Dispose() }
