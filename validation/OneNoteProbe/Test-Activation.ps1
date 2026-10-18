$ErrorActionPreference = 'Stop'
$id = '{AE4143BE-ABDB-4B51-989B-D894A7BD63D5}'
$classKey = "HKCU:\Software\Classes\CLSID\$id"
$appKey = "HKCU:\Software\Classes\AppID\$id"
$addinKey = 'HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteMarkdown.ValidationProbe'
foreach ($key in @($classKey, $appKey, $addinKey)) {
    if (Test-Path -LiteralPath $key) { throw "Existing registration; refusing: $key" }
}
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$office = "$env:WINDIR\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll"
$extensibility = "$env:WINDIR\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll"
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Run Test-Offline.ps1 first' }
$dll = Join-Path $output 'ProbeAddin.dll'
$exe = Join-Path $output 'SurrogateProbe.exe'
& $compiler /nologo /platform:x64 "/out:$exe" "/reference:$office" "/reference:$extensibility" (Join-Path $PSScriptRoot 'SurrogateProbe.cs')
if ($LASTEXITCODE -ne 0) { throw 'Activation client compilation failed' }
$child = $null
try {
    New-Item -Path "$classKey\InprocServer32" -Force | Out-Null
    New-ItemProperty -LiteralPath $classKey -Name AppID -Value $id -PropertyType String | Out-Null
    Set-Item -LiteralPath "$classKey\InprocServer32" -Value 'mscoree.dll'
    $values = @{ThreadingModel='Both';Class='ProbeAddin';Assembly=[Reflection.AssemblyName]::GetAssemblyName($dll).FullName;RuntimeVersion='v4.0.30319';CodeBase=([Uri]$dll).AbsoluteUri}
    foreach ($name in $values.Keys) { New-ItemProperty -LiteralPath "$classKey\InprocServer32" -Name $name -Value $values[$name] -PropertyType String | Out-Null }
    New-Item -Path $appKey -Force | Out-Null
    New-ItemProperty -LiteralPath $appKey -Name DllSurrogate -Value '' -PropertyType String | Out-Null
    $child = Start-Process -FilePath $exe -PassThru -NoNewWindow
    $processHandle = $child.Handle
    if (-not $child.WaitForExit(15000)) { $child.Kill(); throw 'Activation client timeout' }
    if ($child.ExitCode -ne 0) { throw "Activation client failed: $($child.ExitCode)" }
    $deadline = (Get-Date).AddSeconds(15)
    do {
        $hosts = @(Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($id, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
        if ($hosts.Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    if ($hosts.Count -ne 0) { throw 'Activation passed but surrogate did not exit within 15 seconds; no unrelated process will be terminated' }
    'PASS activation client and matching surrogate exited; no OneNote Addins entry was created'
} finally {
    if ($null -ne $child) { if (-not $child.HasExited) { $child.Kill() }; $child.Dispose() }
    foreach ($key in @($classKey, $appKey)) { if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse } }
    'Removed activation-only registry keys'
}
