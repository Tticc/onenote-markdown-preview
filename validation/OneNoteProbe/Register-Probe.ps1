param([switch]$Remove)
$ErrorActionPreference = 'Stop'
$id = '{AE4143BE-ABDB-4B51-989B-D894A7BD63D5}'
$progId = 'OneNoteMarkdown.ValidationProbe'
$classKey = "HKCU:\Software\Classes\CLSID\$id"
$progKey = "HKCU:\Software\Classes\$progId"
$appKey = "HKCU:\Software\Classes\AppID\$id"
$addinKey = "HKCU:\Software\Microsoft\Office\OneNote\Addins\$progId"
if ($Remove) {
    foreach ($key in @($addinKey, $progKey, $classKey, $appKey)) {
        if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse }
    }
    Write-Output 'Only the validation probe registration was removed. Restart OneNote to unload it.'
    return
}
throw 'Host registration is suspended after the shutdown incident. Use Test-Offline.ps1 only; -Remove remains available.'
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit Windows PowerShell.' }
$dll = Join-Path $PSScriptRoot 'ProbeAddin.dll'
if (-not (Test-Path -LiteralPath $dll)) { throw 'Build ProbeAddin.dll first.' }
foreach ($key in @($classKey, $progKey, $addinKey, $appKey)) {
    if (Test-Path -LiteralPath $key) { throw "Registration already exists; inspect before replacing: $key" }
}
$assembly = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
try {
    New-Item -Path "$classKey\InprocServer32" -Force | Out-Null
    New-ItemProperty -LiteralPath $classKey -Name AppID -Value $id -PropertyType String | Out-Null
    New-Item -Path $appKey -Force | Out-Null
    New-ItemProperty -LiteralPath $appKey -Name DllSurrogate -Value '' -PropertyType String | Out-Null
    Set-Item -LiteralPath "$classKey\InprocServer32" -Value 'mscoree.dll'
    $values = @{ ThreadingModel = 'Both'; Class = 'ProbeAddin'; Assembly = $assembly; RuntimeVersion = 'v4.0.30319'; CodeBase = ([Uri]$dll).AbsoluteUri }
    foreach ($name in $values.Keys) { New-ItemProperty -LiteralPath "$classKey\InprocServer32" -Name $name -Value $values[$name] -PropertyType String | Out-Null }
    New-Item -Path "$classKey\ProgId" -Force | Out-Null
    Set-Item -LiteralPath "$classKey\ProgId" -Value $progId
    New-Item -Path "$progKey\CLSID" -Force | Out-Null
    Set-Item -LiteralPath "$progKey\CLSID" -Value $id
    New-Item -Path $addinKey -Force | Out-Null
    New-ItemProperty -LiteralPath $addinKey -Name FriendlyName -Value 'Markdown validation probe' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath $addinKey -Name Description -Value 'Temporary Ribbon loading probe. Does not edit notes.' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath $addinKey -Name LoadBehavior -Value 3 -PropertyType DWord | Out-Null
    Write-Output 'Registered per-user validation probe. Restart OneNote manually to test loading.'
} catch {
    foreach ($key in @($addinKey, $progKey, $classKey, $appKey)) { if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse } }
    throw
}
