# Only for the corrected, independently tested binary. No startup auto-load.
$ErrorActionPreference = 'Stop'
if (@(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Exit OneNote normally first.' }
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell.' }
$id = '{AE4143BE-ABDB-4B51-989B-D894A7BD63D5}'
$classKey = "HKCU:\Software\Classes\CLSID\$id"
$appKey = "HKCU:\Software\Classes\AppID\$id"
$progKey = 'HKCU:\Software\Classes\OneNoteMarkdown.ValidationProbe'
$addinKey = 'HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteMarkdown.ValidationProbe'
$keys = @($classKey, $appKey, $progKey, $addinKey)
foreach ($key in $keys) { if (Test-Path -LiteralPath $key) { throw "Existing registration: $key" } }
$dll = Join-Path $PSScriptRoot 'offline\ProbeAddin.dll'
$ext = [Reflection.Assembly]::LoadFrom("$env:WINDIR\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll")
$assembly = [Reflection.Assembly]::LoadFrom($dll)
if (-not $ext.GetType('Extensibility.IDTExtensibility2').IsAssignableFrom($assembly.GetType('ProbeAddin'))) { throw 'Official contract check failed' }
try {
    New-Item -Path "$classKey\InprocServer32" -Force | Out-Null
    New-ItemProperty -LiteralPath $classKey -Name AppID -Value $id -PropertyType String | Out-Null
    Set-Item -LiteralPath "$classKey\InprocServer32" -Value 'mscoree.dll'
    $values = @{ThreadingModel='Both';Class='ProbeAddin';Assembly=$assembly.FullName;RuntimeVersion='v4.0.30319';CodeBase=([Uri]$dll).AbsoluteUri}
    foreach ($name in $values.Keys) { New-ItemProperty -LiteralPath "$classKey\InprocServer32" -Name $name -Value $values[$name] -PropertyType String | Out-Null }
    New-Item -Path $appKey -Force | Out-Null
    New-ItemProperty -LiteralPath $appKey -Name DllSurrogate -Value '' -PropertyType String | Out-Null
    New-Item -Path "$classKey\ProgId" -Force | Out-Null
    Set-Item -LiteralPath "$classKey\ProgId" -Value 'OneNoteMarkdown.ValidationProbe'
    New-Item -Path "$progKey\CLSID" -Force | Out-Null
    Set-Item -LiteralPath "$progKey\CLSID" -Value $id
    New-Item -Path $addinKey -Force | Out-Null
    New-ItemProperty -LiteralPath $addinKey -Name FriendlyName -Value 'Markdown validation probe (manual only)' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath $addinKey -Name Description -Value 'Corrected loading probe; no note or clipboard edits.' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath $addinKey -Name LoadBehavior -Value 0 -PropertyType DWord | Out-Null
    'Registered corrected probe with LoadBehavior=0. Manual connection only.'
} catch {
    foreach ($key in $keys) { if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse } }
    throw
}
