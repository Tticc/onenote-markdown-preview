param([switch]$ReadOnlyHost)
$ErrorActionPreference = 'Stop'
# Default run does not read/write clipboard, register plugins, or access OneNote.
# COM fixture lifecycle logs contain only synthetic callback metadata.
& (Join-Path $PSScriptRoot 'Test-Offline.ps1')
if (-not $?) { throw 'Offline validation failed' }
& (Join-Path $PSScriptRoot 'Start-PasteFixture.ps1') -CheckOnly
if (-not $?) { throw 'Input fixture compilation failed' }
if ($ReadOnlyHost) {
    & (Join-Path $PSScriptRoot 'Test-ExportComparison.ps1')
    if (-not $?) { throw 'Read-only host comparison failed' }
}
$keys = @(
    'HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteMarkdown.ValidationProbe',
    'HKCU:\Software\Classes\CLSID\{AE4143BE-ABDB-4B51-989B-D894A7BD63D5}',
    'HKCU:\Software\Classes\AppID\{AE4143BE-ABDB-4B51-989B-D894A7BD63D5}'
)
foreach ($key in $keys) {
    if (Test-Path -LiteralPath $key) { throw "Unexpected probe registration remains: $key" }
}
$leftovers = @(Get-Process -Name CrossProcessTests,PasteFixture,HotkeyProbe,CommonClipboardTests,OleInventory -ErrorAction SilentlyContinue)
if ($leftovers.Count -ne 0) { throw 'A validation process is still running; inspect before continuing' }
'PASS automated validation and cleanup audit. This does not release the OpenSpec phase-one gate.'
