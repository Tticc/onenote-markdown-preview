$ErrorActionPreference='Stop'
# Read-only recovery baseline. Never activates COM/OneNote, registers or kills.
function IsDesktopHost($process){try{return [bool]($process.Path -and $process.Path -notlike '*\WindowsApps\*')}catch{return $false}}
$all=@(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue)
$desktop=@($all|Where-Object{IsDesktopHost $_})
$store=@($all|Where-Object{-not (IsDesktopHost $_)})
$preview=@(Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" | Where-Object{$_.CommandLine -and $_.CommandLine -like '*E036DE73-418B-45F7-973F-A4B09D58E515*'})
[pscustomobject]@{
    DesktopOneNoteProcesses=$desktop.Count
    StoreOneNoteProcesses=$store.Count
    VisibleDesktopWindows=@($desktop|Where-Object{$_.MainWindowHandle -ne 0}).Count
    PreviewSurrogates=$preview.Count
    PreviewRegistered=Test-Path -LiteralPath 'HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteMarkdown.Preview'
    CleanForFutureControlledLoad=($desktop.Count -eq 0 -and $store.Count -eq 0 -and $preview.Count -eq 0 -and -not(Test-Path -LiteralPath 'HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteMarkdown.Preview'))
}|Format-List
'Store OneNote is a UWP app and never loads this COM add-in; desktop OneNote is the target. No startup attempt or process termination was performed. A clean result does not authorize installation.'
