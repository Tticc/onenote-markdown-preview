param([switch]$Remove,[switch]$CheckOnly,[switch]$LifecycleOnly,[switch]$StackDiagnostic,[switch]$FixTrial,[switch]$Acceptance)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$id='{E036DE73-418B-45F7-973F-A4B09D58E515}'
$prog='OneNoteMarkdown.Preview'
$class="HKCU:\Software\Classes\CLSID\$id"
$app="HKCU:\Software\Classes\AppID\$id"
$alias="HKCU:\Software\Classes\$prog"
$addin="HKCU:\Software\Microsoft\Office\OneNote\Addins\$prog"
$keys=@($class,$app,$alias,$addin)
if($Remove){
    foreach($key in $keys){if(Test-Path -LiteralPath $key){Remove-Item -LiteralPath $key -Recurse}}
    'Preview registration removed. Disconnect/restart OneNote normally to unload; no processes or notes changed.'
    return
}
if(-not[Environment]::Is64BitProcess){throw 'Use 64-bit Windows PowerShell'}
$output=Join-Path $root 'artifacts\preview'
$dll=Join-Path $output 'OneNoteMarkdown.Addin.dll'
$manifest=Join-Path $output 'SHA256SUMS.txt'
if(-not(Test-Path -LiteralPath $manifest)){throw 'Run build.ps1 -Test first'}
foreach($line in [IO.File]::ReadAllLines($manifest)){
    if($line -notmatch '^([0-9A-F]{64})  ([A-Za-z0-9_.-]+)$'){throw 'Invalid integrity manifest'}
    $expected=$Matches[1];$name=$Matches[2]
    $path=Join-Path $output $name
    if(-not(Test-Path -LiteralPath $path)){throw "Missing preview file: $name"}
    if((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $expected){throw "Integrity mismatch: $name"}
}
foreach($key in $keys){if(Test-Path -LiteralPath $key){throw "Registration exists; inspect or remove first: $key"}}
if($CheckOnly){'PASS preview integrity and registration preflight; no changes';return}
if(-not $FixTrial -and -not $Acceptance){throw 'Loading suspended. -Remove and -CheckOnly remain available; a controlled host run requires -FixTrial (lifecycle-only) or -Acceptance (full functional) per the reviewed plan in openspec/changes/onenote-safe-markdown-addin/lifecycle-investigation.md.'}
if($FixTrial -and -not(Test-Path -LiteralPath 'C:\Users\18883\AppData\Local\Temp\opencode\debuggers-extracted\Windows Kits\10\Debuggers\x64\cdb.exe')){throw 'Verified stack debugger must be prepared first'}
if(@(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue).Count -ne 0){throw 'Exit OneNote normally first. No process will be terminated.'}
$assembly=[Reflection.AssemblyName]::GetAssemblyName($dll)
$config=Join-Path $env:LOCALAPPDATA 'OneNoteMarkdownPreview'
if(-not(Test-Path -LiteralPath $config)){throw 'Expected preview configuration directory missing; inspect before registering'}
$marker=Join-Path $config 'lifecycle-only'
if($FixTrial){[IO.File]::WriteAllText($marker,'1')}
elseif(Test-Path -LiteralPath $marker){Remove-Item -LiteralPath $marker}
try{
    New-Item -Path "$class\InprocServer32" -Force|Out-Null
    New-ItemProperty -LiteralPath $class -Name AppID -Value $id -PropertyType String|Out-Null
    Set-Item -LiteralPath "$class\InprocServer32" -Value 'mscoree.dll'
    $values=@{ThreadingModel='Both';Class='OneNoteMarkdown.Addin.Addin';Assembly=$assembly.FullName;RuntimeVersion='v4.0.30319';CodeBase=([Uri]$dll).AbsoluteUri}
    foreach($name in $values.Keys){New-ItemProperty -LiteralPath "$class\InprocServer32" -Name $name -Value $values[$name] -PropertyType String|Out-Null}
    New-Item -Path $app -Force|Out-Null
    New-ItemProperty -LiteralPath $app -Name DllSurrogate -Value '' -PropertyType String|Out-Null
    New-Item -Path "$class\ProgId" -Force|Out-Null;Set-Item -LiteralPath "$class\ProgId" -Value $prog
    New-Item -Path "$alias\CLSID" -Force|Out-Null;Set-Item -LiteralPath "$alias\CLSID" -Value $id
    New-Item -Path $addin -Force|Out-Null
    New-ItemProperty -LiteralPath $addin -Name FriendlyName -Value 'OneNote Markdown Preview (manual load)' -PropertyType String|Out-Null
    New-ItemProperty -LiteralPath $addin -Name Description -Value 'Controlled preview; unsupported ranges/formats are refused.' -PropertyType String|Out-Null
    New-ItemProperty -LiteralPath $addin -Name LoadBehavior -Value 0 -PropertyType DWord|Out-Null
    if($FixTrial){'Registered lifecycle-only fix-verification trial (LoadBehavior=0); conversion/hook locked; add-in exits its surrogate on disconnect. No OneNote launched.'}
    else{'Registered full functional acceptance preview (LoadBehavior=0); conversion and Ctrl+Shift+D enabled on load; add-in exits its surrogate on disconnect. No OneNote launched.'}
}catch{
    foreach($key in $keys){if(Test-Path -LiteralPath $key){Remove-Item -LiteralPath $key -Recurse}}
    throw
}
