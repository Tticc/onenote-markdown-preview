param(
    [string]$RegistryRoot='HKCU:\Software',
    [switch]$Uninstall,
    [switch]$Verify,
    [switch]$DryRun,
    [string]$FailAfter='',
    [switch]$TestAllowExisting
)
# OneNote Markdown Preview - install / uninstall / verify / rollback.
# Registers the preview add-in per-user (HKCU) with LoadBehavior=0 (manual load).
# Auto-load (LoadBehavior=3) is intentionally NOT implemented: the gate review has
# not authorized launch-time loading. Enable/disable is done from OneNote's COM
# Add-ins dialog once registered.
# -RegistryRoot exists so the rollback/isolation test can run in an HKCU sandbox.
# -FailAfter and -TestAllowExisting are test-only injection helpers.
$ErrorActionPreference='Stop'

$id='{E036DE73-418B-45F7-973F-A4B09D58E515}'
$prog='OneNoteMarkdown.Preview'
$class=Join-Path $RegistryRoot "Classes\CLSID\$id"
$app=Join-Path $RegistryRoot "Classes\AppID\$id"
$alias=Join-Path $RegistryRoot "Classes\$prog"
$addinsRoot=Join-Path $RegistryRoot 'Microsoft\Office\OneNote\Addins'
$addin=Join-Path $addinsRoot $prog
$keys=@($class,$app,$alias,$addin)
$realRoot=($RegistryRoot -eq 'HKCU:\Software')
$root=Split-Path -Parent $PSScriptRoot

function ToRegPath([string]$psPath){ return $psPath.Replace('HKCU:\','HKCU\') }

function Get-TreeHash([string]$psPath){
    if(-not (Test-Path -LiteralPath $psPath)){ return 'MISSING' }
    $tmp=[IO.Path]::GetTempFileName()
    try{
        & reg.exe export (ToRegPath $psPath) $tmp /y | Out-Null
        if($LASTEXITCODE -ne 0){ throw "reg export failed: $psPath" }
        return (Get-FileHash -LiteralPath $tmp -Algorithm SHA256).Hash
    } finally { if(Test-Path -LiteralPath $tmp){ Remove-Item -LiteralPath $tmp -Force } }
}

function Get-SiblingNames(){
    if(-not (Test-Path -LiteralPath $addinsRoot)){ return @() }
    $names=@(Get-ChildItem -LiteralPath $addinsRoot -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName)
    return @($names | Where-Object { $_ -and $_ -ne $prog })
}

function Get-SiblingHashes(){
    $map=@{}
    foreach($name in Get-SiblingNames){ $map[$name]=Get-TreeHash (Join-Path $addinsRoot $name) }
    return $map
}

function Test-SameMap($a,$b){
    if($a.Count -ne $b.Count){ return $false }
    foreach($k in $a.Keys){ if(-not $b.ContainsKey($k)){ return $false }; if($a[$k] -ne $b[$k]){ return $false } }
    return $true
}

function Maybe-Fail([string]$stage){ if($FailAfter -eq $stage){ throw "Injected failure after stage '$stage' (test only)" } }

$registered=$false
foreach($key in $keys){ if(Test-Path -LiteralPath $key){ $registered=$true } }

if($Uninstall){
    foreach($key in $keys){ if(Test-Path -LiteralPath $key){ Remove-Item -LiteralPath $key -Recurse } }
    "Uninstalled preview registration from $RegistryRoot. Disconnect or restart OneNote normally to unload; no processes or notes changed."
    exit 0
}

if($Verify){
    $state=if($registered){ 'STATE=REGISTERED' } else { 'STATE=NOT_REGISTERED' }
    [Console]::Out.WriteLine($state)
    $behavior='(none)'
    if(Test-Path -LiteralPath $addin){
        $value=Get-ItemProperty -LiteralPath $addin -Name LoadBehavior -ErrorAction SilentlyContinue
        if($null -ne $value){ $behavior=[string]$value.LoadBehavior }
    }
    [Console]::Out.WriteLine("LoadBehavior=$behavior")
    [Console]::Out.WriteLine("Siblings=$((Get-SiblingNames) -join ',')")
    exit 0
}

$output=Join-Path $root 'artifacts\preview'
$dll=Join-Path $output 'OneNoteMarkdown.Addin.dll'
$manifest=Join-Path $output 'SHA256SUMS.txt'
if(-not (Test-Path -LiteralPath $manifest)){ throw 'Run build.ps1 -Test first' }
foreach($line in [IO.File]::ReadAllLines($manifest)){
    if($line -notmatch '^([0-9A-F]{64})  ([A-Za-z0-9_.-]+)$'){ throw 'Invalid integrity manifest' }
    $expected=$Matches[1];$name=$Matches[2]
    $path=Join-Path $output $name
    if(-not (Test-Path -LiteralPath $path)){ throw "Missing preview file: $name" }
    if((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $expected){ throw "Integrity mismatch: $name" }
}
$assembly=[Reflection.AssemblyName]::GetAssemblyName($dll)

if($registered -and -not $TestAllowExisting){ throw 'Registration exists; run -Verify, -Uninstall, or inspect before installing.' }
if($realRoot -and @(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue).Count -ne 0){ throw 'Exit OneNote normally first. No process will be terminated.' }

if($DryRun){
    foreach($key in $keys){ [Console]::Out.WriteLine("would create: $key") }
    [Console]::Out.WriteLine('would set LoadBehavior=0 (manual load); auto-load not implemented')
    exit 0
}

$siblingsBefore=Get-SiblingHashes
$backupFiles=@{}
foreach($key in $keys){
    if(Test-Path -LiteralPath $key){
        $file=[IO.Path]::GetTempFileName()
        & reg.exe export (ToRegPath $key) $file /y | Out-Null
        if($LASTEXITCODE -ne 0){ throw "backup export failed: $key" }
        $backupFiles[$key]=$file
    }
}

function Undo-Install(){
    foreach($key in $keys){ if(Test-Path -LiteralPath $key){ Remove-Item -LiteralPath $key -Recurse } }
    foreach($key in $backupFiles.Keys){
        & reg.exe import $backupFiles[$key] | Out-Null
        if($LASTEXITCODE -ne 0){ throw "rollback import failed: $key" }
    }
}

try{
    New-Item -Path "$class\InprocServer32" -Force | Out-Null
    New-ItemProperty -LiteralPath $class -Name AppID -Value $id -PropertyType String | Out-Null
    Set-Item -LiteralPath "$class\InprocServer32" -Value 'mscoree.dll'
    $values=@{ThreadingModel='Both';Class='OneNoteMarkdown.Addin.Addin';Assembly=$assembly.FullName;RuntimeVersion='v4.0.30319';CodeBase=([Uri]$dll).AbsoluteUri}
    foreach($name in $values.Keys){ New-ItemProperty -LiteralPath "$class\InprocServer32" -Name $name -Value $values[$name] -PropertyType String | Out-Null }
    New-Item -Path "$class\ProgId" -Force | Out-Null; Set-Item -LiteralPath "$class\ProgId" -Value $prog
    Maybe-Fail 'class'

    New-Item -Path $app -Force | Out-Null
    New-ItemProperty -LiteralPath $app -Name DllSurrogate -Value '' -PropertyType String | Out-Null
    Maybe-Fail 'app'

    New-Item -Path "$alias\CLSID" -Force | Out-Null; Set-Item -LiteralPath "$alias\CLSID" -Value $id
    Maybe-Fail 'alias'

    New-Item -Path $addin -Force | Out-Null
    New-ItemProperty -LiteralPath $addin -Name FriendlyName -Value 'OneNote Markdown Preview (manual load)' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath $addin -Name Description -Value 'Controlled preview; unsupported ranges/formats are refused.' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath $addin -Name LoadBehavior -Value 0 -PropertyType DWord | Out-Null
    Maybe-Fail 'addin'

    foreach($key in $keys){ if(-not (Test-Path -LiteralPath $key)){ throw "verify failed, missing: $key" } }
    $behavior=(Get-ItemProperty -LiteralPath $addin -Name LoadBehavior).LoadBehavior
    if([int]$behavior -ne 0){ throw "verify failed, LoadBehavior=$behavior" }

    $siblingsAfter=Get-SiblingHashes
    if(-not (Test-SameMap $siblingsBefore $siblingsAfter)){ throw 'other add-ins changed during install; rolling back' }
} catch {
    Undo-Install
    foreach($file in $backupFiles.Values){ if(Test-Path -LiteralPath $file){ Remove-Item -LiteralPath $file -Force } }
    throw
}
foreach($file in $backupFiles.Values){ if(Test-Path -LiteralPath $file){ Remove-Item -LiteralPath $file -Force } }
"Installed preview registration into $RegistryRoot (LoadBehavior=0, manual load). No OneNote launched. Enable it from OneNote COM Add-ins."
