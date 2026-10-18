param()
# Offline isolation test for Install-Addin.ps1.
# Runs the real installer against an HKCU sandbox that also contains foreign
# add-in and CLSID keys, then drives install / rollback / uninstall and asserts:
#   - only the preview's four keys are ever created or removed
#   - foreign keys and values are byte-for-byte unchanged
#   - a failed install rolls back completely (including restoring pre-existing keys)
# No real registration, host launch, or clipboard use. Sandbox is removed at the end.
$ErrorActionPreference='Stop'

$here=Split-Path -Parent $PSCommandPath
$install=Join-Path $here 'Install-Addin.ps1'
$sandbox='HKCU:\Software\OneNoteMarkdownInstallTest'
$id='{E036DE73-418B-45F7-973F-A4B09D58E515}'
$prog='OneNoteMarkdown.Preview'

$class=Join-Path $sandbox "Classes\CLSID\$id"
$app=Join-Path $sandbox "Classes\AppID\$id"
$alias=Join-Path $sandbox "Classes\$prog"
$addin=Join-Path $sandbox "Microsoft\Office\OneNote\Addins\$prog"
$ourKeys=@($class,$app,$alias,$addin)

$otherAddins=Join-Path $sandbox 'Microsoft\Office\OneNote\Addins'
$otherA=Join-Path $otherAddins 'Contoso.OtherAddin'
$otherB=Join-Path $otherAddins 'Fabrikam.OtherAddin'
$otherClsid=Join-Path $sandbox 'Classes\CLSID\{11111111-1111-1111-1111-111111111111}'
$foreignKeys=@($otherA,$otherB,$otherClsid)

$script:failures=0
function Assert([string]$name,[bool]$ok){
    if($ok){ "PASS $name" } else { "FAIL $name"; $script:failures++ }
}
function ToRegPath([string]$psPath){ return $psPath.Replace('HKCU:\','HKCU\') }
function TreeHash([string]$psPath){
    if(-not (Test-Path -LiteralPath $psPath)){ return 'MISSING' }
    $tmp=[IO.Path]::GetTempFileName()
    try{
        & reg.exe export (ToRegPath $psPath) $tmp /y | Out-Null
        if($LASTEXITCODE -ne 0){ throw "reg export failed: $psPath" }
        return (Get-FileHash -LiteralPath $tmp -Algorithm SHA256).Hash
    } finally { if(Test-Path -LiteralPath $tmp){ Remove-Item -LiteralPath $tmp -Force } }
}
function ForeignSnapshot(){
    $map=@{}
    foreach($key in $foreignKeys){ $map[$key]=TreeHash $key }
    return $map
}
function SameSnapshot($a,$b){
    if($a.Count -ne $b.Count){ return $false }
    foreach($k in $a.Keys){ if($a[$k] -ne $b[$k]){ return $false } }
    return $true
}
function OurKeysMissing(){ foreach($key in $ourKeys){ if(Test-Path -LiteralPath $key){ return $false } }; return $true }
function OurKeysPresent(){ foreach($key in $ourKeys){ if(-not (Test-Path -LiteralPath $key)){ return $false } }; return $true }

function Invoke-Installer([string]$stage,[bool]$allowExisting){
    $argv=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$install,'-RegistryRoot',$sandbox)
    if($stage){ $argv+='-FailAfter'; $argv+=$stage }
    if($allowExisting){ $argv+='-TestAllowExisting' }
    $old=$ErrorActionPreference
    $ErrorActionPreference='Continue'
    try{
        $text=(& powershell.exe @argv 2>&1 | Out-String)
        $code=$LASTEXITCODE
    } finally { $ErrorActionPreference=$old }
    return @{ Code=$code; Text=$text }
}

function Reset-Sandbox(){
    if(Test-Path -LiteralPath $sandbox){ Remove-Item -LiteralPath $sandbox -Recurse -Force }
    New-Item -Path $otherAddins -Force | Out-Null
    foreach($pair in @(@($otherA,'Contoso Other'),@($otherB,'Fabrikam Other'))){
        New-Item -Path $pair[0] -Force | Out-Null
        New-ItemProperty -LiteralPath $pair[0] -Name FriendlyName -Value $pair[1] -PropertyType String | Out-Null
        New-ItemProperty -LiteralPath $pair[0] -Name LoadBehavior -Value 3 -PropertyType DWord | Out-Null
    }
    New-Item -Path (Join-Path $otherClsid 'InprocServer32') -Force | Out-Null
    Set-Item -LiteralPath (Join-Path $otherClsid 'InprocServer32') -Value 'foreign.dll'
}

try{
    Reset-Sandbox
    $baseline=ForeignSnapshot
    Assert 'sandbox foreign keys seeded' ($baseline.Count -eq 3 -and -not (TreeHash $otherA).Equals('MISSING'))

    $dry=(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $install -RegistryRoot $sandbox -DryRun 2>&1 | Out-String)
    Assert 'dry-run reports planned keys' ($dry -match 'would create')
    Assert 'dry-run made no changes' (OurKeysMissing)
    Assert 'dry-run left foreign keys intact' (SameSnapshot $baseline (ForeignSnapshot))

    $r=Invoke-Installer '' $false
    Assert 'install exits 0' ($r.Code -eq 0)
    Assert 'install created all four keys' (OurKeysPresent)
    Assert 'install set LoadBehavior=0' ([int](Get-ItemProperty -LiteralPath $addin -Name LoadBehavior).LoadBehavior -eq 0)
    Assert 'install left foreign keys intact' (SameSnapshot $baseline (ForeignSnapshot))

    $v=(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $install -RegistryRoot $sandbox -Verify 2>&1 | Out-String)
    Assert 'verify reports REGISTERED' ($v -match 'STATE=REGISTERED' -and $v -match 'LoadBehavior=0')

    $u=(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $install -RegistryRoot $sandbox -Uninstall 2>&1 | Out-String)
    Assert 'uninstall removed all four keys' (OurKeysMissing)
    Assert 'uninstall left foreign keys intact' (SameSnapshot $baseline (ForeignSnapshot))

    $v=(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $install -RegistryRoot $sandbox -Verify 2>&1 | Out-String)
    Assert 'verify reports NOT_REGISTERED' ($v -match 'STATE=NOT_REGISTERED')

    $r=Invoke-Installer 'app' $false
    Assert 'mid-install failure is reported' ($r.Code -ne 0 -and $r.Text -match 'Injected failure')
    Assert 'failed install rolled back all four keys' (OurKeysMissing)
    Assert 'failed install left foreign keys intact' (SameSnapshot $baseline (ForeignSnapshot))

    New-Item -Path $addin -Force | Out-Null
    New-ItemProperty -LiteralPath $addin -Name FriendlyName -Value 'PREEXISTING' -PropertyType String | Out-Null
    New-ItemProperty -LiteralPath $addin -Name LoadBehavior -Value 3 -PropertyType DWord | Out-Null
    $preAddin=TreeHash $addin
    $r=Invoke-Installer 'class' $true
    Assert 'failure with pre-existing key is reported' ($r.Code -ne 0)
    Assert 'rollback restored the pre-existing key' ((TreeHash $addin) -eq $preAddin)
    Assert 'rollback removed the other three keys' ((Test-Path -LiteralPath $class) -eq $false -and (Test-Path -LiteralPath $app) -eq $false -and (Test-Path -LiteralPath $alias) -eq $false)
    Assert 'rollback left foreign keys intact' (SameSnapshot $baseline (ForeignSnapshot))

    # Regression: real environments often have no sibling add-ins and no Addins key at all.
    if(Test-Path -LiteralPath $sandbox){ Remove-Item -LiteralPath $sandbox -Recurse -Force }
    New-Item -ItemType Directory -Path $sandbox -Force | Out-Null
    $r=Invoke-Installer '' $false
    Assert 'install works with no existing Addins key' ($r.Code -eq 0 -and (OurKeysPresent))

    if($script:failures -eq 0){ 'PASS install/uninstall/rollback isolation; only the preview keys were touched' }
    else { "FAIL install rollback isolation with $($script:failures) failure(s)"; exit 1 }
} finally {
    if(Test-Path -LiteralPath $sandbox){ Remove-Item -LiteralPath $sandbox -Recurse -Force }
    'Sandbox removed; no real registration changed.'
}
