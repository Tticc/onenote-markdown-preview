param([int]$MaximumSeconds=600)
$ErrorActionPreference='Stop'
if($MaximumSeconds -lt 30 -or $MaximumSeconds -gt 600){throw 'Invalid observation limit'}
$root=Split-Path -Parent $PSScriptRoot
$output=Join-Path $root 'artifacts\diagnostics'
if(-not(Test-Path -LiteralPath $output)){throw 'Prepare diagnostics directory first'}
$journal=Join-Path $output 'exit-stack-watcher.log'
function Record([string]$text){[IO.File]::AppendAllText($journal,([DateTime]::UtcNow.ToString('o')+' '+$text+[Environment]::NewLine))}
function AllHosts(){@(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue)}
function IsDesktopHost($process){try{return [bool]($process.Path -and $process.Path -notlike '*\WindowsApps\*')}catch{return $false}}
function DesktopHosts(){@(AllHosts|Where-Object{IsDesktopHost $_})}
function Surrogates(){@(Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" -ErrorAction SilentlyContinue|Where-Object{$_.CommandLine -and $_.CommandLine -like '*E036DE73-418B-45F7-973F-A4B09D58E515*'}).Count}
$deadline=(Get-Date).AddSeconds($MaximumSeconds)
$seen=$false;$started=$false;$storeNoted=$false;$targetId=0;$hiddenAt=$null
Record 'START; watch only; no process termination or auto-launch'
try{
    while((Get-Date) -lt $deadline){
        $all=AllHosts
        $store=@($all|Where-Object{-not (IsDesktopHost $_)})
        if($store.Count -gt 0 -and -not $storeNoted){$storeNoted=$true;Record ('ignoring '+$store.Count+' store/WindowsApps ONENOTE instance(s); they never load this add-in')}
        $hosts=@($all|Where-Object{IsDesktopHost $_})
        if($hosts.Count -gt 1){Record 'STOP multiple desktop hosts; no capture';break}
        if($hosts.Count -eq 0){
            if($seen){Record ('PASS host process exited naturally; preview dllhost='+(Surrogates)+'; no retained-process stacks available');break}
        }else{
            if(-not $started){$started=$true;Record ('host process started pid='+$hosts[0].Id+' (launch-to-window clock begins)')}
            $current=$hosts[0]
            if($targetId -ne 0 -and $targetId -ne $current.Id){Record 'STOP host PID changed';break}
            if($current.MainWindowHandle -ne 0){
                if(-not $seen){$seen=$true;$targetId=$current.Id;Record "Visible host observed pid=$targetId"}
                $hiddenAt=$null
            }elseif($seen){
                if($null -eq $hiddenAt){$hiddenAt=Get-Date;Record ('Window closed; waiting 10 seconds for normal exit; preview dllhost='+(Surrogates))}
                if(((Get-Date)-$hiddenAt).TotalSeconds -ge 10){
                    Record ("Retained invisible host pid=$targetId; preview dllhost="+(Surrogates)+"; capturing local non-invasive stacks")
                    & (Join-Path $PSScriptRoot 'Capture-LocalStacks.ps1') -TargetProcessId $targetId
                    Record 'First capture finished; waiting three seconds'
                    Start-Sleep -Seconds 3
                    if(Get-Process -Id $targetId -ErrorAction SilentlyContinue){& (Join-Path $PSScriptRoot 'Capture-LocalStacks.ps1') -TargetProcessId $targetId;Record 'Second capture finished'}
                    break
                }
            }
        }
        Start-Sleep -Milliseconds 500
    }
    if((Get-Date) -ge $deadline){Record 'Observation deadline reached; no new experiment started'}
}catch{Record ('ERROR '+$_.Exception.GetType().Name);throw}
finally{
    & (Join-Path $PSScriptRoot 'Register-Preview.ps1') -Remove
    Record 'ROLLBACK preview registration removed; no target killed'
}
