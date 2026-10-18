param([string]$Label='unlabeled',[int]$Samples=1,[int]$TimeoutSeconds=180,[switch]$WaitForExit)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$output=Join-Path $root 'artifacts\diagnostics'
if(-not(Test-Path -LiteralPath $output)){throw 'Prepare diagnostics directory first'}
$journal=Join-Path $output 'startup-measurements.log'
function Record([string]$text){[IO.File]::AppendAllText($journal,([DateTime]::UtcNow.ToString('o')+' ['+$Label+'] '+$text+[Environment]::NewLine))}
function IsDesktopHost($process){try{return [bool]($process.Path -and $process.Path -notlike '*\WindowsApps\*')}catch{return $false}}
function AllHosts(){@(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue)}
function DesktopHosts(){@(AllHosts|Where-Object{IsDesktopHost $_})}
function Describe($process){
    $info=Get-CimInstance Win32_Process -Filter ('ProcessId='+$process.Id) -ErrorAction SilentlyContinue
    $parent=0;if($info){$parent=$info.ParentProcessId}
    $parentName='';if($parent -ne 0){$parentName=(Get-Process -Id $parent -ErrorAction SilentlyContinue).Name}
    return ('pid='+$process.Id+' parent='+$parent+'('+$parentName+') start='+$process.StartTime.ToString('HH:mm:ss.fff'))
}
if((DesktopHosts).Count -ne 0){throw 'Exit desktop OneNote first. No process will be terminated.'}
$storeIgnored=(@(AllHosts|Where-Object{-not (IsDesktopHost $_)}).Count)
if($storeIgnored -gt 0){Record ('ignoring '+$storeIgnored+' store/WindowsApps ONENOTE instance(s); they never load this add-in')}
Write-Output ('['+$Label+'] '+$Samples+' sample(s); launch one desktop OneNote per sample. No process is started by this tool.')
for($sample=1;$sample -le $Samples;$sample++){
    $known=@((DesktopHosts)|ForEach-Object{$_.Id})
    Record ('sample '+$sample+' of '+$Samples+': awaiting launch; no process launched by this tool')
    $deadline=(Get-Date).AddSeconds($TimeoutSeconds)
    $recorded=@{};$target=$null
    while((Get-Date) -lt $deadline){
        $hosts=DesktopHosts
        $fresh=@($hosts|Where-Object{$known -notcontains $_.Id})
        foreach($one in $fresh){if(-not $recorded.ContainsKey($one.Id)){$recorded[$one.Id]=$true;Record ('observed new '+(Describe $one))}}
        $withWindow=@($fresh|Where-Object{$_.MainWindowHandle -ne 0})
        if($withWindow.Count -ge 1){Record ('fresh hosts='+$fresh.Count+' with a window='+$withWindow.Count);$target=$withWindow[0];break}
        Start-Sleep -Milliseconds 100
    }
    if($null -eq $target){Record ('sample '+$sample+' STOP no desktop host with a window before the deadline');throw 'No desktop host with a window before the deadline'}
    $started=$target.StartTime
    $latency=((Get-Date)-$started).TotalMilliseconds
    Record ('sample '+$sample+' target pid='+$target.Id+' window shown after '+[int]$latency+' ms (launch-to-window)')
    $extras=@(DesktopHosts|Where-Object{$_.Id -ne $target.Id})
    Record ('sample '+$sample+' extra desktop hosts besides the target='+$extras.Count)
    foreach($one in $extras){Record ('sample '+$sample+' extra detail '+(Describe $one))}
    Write-Output ('['+$Label+' sample '+$sample+'] launch-to-window '+[int]$latency+' ms (pid='+$target.Id+', extras='+$extras.Count+')')
    if($WaitForExit){
        $exitDeadline=(Get-Date).AddSeconds($TimeoutSeconds)
        while((Get-Process -Id $target.Id -ErrorAction SilentlyContinue) -and (Get-Date) -lt $exitDeadline){Start-Sleep -Milliseconds 500}
        if(Get-Process -Id $target.Id -ErrorAction SilentlyContinue){Record ('sample '+$sample+' target still running; next sample starts fresh');throw 'Target did not exit; close desktop OneNote before continuing'}
        $end=Get-Date
        Record ('sample '+$sample+' target exited after '+[int]($end-$started).TotalMilliseconds+' ms (launch-to-exit)')
        Write-Output ('['+$Label+' sample '+$sample+'] launch-to-exit '+[int]($end-$started).TotalMilliseconds+' ms')
    }
    Start-Sleep -Seconds 2
}
'Samples appended to '+$journal
