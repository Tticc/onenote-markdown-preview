param(
    [int]$Seconds=20,
    [int]$IntervalMs=1000,
    [string]$Label='footprint'
)
# Read-only host footprint sampler. Never registers, launches or kills anything.
# Samples the desktop OneNote process and the preview surrogate dllhost:
#   - idle CPU% (processor-time delta / wall time / logical cores)
#   - working set and private bytes
# Conversion timing is not sampled here; it is traced by the add-in as
#   "conversion-ms=<n> chars=<m>" in %LOCALAPPDATA%\OneNoteMarkdownPreview\lifecycle.log
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$log=Join-Path $root 'artifacts\diagnostics\footprint.log'
$cores=[Environment]::ProcessorCount

function DesktopHosts(){ @(Get-Process -Name ONENOTE -ErrorAction SilentlyContinue | Where-Object { try{ $_.Path -and $_.Path -notlike '*\WindowsApps\*' }catch{ $false } }) }
function Surrogates(){ @(Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine -like '*E036DE73-418B-45F7-973F-A4B09D58E515*' } | ForEach-Object { Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue }) }

function Sample(){
    $hp=DesktopHosts; $sp=Surrogates
    $cpu=0.0; $ws=0L; $priv=0L
    foreach($p in @($hp)+@($sp)){ if($p){ $cpu+=$p.TotalProcessorTime.TotalSeconds; $ws+=$p.WorkingSet64; $priv+=$p.PrivateMemorySize64 } }
    [pscustomobject]@{ Host=($hp.Count); Surrogates=($sp.Count); Cpu=($cpu); Ws=($ws); Priv=($priv) }
}

$first=Sample
if($first.Host -eq 0){ throw 'No desktop OneNote process; start it and enable the add-in first' }

$cpuMax=0.0; $wsMax=$first.Ws; $privMax=$first.Priv
$prev=$first; $prevTime=Get-Date
$rows=@()
for($i=0; $i -lt $Seconds; $i++){
    Start-Sleep -Milliseconds $IntervalMs
    $now=Sample; $t=Get-Date
    $wall=($t-$prevTime).TotalSeconds
    $cpuPct=if($wall -gt 0){ (($now.Cpu-$prev.Cpu)/$wall/$cores)*100 } else { 0 }
    if($cpuPct -gt $cpuMax){ $cpuMax=$cpuPct }
    if($now.Ws -gt $wsMax){ $wsMax=$now.Ws }
    if($now.Priv -gt $privMax){ $privMax=$now.Priv }
    $rows+=[pscustomobject]@{ T=('[+{0:00}s]' -f ($i+1)); CpuPct=('{0:N2}' -f $cpuPct); WsMB=('{0:N1}' -f ($now.Ws/1MB)); PrivMB=('{0:N1}' -f ($now.Priv/1MB)); Host=$now.Host; Surrogates=$now.Surrogates }
    $prev=$now; $prevTime=$t
}

"=== $Label ($Seconds s, ${IntervalMs}ms interval, $cores logical cores) ==="
$rows | Format-Table -AutoSize
$summary="label=$Label samples=$Seconds cores=$cores cpuMaxPct={0:N2} wsMaxMB={1:N1} privMaxMB={2:N1} host={3} surrogates={4}" -f $cpuMax,($wsMax/1MB),($privMax/1MB),$first.Host,$first.Surrogates
$summary
[IO.File]::AppendAllText($log,(Get-Date).ToUniversalTime().ToString('o')+" $summary"+[Environment]::NewLine)
"footprint logged to $log"
