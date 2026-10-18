param(
    [string]$Version,
    [string]$Thumbprint,
    [string]$TimestampUrl='http://timestamp.digicert.com',
    [switch]$SkipBuild
)
# Builds, verifies, stages and packages a release archive for the open-source
# project. Never registers the add-in, launches OneNote or touches the clipboard.
# Optional Authenticode signing via -Thumbprint (Windows SDK signtool must be on
# PATH). Without a certificate the archive is unsigned, which is normal for an
# open-source project; SHA-256 checksums are always produced.
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
if(-not $Version){
    $file=Join-Path $root 'VERSION'
    if(Test-Path -LiteralPath $file){$Version=(Get-Content -LiteralPath $file -Raw).Trim()}
}
if(-not $Version){throw 'No version: pass -Version or create a VERSION file'}
$archiveName="OneNoteMarkdown-$Version-x64"

if(-not $SkipBuild){
    & (Join-Path $root 'build.ps1') -Test
    if($LASTEXITCODE -ne 0){throw 'build.ps1 -Test failed'}
}

$preview=Join-Path $root 'artifacts\preview'
$manifest=Join-Path $preview 'SHA256SUMS.txt'
if(-not (Test-Path -LiteralPath $manifest)){throw 'Run build.ps1 -Test first'}
foreach($line in [IO.File]::ReadAllLines($manifest)){
    if($line -notmatch '^([0-9A-F]{64})  ([A-Za-z0-9_.-]+)$'){throw 'Invalid integrity manifest'}
    $expected=$Matches[1];$name=$Matches[2]
    $path=Join-Path $preview $name
    if(-not (Test-Path -LiteralPath $path)){throw "Missing preview file: $name"}
    if((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $expected){throw "Integrity mismatch: $name"}
}

$stage=Join-Path $root ("artifacts\release\$archiveName")
if(Test-Path -LiteralPath $stage){Remove-Item -LiteralPath $stage -Recurse -Force}
New-Item -ItemType Directory -Path $stage -Force | Out-Null

$bin=Join-Path $stage 'bin'
New-Item -ItemType Directory -Path $bin -Force | Out-Null
Get-ChildItem -LiteralPath $preview -File -Filter *.dll | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $bin -Force }
Get-ChildItem -LiteralPath $preview -File -Filter *.LICENSE.txt | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $bin -Force }

foreach($item in @('LICENSE','THIRD-PARTY-NOTICES.md','README.md','usage.md','VERSION')){
    $src=Join-Path $root $item
    if(Test-Path -LiteralPath $src){Copy-Item -LiteralPath $src -Destination $stage -Force}
}
$scripts=Join-Path $stage 'tools'
New-Item -ItemType Directory -Path $scripts -Force | Out-Null
foreach($item in @('Install-Addin.ps1','Register-Preview.ps1')){
    $src=Join-Path $root ("tools\$item")
    if(Test-Path -LiteralPath $src){Copy-Item -LiteralPath $src -Destination $scripts -Force}
}
$docs=Join-Path $stage 'docs'
New-Item -ItemType Directory -Path $docs -Force | Out-Null
foreach($item in @('acceptance\install.md','acceptance\preview-session.md','acceptance\usage.md')){
    $src=Join-Path $root $item
    if(Test-Path -LiteralPath $src){Copy-Item -LiteralPath $src -Destination $docs -Force}
}

$dlls=@(Get-ChildItem -LiteralPath $bin -File -Filter *.dll)
$info=@()
$info+="name: OneNoteMarkdown Preview"
$info+="version: $Version"
$info+="builtUtc: $((Get-Date).ToUniversalTime().ToString('o'))"
$info+="platform: x64"
$info+="runtime: .NET Framework 4.8"
$info+="host: OneNote desktop (x64); the Store/UWP edition never loads this add-in"
$info+="loadBehavior: 0 (manual load)"
$info+=""
$info+="files:"
foreach($dll in $dlls){ $hash=(Get-FileHash -LiteralPath $dll.FullName -Algorithm SHA256).Hash; $info+="$hash  bin/$($dll.Name)" }
[IO.File]::WriteAllLines((Join-Path $stage 'VERSION-INFO.txt'),[string[]]$info,[Text.Encoding]::ASCII)

function Invoke-Sign([string[]]$paths){
    if(-not $Thumbprint){ return }
    $signtool=(Get-Command signtool.exe -ErrorAction SilentlyContinue).Source
    if(-not $signtool){ throw 'signtool.exe not found on PATH; install the Windows SDK or omit -Thumbprint' }
    foreach($path in $paths){
        & $signtool sign /sha1 $Thumbprint /fd sha256 /tr $TimestampUrl /td sha256 $path
        if($LASTEXITCODE -ne 0){ throw "Signing failed: $path" }
        & $signtool verify /pa $path
        if($LASTEXITCODE -ne 0){ throw "Signature verification failed: $path" }
    }
}
Invoke-Sign @($dlls | ForEach-Object { $_.FullName })

$stageSums=@()
foreach($file in @(Get-ChildItem -LiteralPath $stage -Recurse -File | Where-Object { $_.Name -ne 'SHA256SUMS.txt' } | Sort-Object FullName)){
    $relative=$file.FullName.Substring($stage.Length+1).Replace('\','/')
    $stageSums+="$((Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash)  $relative"
}
[IO.File]::WriteAllLines((Join-Path $stage 'SHA256SUMS.txt'),[string[]]$stageSums,[Text.Encoding]::ASCII)

$release=Join-Path $root 'artifacts\release'
$zip=Join-Path $release "$archiveName.zip"
if(Test-Path -LiteralPath $zip){Remove-Item -LiteralPath $zip -Force}
Compress-Archive -Path $stage -DestinationPath $zip -Force
$zipHash=(Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
[IO.File]::WriteAllText("$zip.sha256","$zipHash  $archiveName.zip`n",[Text.Encoding]::ASCII)

"Packaged $zip"
"  signed: $(if($Thumbprint){'yes'}else{'no (open-source default)'})"
"  zip sha256: $zipHash"
"  contents: $($stageSums.Count) files staged"
"Release packaging finished. No registry changes, host activation or clipboard access."
