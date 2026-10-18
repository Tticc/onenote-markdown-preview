param([switch]$Test,[string]$OutputDirectory='artifacts\preview',[string]$Version)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$rawVersion=$Version
if(-not $rawVersion){
    $versionFile=Join-Path $root 'VERSION'
    if(Test-Path -LiteralPath $versionFile){$rawVersion=(Get-Content -LiteralPath $versionFile -Raw).Trim()}
}
if($rawVersion -match '^\d+\.\d+\.\d+$'){$versionStamp="$rawVersion.0"}
elseif($rawVersion -match '^\d+\.\d+\.\d+\.\d+$'){$versionStamp=$rawVersion}
else{$versionStamp='0.0.0.0'}
$packages=Join-Path $root '.packages'
$output=Join-Path $root $OutputDirectory
foreach($directory in @($packages,$output)){if(-not(Test-Path -LiteralPath $directory)){New-Item -ItemType Directory -Path $directory -Force|Out-Null}}
$versionSource=Join-Path $output 'Version.g.cs'
[IO.File]::WriteAllText($versionSource,"using System.Reflection;`r`n[assembly: AssemblyVersion(`"$versionStamp`")]`r`n[assembly: AssemblyFileVersion(`"$versionStamp`")]`r`n",[Text.Encoding]::ASCII)
$dependencies=@(
    @{Id='Markdig';Version='0.37.0';Dll='lib\net462\Markdig.dll'},
    @{Id='System.Memory';Version='4.5.5';Dll='lib\net461\System.Memory.dll'},
    @{Id='System.Buffers';Version='4.5.1';Dll='lib\net461\System.Buffers.dll'},
    @{Id='System.Numerics.Vectors';Version='4.5.0';Dll='lib\net46\System.Numerics.Vectors.dll'},
    @{Id='System.Runtime.CompilerServices.Unsafe';Version='4.5.3';Dll='lib\net461\System.Runtime.CompilerServices.Unsafe.dll'}
)
$references=@()
foreach($dependency in $dependencies){
    $id=$dependency.Id.ToLowerInvariant();$version=$dependency.Version
    $directory=Join-Path $packages "$id.$version"
    $dll=Join-Path $directory $dependency.Dll
    if(-not(Test-Path -LiteralPath $dll)){
        $zip=Join-Path $packages "$id.$version.zip"
        & curl.exe -fsSL "https://api.nuget.org/v3-flatcontainer/$id/$version/$id.$version.nupkg" -o $zip
        if($LASTEXITCODE -ne 0){throw "Package download failed: $id"}
        Expand-Archive -LiteralPath $zip -DestinationPath $directory -Force
    }
    $references+="/reference:$dll"
    Copy-Item -LiteralPath $dll -Destination $output -Force
    $license=Join-Path $directory 'LICENSE.TXT'
    if(Test-Path -LiteralPath $license){Copy-Item -LiteralPath $license -Destination (Join-Path $output ($dependency.Id+'.LICENSE.txt')) -Force}
}
$compiler="$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$core=Join-Path $output 'OneNoteMarkdown.Core.dll'
$sources=@('MarkdownConverter.cs','HtmlClipboard.cs','OperationState.cs','TextFragment.cs','ParagraphPolicy.cs')|ForEach-Object{Join-Path $root "src\Core\$_"};$sources+=$versionSource
& $compiler /nologo /target:library /platform:x64 /optimize+ "/out:$core" $references $sources
if($LASTEXITCODE -ne 0){throw 'Core build failed'}
$ext="$env:WINDIR\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll"
$office="$env:WINDIR\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll"
$one="$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll"
foreach($required in @($ext,$office,$one)){if(-not(Test-Path -LiteralPath $required)){throw "Missing Office PIA: $required"}}
$addin=Join-Path $output 'OneNoteMarkdown.Addin.dll'
$addinSources=@('Native.cs','Selection.cs','ClipboardSession.cs','Addin.cs')|ForEach-Object{Join-Path $root "src\Addin\$_"};$addinSources+=$versionSource
& $compiler /nologo /target:library /platform:x64 /optimize+ "/out:$addin" "/reference:$core" "/reference:$ext" "/reference:$office" "/reference:$one" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll $addinSources
if($LASTEXITCODE -ne 0){throw 'Addin build failed'}
if($Test){
    $exe=Join-Path $output 'OneNoteMarkdown.Tests.exe'
    & $compiler /nologo /platform:x64 "/out:$exe" "/reference:$core" (Join-Path $root 'tests\CoreTests.cs')
    if($LASTEXITCODE -ne 0){throw 'Core test build failed'}
    & $exe
    if($LASTEXITCODE -ne 0){throw 'Core tests failed'}
    $audit=Join-Path $output 'OneNoteMarkdown.PreviewAudit.exe'
    & $compiler /nologo /platform:x64 "/out:$audit" "/reference:$addin" "/reference:$ext" "/reference:$office" (Join-Path $root 'tests\PreviewAudit.cs')
    if($LASTEXITCODE -ne 0){throw 'Preview audit build failed'}
    & $audit $output
    if($LASTEXITCODE -ne 0){throw 'Preview audit failed'}
    $lifecycle=Join-Path $output 'OneNoteMarkdown.LifecycleTests.exe'
    & $compiler /nologo /platform:x64 "/out:$lifecycle" "/reference:$addin" "/reference:$core" "/reference:$ext" "/reference:$office" "/reference:$one" /reference:System.Windows.Forms.dll (Join-Path $root 'tests\LifecycleTests.cs')
    if($LASTEXITCODE -ne 0){throw 'Lifecycle test build failed'}
    $child=Start-Process -FilePath $lifecycle -ArgumentList ('"{0}"' -f $output) -NoNewWindow -PassThru
    $testHandle=$child.Handle
    try{
        if(-not $child.WaitForExit(20000)){$child.Kill();throw 'Offline lifecycle tests timed out; only test process stopped'}
        if($child.ExitCode -ne 0){throw 'Offline lifecycle tests failed'}
    }finally{$child.Dispose()}
    $cross=Join-Path $output 'OneNoteMarkdown.CrossProcessLifecycle.exe'
    & $compiler /nologo /platform:x64 /define:PRODUCT_TEST "/out:$cross" "/reference:$addin" "/reference:$core" "/reference:$ext" "/reference:$office" "/reference:$one" /reference:System.Windows.Forms.dll (Join-Path $root 'validation\OneNoteProbe\CrossProcessTests.cs')
    if($LASTEXITCODE -ne 0){throw 'Actual product cross-process test build failed'}
    & $cross
    if($LASTEXITCODE -ne 0){throw 'Actual product cross-process lifecycle failed'}
    $nativeRefs=Join-Path $output 'OneNoteMarkdown.NativeReferenceTests.exe'
    & $compiler /nologo /platform:x64 "/out:$nativeRefs" "/reference:$addin" "/reference:$core" "/reference:$ext" "/reference:$office" "/reference:$one" /reference:System.Windows.Forms.dll (Join-Path $root 'tests\NativeReferenceTests.cs')
    if($LASTEXITCODE -ne 0){throw 'Native reference fixture build failed'}
    $child=Start-Process -FilePath $nativeRefs -ArgumentList ('"{0}"' -f $output) -NoNewWindow -PassThru
    $testHandle=$child.Handle
    try{
        if(-not $child.WaitForExit(15000)){$child.Kill();throw 'Native ref fixture timeout; only test process stopped'}
        if($child.ExitCode -ne 0){throw 'Native ref fixture failed'}
    }finally{$child.Dispose()}
}
Copy-Item -LiteralPath (Join-Path $root 'licenses\Markdig.txt') -Destination (Join-Path $output 'Markdig.LICENSE.txt') -Force
$names=@('OneNoteMarkdown.Core.dll','OneNoteMarkdown.Addin.dll','Markdig.dll','System.Memory.dll','System.Buffers.dll','System.Numerics.Vectors.dll','System.Runtime.CompilerServices.Unsafe.dll')
$lines=foreach($name in $names){$hash=(Get-FileHash -LiteralPath (Join-Path $output $name) -Algorithm SHA256).Hash;"$hash  $name"}
[IO.File]::WriteAllLines((Join-Path $output 'SHA256SUMS.txt'),[string[]]$lines,[Text.Encoding]::ASCII)
'Offline build finished. No registry changes, host activation or clipboard access.'
