$ErrorActionPreference = 'Stop'
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$extensibility = "$env:WINDIR\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll"
$office = "$env:WINDIR\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll"
foreach ($path in @($compiler, $extensibility, $office)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing build prerequisite: $path" }
}
# Separate output: never overwrite a previously registered/loaded probe binary.
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $PSScriptRoot)) { throw 'Missing source directory' }
if (-not (Test-Path -LiteralPath $output)) { New-Item -ItemType Directory -Path $output | Out-Null }
$dll = Join-Path $output 'ProbeAddin.dll'
$exe = Join-Path $output 'ContractTests.exe'
& $compiler /nologo /target:library /platform:x64 "/out:$dll" "/reference:$extensibility" "/reference:$office" (Join-Path $PSScriptRoot 'ProbeAddin.cs')
if ($LASTEXITCODE -ne 0) { throw 'Probe compilation failed' }
& $compiler /nologo /platform:x64 "/out:$exe" "/reference:$dll" "/reference:$extensibility" "/reference:$office" (Join-Path $PSScriptRoot 'ContractTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed' }
& $exe
if ($LASTEXITCODE -ne 0) { throw 'Offline contract tests failed' }
$cross = Join-Path $output 'CrossProcessTests.exe'
& $compiler /nologo /platform:x64 "/out:$cross" "/reference:$dll" "/reference:$extensibility" "/reference:$office" /reference:System.Windows.Forms.dll (Join-Path $PSScriptRoot 'CrossProcessTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Cross-process test compilation failed' }
& $cross
if ($LASTEXITCODE -ne 0) { throw 'Cross-process tests failed' }
$screen = Join-Path $output 'SelectionScreenTests.exe'
& $compiler /nologo /platform:x64 "/out:$screen" (Join-Path $PSScriptRoot 'SelectionScreen.cs') (Join-Path $PSScriptRoot 'SelectionEvidence.cs') (Join-Path $PSScriptRoot 'SelectionScreenTests.cs') (Join-Path $PSScriptRoot 'FixtureText.cs')
if ($LASTEXITCODE -ne 0) { throw 'Selection screening tests compile failed' }
& $screen
if ($LASTEXITCODE -ne 0) { throw 'Selection screening tests failed' }
$html = Join-Path $output 'HtmlPayloadTests.exe'
& $compiler /nologo /platform:x64 "/out:$html" (Join-Path $PSScriptRoot 'HtmlPayloadTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'HTML fixture tests compile failed' }
& $html
if ($LASTEXITCODE -ne 0) { throw 'HTML fixture tests failed' }
$post = Join-Path $output 'InlinePostconditionTests.exe'
& $compiler /nologo /platform:x64 "/out:$post" (Join-Path $PSScriptRoot 'InlinePostcondition.cs') (Join-Path $PSScriptRoot 'InlinePostconditionTests.cs') (Join-Path $PSScriptRoot 'FixtureText.cs') (Join-Path $PSScriptRoot 'SelectionScreen.cs') (Join-Path $PSScriptRoot 'OutsideDiff.cs')
if ($LASTEXITCODE -ne 0) { throw 'Inline postcondition compile failed' }
& $post
if ($LASTEXITCODE -ne 0) { throw 'Inline postcondition tests failed' }
$restore = Join-Path $output 'RestorePolicyTests.exe'
& $compiler /nologo /platform:x64 "/out:$restore" (Join-Path $PSScriptRoot 'RestorePolicy.cs') (Join-Path $PSScriptRoot 'RestorePolicyTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Restore policy compile failed' }
& $restore
if ($LASTEXITCODE -ne 0) { throw 'Restore policy tests failed' }
$formats = Join-Path $output 'ClipboardFormatPolicyTests.exe'
& $compiler /nologo /platform:x64 "/out:$formats" (Join-Path $PSScriptRoot 'ClipboardFormatPolicy.cs') (Join-Path $PSScriptRoot 'ClipboardFormatPolicyTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Clipboard format policy compile failed' }
& $formats
if ($LASTEXITCODE -ne 0) { throw 'Clipboard format policy tests failed' }
$block = Join-Path $output 'BlockFixtureTests.exe'
$sources = @('BlockFixture.cs','BlockFixtureTests.cs','SelectionScreen.cs','FixtureText.cs','CodeWhitespaceFixture.cs') | ForEach-Object { Join-Path $PSScriptRoot $_ }
& $compiler /nologo /platform:x64 "/out:$block" $sources
if ($LASTEXITCODE -ne 0) { throw 'Block fixture compile failed' }
& $block
if ($LASTEXITCODE -ne 0) { throw 'Block fixture tests failed' }
$rtf = Join-Path $output 'RtfCodeFixtureTests.exe'
& $compiler /nologo /platform:x64 "/out:$rtf" /reference:System.Windows.Forms.dll (Join-Path $PSScriptRoot 'RtfCodeFixture.cs') (Join-Path $PSScriptRoot 'RtfCodeFixtureTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'RTF code fixture build failed' }
& $rtf
if ($LASTEXITCODE -ne 0) { throw 'RTF code fixture tests failed' }
$coordinator = Join-Path $output 'RestoreCoordinatorTests.exe'
& $compiler /nologo /platform:x64 "/out:$coordinator" (Join-Path $PSScriptRoot 'RestoreCoordinator.cs') (Join-Path $PSScriptRoot 'RestoreCoordinatorTests.cs') (Join-Path $PSScriptRoot 'RestorePolicy.cs')
if ($LASTEXITCODE -ne 0) { throw 'Restore coordinator build failed' }
& $coordinator
if ($LASTEXITCODE -ne 0) { throw 'Restore coordinator tests failed' }
$limits = Join-Path $output 'SnapshotLimitsTests.exe'
& $compiler /nologo /platform:x64 "/out:$limits" (Join-Path $PSScriptRoot 'SnapshotLimits.cs') (Join-Path $PSScriptRoot 'SnapshotLimitsTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Snapshot limits build failed' }
& $limits
if ($LASTEXITCODE -ne 0) { throw 'Snapshot limits tests failed' }
