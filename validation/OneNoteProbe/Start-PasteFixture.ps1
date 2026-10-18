param([switch]$CheckOnly, [switch]$ResetDisposable, [switch]$BlockFixture, [switch]$CodeFixture, [switch]$FencedFixture, [switch]$ExplicitWhitespace, [switch]$RtfCode, [switch]$RemainingFixture, [ValidateSet(200,500)][int]$RestoreDelayMs=200)
$ErrorActionPreference = 'Stop'
$output = Join-Path $PSScriptRoot 'offline'
if (-not (Test-Path -LiteralPath $output)) { throw 'Missing build output' }
if (@(Get-Process -Name PasteFixture -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Existing fixture still running' }
$compiler = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$interop = "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll"
$exe = Join-Path $output 'PasteFixture.exe'
$sources = @('PasteFixture.cs','ClipboardRoundtrip.cs','HtmlPayloadTests.cs','SelectionScreen.cs','SelectionEvidence.cs','FixtureText.cs','InlinePostcondition.cs','OutsideDiff.cs','RestorePolicy.cs','BlockFixture.cs','CodeWhitespaceFixture.cs','RtfCodeFixture.cs') | ForEach-Object { Join-Path $PSScriptRoot $_ }
& $compiler /nologo /platform:x64 /main:PasteFixture "/out:$exe" "/reference:$interop" /reference:System.Windows.Forms.dll $sources
if ($LASTEXITCODE -ne 0) { throw 'Paste fixture compile failed' }
& $exe --check-input-layout
if ($LASTEXITCODE -ne 0) { throw 'Keyboard input layout checks failed' }
if ($CheckOnly) { return }
$log = Join-Path $output 'paste-fixture.log'
$eventName = 'Local\OneNoteMarkdownPasteFixtureAck'
$arguments = @(('"{0}"' -f $log), ('"{0}"' -f $eventName))
$arguments += "--delay=$RestoreDelayMs"
if ($ResetDisposable) { $arguments += '--reset-disposable' }
if ($BlockFixture) { $arguments += '--block-fixture' }
if ($CodeFixture) { $arguments += '--code-fixture' }
if ($FencedFixture) { $arguments += '--fenced-fixture' }
if ($ExplicitWhitespace) { throw 'NBSP candidate disabled: strict original whitespace required' }
if ($RtfCode) {if(-not $FencedFixture){throw 'RTF code requires fenced fixture'}; $arguments += '--rtf-code'}
if ($RemainingFixture) { $arguments += '--remaining-fixture' }
$child = Start-Process -FilePath $exe -ArgumentList $arguments -WindowStyle Hidden -PassThru
"Started bounded authorized paste fixture PID=$($child.Id); log=$log"
