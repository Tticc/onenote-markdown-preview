# OneNote metadata probe

> **Host loading suspended:** an earlier probe attempt coincided with a OneNote shutdown incident. Registration is now blocked by the script; `-Remove` remains available. Historical host instructions below are not authorization to repeat the test. Use `Test-Offline.ps1` for the corrected source. The old root-level DLL/EXE are historical binaries and must not be registered.

This is a phase-one validation executable, not the product Add-in. It uses the installed Office interop assembly and .NET Framework compiler. It reads window metadata only, does not read page content, register an Add-in, send input, or modify the clipboard. COM activation may start OneNote if it is not already running; run with OneNote open.

Run in Windows PowerShell from the repository root:

```powershell
$compiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$interop = 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll'
$output = Join-Path $env:TEMP 'OneNoteMetadataProbe.exe'
& $compiler /nologo /platform:x64 "/out:$output" "/reference:$interop" '.\validation\OneNoteProbe\Probe.cs'
if ($LASTEXITCODE -eq 0) { & $output }
```

Exit codes: 0 = metadata accessible and current page is named `Markdown插件验证`; 1 = exception (type and HRESULT only); 2 = not exactly one window; 3 = fixture title mismatch. Only page hierarchy metadata is read. A zero exit code does not validate precise selection, clipboard restoration, Add-in loading, or undo.

The interop path is specific to the inspected machine. No SDK or package download was used for this probe.

## Temporary COM Add-in

`ProbeAddin.cs` supplies only a Ribbon loading/callback probe. Build with the same compiler using `/target:library /platform:x64 /out:ProbeAddin.dll`, with output in this directory. Run `Register-Probe.ps1` in 64-bit Windows PowerShell to register for the current user. It refuses to replace existing registration.

Restart OneNote manually, select **Markdown Probe → Verify callback**, then inspect `%LOCALAPPDATA%\OneNoteMarkdownValidation\probe.log`. It contains callback names, UTC timestamps, PID and elapsed time, not notebook contents. This is not the formatting plugin and has no document-write or clipboard code.

Rollback: run `Register-Probe.ps1 -Remove`, then restart OneNote. This removes only the CLSID, ProgID, AppID and Office Addins registration roots belonging to the validation probe; logs and source/binary remain for inspection. Do not move/delete the DLL while registered.

Registration includes AppID/DllSurrogate for out-of-process COM activation. `SurrogateProbe.cs` separately tests CLSCTX_LOCAL_SERVER activation and the Ribbon interface; build it beside ProbeAddin.dll with a reference to that DLL. Its GetCustomUI log is not a OneNote host callback. Correlate PIDs and timestamps when checking actual loading.

## Corrected offline contract verification

Run `Test-Offline.ps1`. It references the installed official Extensibility and Office PIAs, builds corrected files into `offline/`, and runs managed contract checks without COM activation or contacting OneNote. It invokes lifecycle methods with dummy values, so content-free log entries will have the offline test PID. Those entries do not establish host loading.

The original handwritten interfaces incorrectly used InterfaceIsIDispatch where the installed official contracts are Dual; lifecycle array direction/marshaling and enum types also differed. Corrected source implements official IDTExtensibility2/IRibbonExtensibility and uses IRibbonControl for the callback. Passing these checks does not verify cross-process marshaling, registration, host exit or recovery. A revised host test plan is required before lifting the registration block.

`Test-Offline.ps1` also builds and runs `CrossProcessTests.cs`: a private current-user pipe carries a standard COM marshaled interface from a child STA process. The parent invokes official lifecycle interfaces with a nonempty VARIANT SAFEARRAY, queries the official Ribbon interface, reads its BSTR result, invokes shutdown, releases the proxy and verifies normal child exit. The server pumps its STA messages. No registry entry or COM class activation is used, and no OneNote object is created. A 20-second watchdog terminates only the spawned test process and fails the test if it stalls; each normal exit is checked explicitly.

Limits: OnConnection receives null application/AddIn objects; this does not test real host object marshaling, Ribbon button IDispatch callbacks, COM activation/surrogate policy, OneNote startup overhead or exit. Test pipe framing is fixture code, not the product IPC implementation. A passing test does not authorize host registration.

### Historical expanded suite failure

The latest fixture supplies nonnull synthetic application/AddIn objects, which successfully reach OnConnection. It also tries raw native IDispatch button invocation with a synthetic IRibbonControl, expecting a reverse Id property call. This path currently fails with 0x8002000E (parameter count) before the button method is entered. Test-Offline.ps1 correctly exits with failure. The prior lifecycle-only suite passing does not mean this expanded suite passes. Host registration remains blocked pending a review of the native invocation harness and default-dispatch contract. Synthetic objects are not actual Office objects.

### Current expanded suite status: passing

The historical callback failure above was caused by missing array size metadata in the test client's GetIDsOfNames declaration: output DISPID remained 0. Explicit LPArray/SizeParamIndex fixes now yield the actual callback DISPID. The expanded suite successfully invokes the callback through native IDispatch and verifies one reverse IRibbonControl.Id access. Unknown-name and missing-argument negative cases also pass; the CCW rejects missing arguments with E_INVALIDARG on this machine. Full startup/shutdown lifecycle and normal child exit pass. This remains a synthetic, registration-free test; no OneNote host compatibility is established. New registration is still blocked.

`Test-Activation.ps1` is a separate activation-only fixture: temporarily registers the corrected offline DLL's CLSID/AppID, never creates an Office Addins entry, invokes official lifecycle/Ribbon calls through CLSCTX_LOCAL_SERVER, waits for normal client/surrogate exit and removes registration in finally. It does not authorize or perform OneNote loading. Run only after offline tests have built the corrected DLL and when the probe CLSID/AppID are absent.

## Read-only selection export

`Test-Selection.ps1` builds and runs SelectionProbe.cs. Requires an already-visible single OneNote window on the disposable page titled `Markdown插件验证`. Without a visible window it skips before COM activation. The fixture page is exported in memory with piSelection; output contains structural counts and selection attribute metadata only. No note text or page IDs are printed, no XML is saved, and no clipboard/document edits are performed. Keep the fixture small; the post-export length check does not bound host-side export cost. Success means XML was inspected, not that exact selection precision has passed.

The probe now uses SelectionScreen.cs and SelectionEvidence.cs for conservative structural screening and an optional focused double-read comparison. Full-export hashes stay in memory; a candidate is not paste approval. Foreground host and DocumentCanvas focus are required; default OneNote styles and block endpoints remain unresolved. Test-Offline.ps1 includes 26 synthetic screening/evidence cases.

`Test-ClipboardInventory.ps1` only enumerates clipboard format metadata under a short lock. It does not fetch delayed data, take a backup, change clipboard content or paste. A successful inventory does not prove snapshot/restore compatibility.

The focused evidence probe normalizes the COM window handle with GA_ROOT before comparing foreground roots: this OneNote build's COM handle is not itself the foreground top-level handle. PID-only matching is not used as approval. HtmlPayloadTests.cs adds four offline fixed-fragment CF_HTML UTF-8 offset/whitespace tests, not a Markdown implementation or OneNote import test.

`Test-ClipboardRoundtrip.ps1 [-DelaySeconds 0..60]` is an explicitly authorized disposable-fixture experiment. Requires exact Unicode text clipboard-probe-123 and only Unicode/ANSI/OEM text/locale formats; private/HTML/GDI formats cause refusal before any write. It temporarily publishes a test string and restores independent memory copies, checking sequence under the restoration lock and original format bytes. It does not interact with OneNote. This is not production-safe recovery code; if publishing/restoration fails, manually recopy the disposable fixture. Never run it on important clipboard content or infer broad format support from a pass.

With explicit additional user permission, -SyntheticDisposable replaces the disposable original (including its metadata) with a generated Unicode fixture first. -SimulateNewerCopy requires that switch and tests another owner HWND in the same process; it deliberately leaves a newer test string. Normal rerun resets/restores the baseline fixture. Sequence is captured after CloseClipboard and checked together with owner HWND under the restore lock, accounting for Windows text synthesis. This is limited fixture evidence, not a solution for arbitrary/OLE clipboard preservation or paste completion.

`Test-ExportComparison.ps1` is read-only on the named fixture: compares piBasic and piSelection exports and repeated selection exports using metadata categories only. Observed mode-dependent child-count differences. The paste fixture observer now compares piBasic baseline/post-state (with piSelection for original target location) instead of assuming selection-marked node shapes are stable after a paste collapses selection. No mismatch category is ignored to force a pass; actual observer success still requires a controlled host test.

Current approved policy supersedes mandatory observer completion: Start-PasteFixture.ps1 defaults to -RestoreDelayMs 200, accepts 500, restores only after owner/sequence checks, and no longer waits for human acknowledgement. This is empirical timing, not a completion guarantee. The standalone fixture's sleep is not a product UI implementation. -CheckOnly compiles/checks inputs without clipboard/input writes; real execution still requires explicit fixture selection and disposable clipboard reset authorization. Arbitrary clipboard/OLE snapshot support remains incomplete.

`Test-CommonClipboard.ps1` overwrites clipboard with synthetic text, HTML/RTF, PNG and DIB fixtures and leaves the last 2x2 image. Explicit user authorization required. Verifies supported bytes through independent restore, DIB pixels/dimensions and competing-process new-copy protection. It never accesses OneNote. Snapshot code is fixture-only, rejects unknown/private formats and does not prove arbitrary app-origin data, delayed rendering, alpha/color profiles or OLE metadata preservation.

`Test-OleInventory.ps1` launches a separate five-second bounded read-only process. Enumerates OLE source FORMATETC/TYMED and independently materializes approved HTML/RTF/text/DIB HGLOBAL representations, printing only format names, medium kinds and byte sizes. It releases all media/COM objects, clears copied buffers and does not publish anything. This can request delayed rendering from the original app; supervisor may terminate only its own probe on timeout. It does not copy protection/private/vector data or prove restoration.

The optional `-Roundtrip` mode is MUTATING: after authorized common-content snapshot it temporarily publishes text and restores after 200 ms, verifies saved representation bytes, and deliberately omits conditional OneNote-native/vector extras when appropriate common fallbacks exist. Unknown/protection formats and unsupported media refuse. It must not be confused with default read-only mode. Tested on one small real OneNote sample. Restoration failure/timeout can leave a temporary fixture; not product recovery code. This mode does not contact OneNote page APIs.

Code blocks require exact whitespace. NBSP candidate is disabled for actual paste. RtfCodeFixture.cs has nine RichEdit-only roundtrip cases; Start-PasteFixture.ps1 -FencedFixture -RtfCode selects an RTF-only rich payload plus Unicode fallback for the same fixed fenced sample. It does not establish OneNote whitespace preservation until actual import is inspected. No language highlighting, image rendering or whitespace substitution is performed.

`Run-AutomatedValidation.ps1` is the consolidated non-mutating entry: runs offline COM, selection, payload, policy, restoration fault/ordering and resource-limit tests; check-only builds input fixtures; audits probe registry/process residue. Optional `-ReadOnlyHost` runs fixture export comparison without clipboard or document changes. Mutating tests are NOT invoked by this runner. A pass is not phase-one release or complete plugin acceptance. RestoreCoordinator uses a fake backend for fault injection; SnapshotLimits is applied before independent OLE buffer allocation.
