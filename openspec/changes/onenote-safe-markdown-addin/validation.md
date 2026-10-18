# Phase-one validation record

## STOP: host exit incident and rollback

After the second restart attempt the user reported the Ribbon was still absent and OneNote could not exit. Stop all Add-in loading and editing experiments until normal host operation is restored and the activation path is reviewed.

- Immediately ran Register-Probe.ps1 -Remove. Confirmed this probe's Office Addins, CLSID and AppID keys are absent. The script also removes its ProgID key.
- Observed two ONENOTE processes: PID 28732 without a main window and PID 19780 with a main window; both reported Responding=True. This flag does not establish that application shutdown works.
- Enumerated dllhost command lines; none matched this probe's AppID. Did not terminate any dllhost or OneNote process.
- probe.log still contains only the earlier standalone surrogate GetCustomUI call; no OneNote OnConnection or shutdown callback was recorded. The cause of the exit issue remains unproven; lack of logs does not exclude an activation-time problem.
- No document or clipboard writes were performed by the probe. This does not prove that all existing user edits are synced; do not force-kill OneNote without discussing possible unsaved/unsynced work with the user.
- Next step: user checks whether the visible window can now close normally after unregistering. If not, obtain a recovery decision before force termination. Do not re-register the probe or ask for another loading test.
- Task 1.2 remains incomplete and is now a blocking host lifecycle incident. Phases 2–6 remain gated.

### Recovery action

The user reported the cleanup dialog remained for ten minutes and neither Close nor Cancel worked. After explaining possible unsaved-edit loss, the user explicitly authorized termination. Stopped the remaining ONENOTE process(es) with Stop-Process -Force and verified zero remained after two seconds. Reconfirmed this probe's Office Addins, CLSID and AppID keys are absent. No cache, notebook files or other Add-in registrations were removed. Normal startup, recent-note integrity, sync and normal exit still need user verification; do not claim recovery complete or resume plugin tests until those checks pass.

### User-confirmed recovery and offline contract review

User subsequently confirmed OneNote starts/exits normally and inspected content was not lost. Sync status was not explicitly reported. This is user-reported recovery evidence, not proof of the incident's root cause.

Offline reflection inspected installed official assemblies:

- `Extensibility/7.0.3300.0`: IDTExtensibility2 GUID B65AD801-ABAF-11D0-BB8B-00A0C90F2744, ComImport, FDual/FDispatchable. Lifecycle parameters include typed enums and [In, MarshalAs(SafeArray, SafeArraySubType=VT_VARIANT)] ref Array.
- `office/15.0.0.0`: IRibbonExtensibility GUID 000C0396-0000-0000-C000-000000000046, ComImport, FDual/FDispatchable; IRibbonControl is the official button callback parameter.
- Original source instead declared those same interface IIDs as InterfaceIsIDispatch and omitted official lifecycle marshaling. Reusing an IID with an incompatible interface layout is a definite contract defect. Prior self-declared-interface activation/GetCustomUI tests did not establish compatibility with the host's official interfaces. This defect is a plausible loading-failure cause, not proof of the ten-minute shutdown incident's sole cause.

Corrected source uses the official PIAs. `Test-Offline.ps1` built separate `offline/` binaries and passed ten checks covering imported interface identity, Dual flags, complete method maps, callback parameter and well-formed Ribbon XML. Managed dummy lifecycle sequence completed. Tests do not contact OneNote or register/activate COM. Their diagnostic entries must not be mistaken for actual host callbacks. Original root binaries remain historical and are not updated/registered.

`Register-Probe.ps1` now unconditionally blocks new registration after handling -Remove, preventing accidental reuse of the failed probe. No registration was restored. A revised test plan, official-interface cross-process verification outside OneNote, lifecycle evidence and an explicit host-test decision are required before any further host loading. Task 1.2 remains blocked; no task is marked complete on offline evidence alone.

### Registration-free cross-process verification

Implemented CrossProcessTests.cs and invoked it via Test-Offline.ps1. First compile found an ambiguous Timer type, corrected by qualifying System.Threading.Timer. The final run passed the existing ten managed contract checks plus the following cross-process checks (client PID 17016, server PID 2716):

- Standard CoMarshalInterface/CoUnmarshalInterface transfer of the official IDTExtensibility2 through a private current-user pipe, with no registry entries or COM class activation.
- OnConnection, OnAddInsUpdate, OnStartupComplete with a nonempty VARIANT SAFEARRAY.
- Query of official IRibbonExtensibility and cross-process GetCustomUI BSTR result.
- OnBeginShutdown and OnDisconnection; client releases proxy, sends explicit stop, and observes normal child exit code 0 within three seconds.

The test uses a separate child STA with a message pump and a 20-second failure watchdog. No watchdog termination occurred in the passing run. It does not reference/create OneNote application objects or use the clipboard. Logs produced are synthetic test callbacks, not evidence of OneNote loading.

Remaining limitations: null application and AddInInst inputs, no real Ribbon button callback marshaling, no class activation or surrogate registration path, no real host shutdown or startup measurement. Therefore 1.2 stays incomplete. Host registration remains blocked; any further host test needs an explicit decision and a narrowed test plan. Future options include a controlled automation harness exercising IDispatch callbacks and nonnull test host objects before considering OneNote again.

### Nonnull objects and native callback test — FAIL / PAUSED

Extended the fixture with two nonnull AutoDispatch host objects and an official IRibbonControl implementation. OnConnection logs confirm both objects arrived at the child. Startup lifecycle and Ribbon XML still succeeded. Added a raw native IDispatch GetIDsOfNames/Invoke test intended to call VerifyCallback with a VARIANT holding the test control and observe a reverse Id getter call.

After fixing ambiguous DISPPARAMS/EXCEPINFO names at compilation, the callback path returned TargetParameterCountException HRESULT 0x8002000E. The child log did not enter VerifyCallback. Changed initial interface marshaling from lifecycle interface to default IDispatch and reran; the same error remained (client 24104/server 27184). This does not establish whether the defect is the harness ABI, CCW default dispatch selection, or callback exposure. It is not a passing callback test and not proof of the original OneNote shutdown cause.

Final Test-Offline.ps1 result is FAILURE; earlier passing lifecycle-only evidence remains historical, not the result of the current expanded suite. Child processes were cleaned up by the client failure path; shutdown callbacks in this failing run were not reached, so do not label cleanup as normal COM lifecycle shutdown. Host registration remains blocked and task 1.2 remains incomplete. Pause this integration path for test-harness/dispatch-contract review before further activation tests.

### Callback harness diagnosis and final passing expanded suite

The above failure is resolved in the offline harness. Instrumentation showed GetIDsOfNames returned success but the managed output array retained DISPID 0. The declaration lacked explicit LPArray/SizeParamIndex=2 on the output DISPID array (and explicit size on names). Correcting the declaration returned callback DISPID 1610743818; the same Invoke path then successfully entered VerifyCallback and made one reverse cross-process IRibbonControl.Id call. This identifies a separate client harness marshaling defect; it is not an additional demonstrated defect in the corrected plugin or evidence about the host shutdown incident.

Added negative checks: unknown name returns DISP_E_UNKNOWNNAME and DISPID_UNKNOWN; missing callback argument returns failure without calling the control getter. The latter actually returns E_INVALIDARG (0x80070057) on this CCW, not the initially assumed DISP_E_BADPARAMCOUNT; the assertion checks rejection and absence of callback rather than requiring an undocumented exact error mapping.

Final Test-Offline.ps1 run PASS: ten managed contract checks, nonnull synthetic objects, official cross-process lifecycle/Ribbon calls, native default IDispatch callback, reverse getter, negative cases and normal shutdown/child exit (client 15584/server 9552). Initial marshaled interface is now default IDispatch; official lifecycle/Ribbon interfaces are queried from the same object. No registry or OneNote operations occurred. No watchdog termination in final run.

Offline callback investigation is complete for this fixture. Not covered: actual Office objects, COM class activation/registration/surrogate policy, actual Ribbon host, host startup overhead and host shutdown. Do not infer that a custom message-pumped executable proves dllhost behavior. Registration remains blocked pending an explicitly agreed revised host test; overall task 1.2 is still incomplete.

Proposed revised host test: use corrected official-PIA binary only, verify class activation/lifetime outside OneNote first, establish clean single-host baseline, register only for an explicit manual load (no startup auto-load), confirm lifecycle logs and callback before any editing, disconnect and verify clean exit/unregister. If manual load also stalls, remove registration and stop rather than cycling startup tests. Actual activation route must be chosen and verified before registering; this is a plan, not permission or an executed test.

### Actual surrogate activation outside OneNote — PASS

Added Test-Activation.ps1 to temporarily register only this CLSID and AppID, with no ProgID/Office Addins registration. It uses the corrected offline binary and a separate official-PIA client to request CLSCTX_LOCAL_SERVER activation, obtain Ribbon XML, invoke lifecycle startup/shutdown and release objects. A finally block removes the temporary keys. Client wait is bounded; matching surrogate exit is checked for up to 15 seconds without killing unrelated processes.

First run's COM calls passed but PowerShell Process.ExitCode was null due to short-lived Start-Process handle acquisition, so the harness reported failure and cleaned registrations. Fixed by acquiring the process handle immediately. Final rerun returned client exit code 0, passed COM calls, and confirmed no matching dllhost remained. Temporary keys removed; no OneNote Addins entry was created. This completes independent activation/lifetime evidence, not actual OneNote loading or startup measurements. Next prerequisite is a user-controlled normal OneNote exit before preparing a manual-load-only host test.

### Controlled manual-load registration prepared

After the user confirmed normal OneNote exit, verified zero ONENOTE processes and absent probe registrations. Added Register-ManualProbe.ps1, which requires no running host, refuses existing keys and verifies official lifecycle interface identity in the corrected offline DLL. Registered that exact binary per-user with LoadBehavior=0, verified registry value 0. No OneNote process was launched by the script. Await user startup and explicit manual connection; no host loading/exit result is established. Register-Probe.ps1 remains blocked for installs and available for -Remove rollback.

### Actual OneNote manual connection and Ribbon callback — PASS

User manually connected and clicked Verify callback, reporting no visible result. This is expected for the log-only probe. Real host logs show OnConnection mode ext_cm_AfterStartup with nonnull application/AddIn objects at 10:34:47 UTC, GetCustomUI, and successful callbacks at 10:34:57 and 10:35:06 with nonnull control and nonempty control Id.

Process inspection confirms host ONENOTE PID 29312 and probe dllhost PID 21392 with matching AppID; log PID is 21392. Thus this tested plugin entry runs out-of-process, not in OneNote. Architecture must account for that fact before phase 2; an additional conversion worker may be redundant but remains a user-selected design until explicitly revisited. LoadBehavior remains 0. No document or clipboard edits occurred. Next verify user unchecking the Add-in produces OnDisconnection and normal host exit, then remove registration. Task 1.2 still requires unload/exit and startup overhead evidence.

### Manual disconnection passed; host process exit not yet passed

User reported closing OneNote after requested disconnection. Log records OnDisconnection(ext_dm_UserClosed) at 10:36:17 UTC in probe PID 21392. The matching dllhost is absent. Removed all probe registrations with Register-Probe.ps1 -Remove and confirmed Addins/CLSID/AppID keys absent.

Host ONENOTE PID 29312 remains with MainWindowHandle=0, Responding=True and CPU 8.046875 seconds unchanged across a 15-second observation. No force termination performed. Window closure is established; full host-process termination is NOT established. Do not mark task 1.2 complete or conclude normal host exit. Background behavior versus abnormal COM lifetime retention remains unresolved and needs a clean comparison/recovery decision; avoid starting additional host instances while investigating.

### Continued read-only preparation

After the user confirmed no popup and requested progress, host PID 29312 still had no main window; CPU had increased slightly to 8.125 seconds. Probe dllhost remained absent. Read-only registry inspection found EnableMemoryTrayMenuItem=1 and EnableMemoryShortcut=0; these settings alone do not prove why the host remains resident. No process was terminated and no preference changed.

Added SelectionProbe.cs and Test-Selection.ps1. The executable first checks for a visible existing OneNote window; without one it exits SKIP before COM activation. For a visible single-window fixture named Markdown插件验证, it reads piSelection XML in memory and prints only element counts and numeric/boolean selection markers, not note text, IDs or XML. Export length is checked after COM returns (not a bound on OneNote's export work); use the small disposable page only. Precise boundaries, mixed objects and views remain unverified.

Compiled successfully and executed: SKIP because no visible OneNote window. No COM activation, document or clipboard operation occurred in this run. Next needs user reopening the existing host normally onto the fixture and selecting controlled text. Host shutdown remains a separate open gate; selection preparation does not imply phase-one release.

### Inline controlled selection export — PASS for this sample

User selected **测试文字** including four asterisks in the controlled line 前文 **测试文字** 后文. Read-only piSelection export completed in 17 ms initially, then 6–7 ms in small fixture observations. Page has two Outlines, three OEs, five T nodes, no counted Image/InkDrawing/InsertedFile/Table elements. These are whole-page structural counts, not proof of selected-range purity for arbitrary pages.

Export marks Page/Outline/OEChildren/OE as selected=partial and selected T as selected=all. Selected T has eight characters, exactly matches the controlled token, has no serialized markup, and is index 1 among three sibling T nodes, with three serialized characters preceding it. Parent OE has objectID. Actual IDs and content were not printed or saved.

This establishes exact text extraction for one plain inline sample and offers object identity plus relative text position as candidate evidence. T nodes may be split according to current selection; their index alone is not a stable document position. Serialized HTML lengths are not generally visible-text offsets. Repeated-token positions, formatted text normalization, paragraphs, mixed objects and selection-change detection still need tests before defining a reliable fingerprint. No full task checkbox is warranted yet. No writes were performed.

### Repeated-token second position

User selected the second controlled token in **测试文字** 中间 **测试文字**. Export took 7 ms; selected T again exactly matches eight-character token, index 1 among two T siblings, with 12 preceding serialized characters and OE identity present. This differs from the earlier sample's three-character prefix but is not a same-paragraph first-vs-second comparison yet. Await first-token selection on this same line to demonstrate position discrimination with unchanged text content. No writes performed.

### Repeated-token first position — positional discrimination sample PASS

User selected the first token on the same controlled line. Export took 7 ms; selected text remains exactly eight characters, but selected T index is 0 and preceding serialized character count is 0 (second token was index 1/prefix 12). Both observations report OE identity present; actual object IDs are deliberately not logged, so cross-snapshot identity equality was not automatically asserted. This proves plain-text position signals distinguish the two user-designated samples; a complete identity/position revalidation implementation is still pending. Next validate full-paragraph selection markers and selection endpoint semantics. No document or clipboard writes performed.

### Full visible paragraph selected without next line

User selected the entire repeated-token line without selecting into the next line. Export took 7 ms; OE and T both have selected=all, selected T is 20 characters, single T child, prefix 0. This differs from partial-token OE selected=partial. It demonstrates a candidate full-paragraph marker for this plain fixture, not proof whether a paragraph separator is included: selected=all might describe all content independent of separator endpoint. Need endpoint comparison and multi-paragraph tests before deciding block replacement safety. No writes performed.

### Selection extended to next line start

User extended the same line selection to the next line start. Export took 8 ms; page OE count is now 5 (previously 4). Two OEs are marked selected=partial, each with T selected=all; first T length 20, second T length 0, both prefix 0. Thus selecting through a paragraph boundary differs structurally from selecting only visible line text. OE selected=partial alone cannot be used to reject an incomplete paragraph or prove an unsafe selection: it must be interpreted with child ranges and endpoint context. The extra OE may be a newly created blank line from fixture preparation, so the changed page count is not attributed solely to selection export. No automatic boundary classifier has been validated. No writes performed by the probe.

### Mixed text and image selection — rejection evidence PASS for sample

User selected a text-and-image fixture. Export took 7 ms and shows Image selected=all plus selected text in another OE, with Outline selected=all. This directly identifies a selected non-text object; a future validator must reject the whole operation. The selected text also contains serialized markup (353 serialized characters), so it cannot be treated as verified unformatted text. No text/markup or image contents were printed. Whole-page Image count alone was not used as a selection test. Other object categories, nested containers and stale/focus selections remain unverified. No writes performed.

### Title selection — rejection evidence PASS for sample

User selected characters in the page title. Export took 10 ms and marks Title selected=partial with selected OE/T descendants; no selected Outline marker appeared. Therefore Title ancestry distinguishes this sample from supported body selections and must cause rejection even though a T node is selected=all and its OE has object identity. T content has serialized markup; serialized length 108 is not the number of selected visible characters. No title text or XML was printed. Search-box focus/stale body selections still need independent focus validation.

### Search focus leaves stale body selection — critical constraint

User reported focusing search and entering probe after selecting body text. piSelection still exported body Outline/OE/T selection markers (selected T length 3, prefix 6). Thus XML selection state is not proof that paste will target the body. A current foreground/editor focus guard is mandatory; XML-only activation is unsafe.

Added read-only GetGUIThreadInfo/GetClassName metadata to the probe. Subsequent run reported no host focus HWND and no caret, while XML selection remained. Because the user returns to the chat to report readiness, and a console/tool run may affect foreground focus, these runs cannot establish the exact class of the search editor while actually focused. A focus capture initiated from within the host callback/hotkey, or a delayed capture while the user remains in OneNote, is needed. Do not use null focus as evidence that the search box is absent; it must cause fail-closed refusal. No writes or UI activation performed.

### Delayed search-focus sample

User switched to OneNote search; after an eight-second delay, host GUI focus class was RICHEDIT60W and a caret existed. The current page hierarchy name did not match the disposable fixture, so the probe stopped before GetPageContent. No arbitrary current-page content was read and no writes performed. This supplies a search-context class observation but does not establish that this class uniquely identifies search (a class allowlist alone is insufficient). Further comparison must retain the designated fixture and sample body/search focus separately. A different current page is a cancellation condition, not permission to relax the fixture guard.

### Delayed body-focus sample

User returned to the fixture, selected body characters and kept focus in OneNote. After eight seconds, focus class was OneNote::DocumentCanvas with no Win32 caret HWND. Fixture check passed; piSelection took 7 ms. Selected body T contains serialized markup (154 serialized characters; preceding serialized length 155), so offsets must be derived from a parsed/validated text representation, not raw string lengths. Contrast with search RICHEDIT60W supports using focus plus window/page/selection evidence. A caret requirement would wrongly reject this legitimate body selection. Title may share the canvas class, so Title ancestry must still be checked. No broad compatibility claim or completed task from these two samples.

### Multiple windows — rejection PASS

User opened a second OneNote window via View/New Window. Test-Selection.ps1 returned SKIP: requires one window. The guard executes before accessing current page hierarchy, focus metadata or page body. This establishes the probe's multiple-window early refusal in this environment. Remaining 1.3 coverage includes non-normal views, cross-container/cell selections, formatted-text policy, parsed boundaries and reusable evidence comparison; do not mark the whole task complete from this case.

### Cross-container selection — rejection PASS for sample

After closing the extra window, user selected text across two independent containers. Export shows two distinct Outline nodes marked partial, each containing selected text. Added a structural screening result to the probe: selected Outline count 2 yields refusal for not exactly one selected body container; rerun confirmed this result. Also screens selected Title and explicit selected Image/InkDrawing/InsertedFile/Table nodes, but is not a production validator or comprehensive object detector. Focus was unavailable while user was in chat and is not treated as approval. No document/clipboard writes. Same-container nested cells and unsupported embedded objects still need coverage.

### Cross-cell selection — rejection PASS for sample

User selected across two table cells. Export took 9 ms: exactly one selected Outline, Table selected=partial, two distinct Cells selected=all, with selected text in each. Initial conservative object guard rejected the selected Table; added explicit selected Cell count and cross-cell refusal ahead of generic object refusal. Rerun reports selected Cell count 2 and Structural refusal: cross-cell selection. Row/cell layout did not need to be inferred from the user's requested table dimensions. No writes. Single-cell policy is not established by this result; current probe conservatively refuses any selected Table and must not be treated as a confirmed requirement to prohibit single-cell formatting.

### Full-page view and wrapped text sample

User selected the controlled token and pressed F11. Added explicit Window.FullPageView/SideNote output; confirmed FullPageView=True, SideNote=False, focus class OneNote::DocumentCanvas, one selected Outline and zero selected Cells. Small fixture export took 8–9 ms. No selected Image/Table markers despite those objects existing elsewhere on the page, illustrating why whole-page object counts must not reject unrelated selections.

Selected T contains 149 serialized characters. A diagnostic XML-compatible fragment parse (XmlResolver=null; only nbsp normalization) succeeds and extracts exactly the controlled eight-character token from three nested elements. The serialized length mismatch is style/wrapper markup in this sample, not loss of selected text. This parser is diagnostic only, not a general HTML parser/sanitizer or production plain-text classifier. Mere presence of markup must not be equated with user-applied rich formatting; default OneNote wrappers require a deliberate style policy. Full-page view read-only selection access is demonstrated, but actual paste/focus stability remains unverified. No writes performed.

### No selected characters — refusal PASS for sample

User returned to normal view and placed an insertion point without text selection. FullPageView=False. XML still contains Outline/OE selected=partial and an empty T selected=all, so presence of selection markers alone is insufficient. Diagnostic decoded selected text length is zero; added explicit no-selected-text refusal and confirmed on rerun (10–11 ms exports). This conservatively rejects both this insertion-point sample and separator-only selections; paragraph separator semantics are not inferred from zero text length. No writes performed.

### Reusable structural screening and evidence experiment

Added SelectionScreen.cs with bounded DTD-prohibited XML parsing and diagnostic refusal categories for title, container count, cross-cell, unknown selected objects/states, table-not-validated, unparsed/unsupported text markup, missing OE identity and empty text. Only span/br markup is recognized for this experiment; styles and block boundaries are deliberately not approved. Unknown syntax or parsing failure is refused. It is not a product HTML parser or permission to paste.

Added SelectionEvidence.cs: holds page/window/focus identity and SHA-256 of the entire selection-marked page XML in memory only. Requires single window, foreground root matching host and DocumentCanvas focus. Comparing full exports avoids raw HTML offset arithmetic and detects repeated-text relocation, but may over-cancel on metadata changes, incurs page export costs and does not prove atomicity or detect an edit-away-and-back transient. Focus can still change after capture; not a completed commit protocol.

26 offline cases passed, including unchanged evidence, page/content changes, repeated identical text at different positions, unknown foreground, missing/search focus, multiple windows, mixed objects, malformed/oversized XML and DTD rejection. Re-ran full offline COM suite successfully. Integrated screening/evidence into the read-only fixture probe; current insertion-point run correctly refused empty-text and refused evidence capture. A real focused nonempty double-read still needs user setup.

### Read-only clipboard inventory

Added ClipboardInventory.cs/Test-ClipboardInventory.ps1. Enumerates format IDs under a short OpenClipboard lock, never calls GetClipboardData and does not render delayed formats, modify content or take a snapshot. Known names are allowlisted in output; other registered names omitted. Current clipboard contains 13 enumerated formats including HTML, Unicode/ANSI/OEM text, locale, DIB/DIBV5, bitmap/metafile handles and unclassified private formats. Sequence was unchanged during inventory. This demonstrates why a text-only backup would not satisfy the user's preservation requirement; none of these formats has been validated for deep snapshot/restoration yet. Current content would be rejected by a strict unclassified-format policy. Resource medium handling and complete-format preservation remain task 1.5 work.

Next user operation: keep a nonempty plain body selection focused for delayed double-read. After this, clipboard preservation/paste completion require a separate bounded experiment with an agreed disposable clipboard fixture; no paste or clipboard writes have yet been authorized/executed by the probes.

### Focused double-read evidence — PASS for controlled sample

After user kept the controlled token selected, focus was DocumentCanvas and structural screening returned a candidate. Initial evidence capture refused because foreground top-level HWND was compared directly against COM WindowHandle. Diagnostics established foreground PID matches host, and GetAncestor(COM WindowHandle, GA_ROOT) matches the foreground root although raw COM handle does not. Corrected both captures to compare normalized top-level roots, not relaxed to PID-only matching.

Final run: foreground root equals COM window root, nonzero focus, fixture title matches, decoded text equals controlled token, and conservative full-export double-read matches=True. This validates stable evidence for this sample; no action is committed, no cross-snapshot atomicity is claimed. Current path still requires style policy, contiguous range validation and block endpoint handling. Unknown focus fails closed.

### Fixed HTML transport fixtures — offline PASS

Added HtmlPayloadTests.cs with a fixed validation payload builder and four CF_HTML roundtrips: Chinese/emoji, title/list/task symbol, code CRLF/tab/empty lines, quote and combining characters. Checks UTF-8 byte offsets, strict decoding, context and exact fragment recovery. All pass together with 26 selection/evidence cases and COM tests. This is fixture generation, not the product Markdown renderer or proof of OneNote import/undo.

Next blocked experiment needs a user-prepared disposable clipboard fixture because the existing clipboard contains unclassified/GDI/private formats that are not validated for reliable deep backup. Do not replace that clipboard or pretend text-only preservation is complete. Ask user to copy a nonsensitive short plain string from Notepad before a bounded snapshot/restore experiment. No such writes have been performed yet.

### Disposable clipboard authorization; strict-format refusal

User authorized the disposable clipboard test and reported copying the fixture. Inventory shows Unicode/ANSI/OEM text/locale PLUS three registered Chromium source/custom-MIME formats. Temporarily printed registered format names (not payloads) to identify these; then restored allowlisted-name output including the observed Chromium names. This suggests a browser/electron copy path or manager, not proof of which program wrote them.

Added ClipboardRoundtrip.cs with an exact fixture-text check and a strict supported-format gate (CF_UNICODETEXT/CF_TEXT/CF_OEMTEXT/CF_LOCALE only), independent HGLOBAL allocation before writes, owner window, sequence check under lock before restore and byte comparison. It is disposable-fixture code, not a production snapshot service: no delayed-source timeout isolation, GDI/private support, recovery persistence or concurrent-race coverage is established. Failure after clipboard modification can require manual fixture recopy; no real OneNote work uses this code.

Compiled and ran: SKIP non-text/private format present; clipboard unchanged. This is correct refusal under agreed preservation rules; no replacement/restore pass claimed. To test the intended plain-text path, user must copy directly from Notepad while the tool waits, without recopying text from chat. If private formats persist, identify the source instead of silently stripping them.

### Actual Notepad copy also carries OLE/enterprise metadata

After a 15-second delayed run during user copy, strict roundtrip again refused without writing. Inventory confirmed clipboard owner process Notepad, with CF_TEXT/UNICODETEXT/OEMTEXT/LOCALE plus DataObject, Ole Private Data and EnterpriseDataProtectionId. Thus the second refusal is not user failure to copy from Notepad: even this system's Notepad path has unsupported metadata. Registered names were inspected (no content) then added to allowlisted inventory labels; values were not read.

Do not strip EnterpriseDataProtectionId/OLE metadata or claim full clipboard preservation. Need an explicit decision for either a synthetic disposable test clipboard created by the harness (does not validate preservation of Notepad's metadata), or more comprehensive OLE snapshot investigation. Production requirement remains unchanged: unsupported original formats cause skip. No clipboard write has occurred in these runs.

### Authorized synthetic clipboard roundtrip — PASS with discovered close semantics

User explicitly allowed replacing the disposable copy with a program-generated plain-text fixture. Added -SyntheticDisposable. First run published the fixture, then conservatively refused restoration because sequence changed; diagnostic inventory found only standard text formats. Sequence had been recorded inside the publication lock, before CloseClipboard could synthesize associated text formats/update sequence. Changed the fixture to capture sequence after closing publication, then verify BOTH sequence and owner HWND while holding the restoration lock. This is not a universal proof against external interference and remains experimental.

Final normal path passed independent byte-copy/restore checks for every original synthetic format. Added same-process different-owner-window simulated copy: sequence/owner change detected and newer text retained without restoring old snapshot. This is a controlled ownership race sample, not external-process stress testing. A subsequent invocation briefly found the clipboard busy and skipped; retry after two seconds passed. Final clipboard is the restored synthetic disposable fixture, not the discarded original Notepad OLE metadata. No OneNote interaction or paste occurred.

Remaining task 1.5 gaps: actual Notepad/Office/browser private/OLE metadata, GDI handles, delayed-source timeout isolation, restore failure recovery and true competing-process tests. Task 1.6 paste completion is still unverified. No full task marked complete.

### First authorized actual inline paste — submitted, human verification pending

User authorized a one-shot controlled inline paste, selected **测试文字** in the disposable fixture and retained focus. Current clipboard inventory was standard synthetic text formats. Added PasteFixture.cs/Start-PasteFixture.ps1: exact fixture/page/text guards, single-window and foreground-root/DocumentCanvas evidence, full-export revalidation before and after publication, modifier check, strict synthetic clipboard snapshot, fixed bold HTML payload and a single four-event Ctrl+V input batch. It never retries or sends undo. Clipboard restoration requires explicit human acknowledgement through a named event; this is a test mechanism, not automated paste completion or a production feature.

First run refused on XML parsing before any clipboard write/input. OneNote emitted an unquoted attribute in its internal text wrapper. Added a bounded diagnostic tag-lexeme adapter for unquoted attributes, preserving quoted values/text, then validated 29 structural/evidence cases plus four HTML and COM suites. This is not a full HTML parser/style policy. Exact controlled text check remains mandatory.

Second run PID 29648 passed three selection captures, published payload and SendInput reported 4 at 11:45:26 UTC. This means input submitted, NOT proven complete or correct. Await user report of bold/no-star result and immediate single undo. Helper waits up to five minutes for human acknowledgement before restoring clipboard if ownership/sequence unchanged; outcome uncertainty does not trigger automatic undo or retries. Timeout exits without claiming restoration. Snapshot is in memory and disposable only; failure recovery persistence remains unresolved. No full task marked complete.

### User reports no change; keyboard input discriminator defect found

User explicitly reported no visible replacement. Signaled only the existing cleanup acknowledgement event, without sending any further input. Helper exited and logged NEWER_CLIPBOARD preserved; snapshot not restored. It could not verify whether a third-party update or Windows lazy text synthesis caused the sequence change; no overwrite was attempted. Thus the synthetic original was not confirmed restored in this run. Do not ask the user to undo a nonexistent change.

Code inspection found INPUT.Type was left at zero (INPUT_MOUSE) for all four events, rather than INPUT_KEYBOARD=1. SendInput returning 4 therefore counted submitted mouse-type records, not Ctrl+V. This is a definite fixture defect explaining why the earlier result must not be called a successful keyboard paste; it does not independently establish every possible side effect. Corrected the input factory and added a no-injection --check-input-layout mode: verifies x64 size 40, union offset 8, keyboard type/key/up-down flags. Start-PasteFixture.ps1 -CheckOnly compiled and passed. No corrected paste was automatically retried. Requires fresh user selection and clipboard fixture validation before a new one-shot attempt.

### Corrected retry refused before publication

After user reselected, corrected test PID 6432 passed input-layout checks and initial selection captures but ClipboardRoundtrip.Read refused unsupported formats before payload publication or SendInput. Inventory showed owner OpenCode with HTML, Chromium RFH token/source URL, plus standard text. This is not the approved synthetic clipboard fixture and must not be overwritten without a renewed disposable-content decision. No paste occurred, no undo needed; helper exited. Next requires permission to reset the current clipboard to a synthetic fixture immediately before the bounded paste experiment.

User explicitly authorized reset. Added -ResetDisposable to the one-shot fixture, performed only after initial selection/focus guards, before snapshot creation. Run PID 20632 stopped at initial guard after 12 seconds; neither clipboard reset nor paste occurred. Exact cause (selection/foreground/fixture) is not distinguished by this log. Requires user to retain a fresh controlled selection and actual OneNote foreground; no automatic retry was performed.

### Corrected keyboard paste submitted after synthetic snapshot fixes

User selected again and requested continued work. PID 23676 reset the authorized fixture but refused at clipboard sequence check before input. Next run PID 30548 recorded sequence after CloseClipboard and checked actual snapshot byte equality; refused because snapshot had been taken before Windows synthesized standard text formats on close. Neither attempt sent paste input.

Separated fixture reset/publication and snapshot capture into distinct clipboard lock intervals, so original format set is captured after initial CloseClipboard synthesis. Sequence captured after snapshot close; publication rechecks sequence and all snapshot bytes under lock. PID 27912 then passed initial/final evidence and clipboard checks, and submitted four correctly typed keyboard events at 13:03:58 UTC. Awaiting user visual result and immediate one-undo check; helper waits for explicit human acknowledgement before restore. Do not label this paste successful based on input count. No automatic retries after this submission.

### User-confirmed single undo; test process cleaned up

User reports one Ctrl+Z restored the original. This is actual generated-CF_HTML/input-path undo evidence for this single inline fixture, distinct from the earlier manual-source tests. User has not explicitly confirmed the bold visual style or unchanged surrounding content separately; do not infer all visual/structural checks passed from the undo statement alone.

Signaled the existing acknowledgement event only; no new input sent. Helper exited. Cleanup log says NEWER_CLIPBOARD preserved; snapshot not restored. Read-only inventory confirms clipboard owner OpenCode with HTML and Chromium metadata, so retaining newer content was appropriate; this does NOT validate normal paste-followed-by-automatic-restoration. Test Addins registration remains absent. Task 1.4 remains partial (other constructs, surroundings, read-only settings and full plugin path not validated), and task 1.6 remains open. No automated paste-completion evidence was obtained from a human-confirmed undo.

### Preparation for observed postcondition test

Added InlinePostcondition.cs and five offline tests: locate original selected OE identity, derive expected controlled inline replacement, require same target prefix/suffix and exact expected decoded text, compare outside XML after removing selection and known author/modification metadata. Refuses unchanged text, outside edits, suffix loss and changed target identity. Passing means fixture text-state observation only; does not prove bold style, future clipboard requests completed, transaction semantics or concurrency attribution. Exact outside serialization comparison may over-refuse harmless metadata/layout updates.

PasteFixture now optionally observes this postcondition for up to 2.5 seconds of bounded polling after submitted inputs while enforcing page/window identity. It still waits for human acknowledgement before restoring clipboard, never uses a fixed delay as restoration permission, never retries/undoes. COM call duration itself is not bounded by the polling clock, an outstanding production isolation concern. Full offline suite passes (29 screening/evidence, 4 HTML, 5 postcondition plus COM); --CheckOnly compile/layout passes. No new paste was executed. Next requires a fresh user-controlled inline selection and a separate visual bold/surroundings report to test this observer.

### Observer run failed to confirm; undo evidence repeated, no further manual repeat

User authorized the next sample. PID 30768 submitted four keyboard events; the observer did not confirm its strict full outside-XML postcondition in 2.5 seconds. This means the observer failed to establish the predicate, not that paste failed or that outside content was corrupted. User reports full restoration after one undo again. Signaled cleanup acknowledgement and confirmed fixture exited. User questioned repetitive testing: treat the inline one-undo behavior as established for this fixture and do not repeat that manual test without a distinct unresolved change. The new hypothesis was automatic postcondition observation/clipboard completion, which remains unvalidated; stop requesting identical undo demonstrations. Investigate observer mismatch offline or return an explicit design decision for task 1.6 rather than claiming success or burning repeated user operations.

### Offline observer diagnosis refinement

Refactored InlinePostcondition into categorized diagnostics without weakening success conditions. Twelve offline cases now separate unchanged target, missing identity, unexpected target content, matching target with differing outside serialization and true fixture match. Tests also explicitly show that target text matching can pass even when bold style is absent; no visual correctness or clipboard-consumption claim is attached. All offline suites pass, and the paste binary compiles/checks layout without sending input.

Prior real run stored only a single boolean outcome and did not persist page XML by design. Therefore its exact mismatch cannot be reconstructed offline. No new paste was run. A distinct diagnostic experiment would need a fresh authorized replacement with categorized observations; it does not require another manual undo demonstration. If user declines further paste experiments, task 1.6 remains blocked and design must return for explicit interaction/clipboard trade-off decisions. Do not invent the missing mismatch reason or automatically proceed to full implementation.

### Authorized categorized diagnosis result

User authorized the diagnostic run. PID 14096 submitted the single keyboard paste at 13:16:32 UTC. Within approximately 56 ms the observer reported expected-target-text-but-outside-serialization-differs, persisting through the observation window. Thus target identity, expected replacement text and target paragraph prefix/suffix matched; failure is in the exact XML comparison outside the replaced OE. It does not show whether the mismatch is benign layout/export metadata or actual outside content/style changes, and does not prove complete clipboard consumption or bold style.

Signaled the test event to end diagnostic waiting and attempt cleanup; this signal was agent-controlled diagnostic completion, not a user visual acknowledgement, despite historical helper log wording. Helper exited, no automatic undo/retry. Need a direct visual outside-content check and finer non-content diff categories before allowing a semantic comparison to replace strict serialization equality. Never broadly ignore outside XML or call this a proven no-damage result.

User visually confirmed the target text is bold and there are no visible changes elsewhere. Together with target/prefix/suffix observation, this is a passing visual/inline-replacement sample. Exact outside serialization equality still fails and no automatic clipboard-completion guarantee follows. Do not request the same bold/outside/undo checks again for this unchanged inline fixture.

Added OutsideDiff.cs category-only diagnostics (layout, identity, style, content, child shape, other attributes) and five additional offline cases. Seventeen inline diagnostic cases now pass along with all earlier suites; production acceptance remains unchanged. No new host paste performed. Need architecture decision based on actual observed host behavior: the corrected COM Add-in is already hosted in dllhost, not inside ONENOTE. The planned extra conversion worker would create another external process; confirm whether to retain separate conversion isolation or combine conversion in the surrogate before creating product projects. Neither choice resolves clipboard completion or authorizes phase-two implementation ahead of open gates.

### User decision: merge conversion into the COM surrogate

User explicitly accepted extreme conversion faults making the plugin unavailable and possibly requiring OneNote restart, and chose the merged design. Updated proposal/design/specs/tasks: no dedicated conversion worker or product IPC; conversion runs asynchronously within the registered dllhost. Persistence is feature enablement/fault state, not a promise to kill the Windows-managed surrogate when disabled. Timeout invalidates results but cannot forcibly stop a stuck managed parser; do not restart work while a previous task is still running. Abrupt process failure may prevent cleanup/persistence, so an unfinished-operation marker must lead to manual recovery on next initialization. No automatic proxy restart or Environment.Exit/kill-by-name allowed.

This is an approved architecture update, not gate completion. Host background lifetime, automatic paste-completion/clipboard restoration, private format preservation, block syntax and shortcut coverage remain open. No new installation or document/clipboard mutation occurred for this decision.

### Export-mode structural difference found without further paste

Added ExportComparison.cs and ran it read-only on the designated fixture: piBasic vs piSelection yields child-count and selection-metadata differences, while two piSelection exports yield none. This independently establishes that selection export mode changes node structure without editing. It is a plausible contributor to the earlier outside comparison failure (selection collapse/segmentation), not proof of that historical diff's exact cause because those exports were not retained.

Changed the experimental observer to use a separate piBasic baseline and piBasic post-export for outside comparison, while piSelection remains only for locating the original target/range. Verifies baseline target decoded text equals original selection export target text and revalidates selection after baseline acquisition. Does not broadly normalize away child-count/content/style changes. Added four cases: selection-only T splitting vs same-mode baseline, baseline target edit rejection and outside edit preservation. All 21 postcondition cases, 29 selection cases, 4 CF_HTML cases and COM suites pass; paste fixture layout/check-only compilation passes. No new paste or clipboard modification occurred.

Actual observer and clipboard completion remain unverified. A new run, if authorized, tests comparable export modes and categorized result only; no need to repeat manual undo confirmation. Clipboard restoration stays behind explicit acknowledgement until the consumption/completion strategy is separately accepted and validated. First-phase gate remains closed.

### Comparable-mode observer experiment still blocked

Authorized run PID 14188 submitted one paste. Observer reported baseline-target-text-changed (selected target text and piBasic pre-state failed the observer's equivalence check), so no completion was established. Sent cleanup event only and verified helper ended. No automatic undo/retry. This category concerns the observer's baseline/export interpretation, not proof that the user edited during capture or the paste damaged notes.

Added a mandatory pre-paste BaselineMatches guard: if selected target and piBasic baseline do not have identical decoded text, skip before input instead of discovering an invalid baseline after a write. Three additional offline cases pass, bringing postcondition tests to 24; all earlier suites and check-only build pass. No further actual paste executed.

Task 1.6 remains a design gate: the current export-based observer is not yet reliable enough to authorize automatic clipboard restoration. Stop repeated visual/undo tests. Need user decision whether to continue a dedicated read-only/export investigation or revise the clipboard promise to leave the generated rich text on the clipboard after formatting (with explicit user acceptance and spec updates). Successful inline formatting/one undo remains established; neither option erases that evidence. Formal implementation remains blocked until the decision and applicable validation are completed.

### Approved empirical restoration policy

After source comparison with OneMore and explanation of timing limitations, user explicitly chose default 200 ms, configurable 500 ms, retaining ownership protection. Updated current proposal/design/specs/tasks to replace mandatory automatic completion proof with this accepted empirical compatibility trade-off. Old observer history above remains evidence, not current restoration requirements. Full snapshot preservation for supported formats, refusal for unsupported formats and no overwriting newer copies remain unchanged.

Added RestorePolicy.cs with ten offline delay/default/ownership cases. Test fixture now defaults to --delay=200 (optional 500), waits without holding clipboard lock, then checks owner+sequence in restoration lock. It does not need XML postcondition approval or human acknowledgement for delay mode. The diagnostic observer code remains unused by default. Test fixture uses Thread.Sleep on its independent hidden test thread, not in a product/UI callback; actual product must use asynchronous delay. Historical log label PASTE_SUBMITTED waiting for human acknowledgement is stale in delay mode and must not be interpreted as a live wait.

All offline suites pass, fixture check-only compiles, OpenSpec strict validation passes. No new real paste executed for the new policy; task 1.6 remains pending real 200/500 ms restore, collision and failure tests. User setup required for one controlled timing/restore sample, not another proof of inline undo. No task prematurely marked complete.

### Default 200 ms host experiment — restoration confirmed

User requested the 200 ms sample. PID 26272 passed fixture/evidence guards, reset the authorized disposable clipboard, submitted four keyboard events at 14:10:12.499 UTC, and logged RESTORED at 14:10:12.705 UTC (approximately 207 ms later). Independent read-only Unicode clipboard check equals clipboard-probe-123; test process has exited. Thus restoration occurred and the disposable original text is verified restored for this sample. User visual confirmation of the resulting OneNote formatting is still pending; no new undo demonstration requested. No statement of timing guarantee, arbitrary-format support or full task completion follows. Next work can test 500 ms/collision cases independently; actual product integration remains gated.

User subsequently confirmed bold text, removed markers and unchanged surroundings. Default 200 ms inline paste/restoration is therefore a passing integrated disposable sample; no further same-sample manual repetition needed.

Added and ran a 500 ms standalone clipboard roundtrip against the still-exact disposable string without resetting content: all snapshot format bytes restored. This is not an additional OneNote 500 ms compatibility test. Remaining strict requirement gap is common OLE/private original formats: actual Notepad/Office/browser copies currently cause refusal. A practical product cannot silently discard those formats or simply call that full preservation. Need explicit scope decision or implement/test OLE-aware snapshot support before declaring gate complete. No full task checkbox changed.

### Approved common-content preservation scope

User accepted preserving common content (text, HTML/RTF, images) rather than every app-private paste capability. Updated proposal/design/spec to permit individually justified auxiliary metadata handling, not blanket private-format stripping. Unknown formats and unvalidated protection/privacy semantics still refuse. Added clipboard-support.md separating requested support from observed implementation; actual Notepad OLE/protection support is not claimed.

Added twelve offline classification cases including text, rich text, images, OLE transport, protection flags, files and unknown Chromium metadata. Full offline suite and strict OpenSpec validation pass. Classification does not read/drop payloads or authorize publishing. No clipboard or document write for this update. Stage-one task 1.5 remains partial; selection/paste regressions need not be repeated to implement richer snapshots.

### Authorized synthetic common-content and competing-process tests — PASS

User authorized replacing current clipboard with generated text/rich-text/image fixtures and leaving the final test image. Implemented CommonClipboardTests.cs and ran Test-CommonClipboard.ps1. Unicode text, simultaneous HTML/RTF/plain representations, PNG stream and 2x2 DIB passed independent byte snapshots → temporary text → 200 ms restoration, with owner/sequence check under lock. Restored DIB decoded to expected dimensions and pixel colors. Synthesized CF_BITMAP is omitted only when an independent DIB representation exists; this is a controlled example of representation substitution, not a blanket private-format exemption.

Added child-process competing-copy case; verified guard refuses old restore and actual newer text remains. Final run passed all cases; final clipboard is the authorized small image. No OneNote access. Helper bounds individual HGLOBAL size and enumerated count; robust total-size budget, external delayed-render timeout, real application-origin payloads, OLE/enterprise classification, restore failure persistence, and image edge cases still pending. Therefore tasks 1.5/1.6 are partial, not fully complete.

### Real OneNote clipboard sample — read-only independent snapshot PASS

User copied real bold content. Win32 inventory owner is ONENOTE with 13 formats. Added OleInventory.cs and a standalone process with a five-second supervisor timeout. Initial OLE enumeration reveals six source representations: OneNote 2016 Internal (HGLOBAL), HTML (HGLOBAL), Unicode/ANSI text (HGLOBAL), enhanced metafile (ENHMF), and DIB (HGLOBAL). DataObject/Ole Private Data wrappers and some synthesized formats in Win32 inventory are not enumerated as original application representations by OLE. This supports using OLE enumeration, not blindly copying transport bookkeeping bytes.

Extended read-only probe to materialize only approved HGLOBAL common representations, release STGMEDIUM correctly and discard/clear copied arrays without printing content. Actual copies succeeded: HTML 695 bytes, Unicode 8 bytes, ANSI 7 bytes, DIB 16936 bytes. Clipboard sequence unchanged; no data published or restored. Enhanced metafile and OneNote Internal were not requested. This validates independent capture of common fallback representations only, not preservation of native OneNote private behavior/vector fidelity. Production policy still needs clear conditional omission rules. Original clipboard remains intact.

Build first exposed IEnumFORMATETC.Next third parameter type mismatch, corrected to int[]. Subsequent metadata and snapshot runs passed. No OneNote page API/edit used, no plugin registered. Next real roundtrip requires permission to replace/restore this sample under common-content scope; previous user instruction for this sample authorized read-only inspection only.

### Real OneNote clipboard roundtrip — PASS

User authorized restore testing and requested no repetitive authorization within the agreed scope. Extended the OLE fixture with -Roundtrip. It enumerates and materializes common HGLOBAL representations, caps each at 4 MiB and aggregate at 8 MiB, releases source refs/media before mutation, preallocates restoration storage, checks sequence before publishing and owner/sequence under lock before restore. Only OneNote 2016 Internal with rich+Unicode fallback, and GDI/vector extras with DIB fallback, are conditionally omitted; unknown formats or unsupported content media refuse. These omissions are specific approved common-content fidelity trade-offs, not blanket dropping private/enterprise flags.

Actual sample HTML 695 bytes, Unicode 8, ANSI 7, DIB 16936 all restored byte-for-byte after temporary test text and 200 ms wait. Original OneNote-private/vector representations intentionally not restored. Helper exited normally; no page edits or plugin loading. First compile found ambiguous IDataObject namespaces, fixed with official ComTypes alias before executing. Further media/aspect guard added to reject any supported format not offered as ordinary HGLOBAL content rather than silently omitting it. Later build-only verification does not repeat mutation.

### Actual Notepad OLE source investigation

User provided the requested Notepad copy. Win32 inventory includes DataObject/Ole Private Data/EnterpriseDataProtectionId, but OLE source enumeration offers only CF_TEXT, CF_UNICODETEXT and EnterpriseDataProtectionId. Text independent copies succeeded (80/40 allocated bytes, contents not printed). Read-only materialization of the enterprise marker returned no usable allocated bytes; report this as advertised but unmaterialized, not permission to drop it. First probe rejected zero size, then diagnostic output was refined to report it without changing clipboard. Clipboard sequence unchanged, no roundtrip attempted. Broad Notepad source support remains unapproved pending semantic handling; neither user mistake nor blanket metadata stripping is an acceptable explanation.

### Observational shortcut fixture prepared

Added 90-second HotkeyProbe.cs: WH_KEYBOARD_LL observation only, no suppression/injection, never edits documents/clipboard. Logs only Ctrl+Shift+D target/non-target category, no other keys or text. Captures GUI class in hook, defers process lookup/logging to message loop, debounces D key repeat, ignores injected flag; always CallNextHookEx. Compiled and started PID 30112 for user-directed body and Notepad comparison. This is preparation for task 1.7, not an enabled intercept implementation or full selection eligibility test. Requires user keypresses and automatically unhooks/exits.

User pressed in both apps. Log confirms body-focus-candidate at 14:36:57 and non-target at 14:37:01, both passed through, followed by STOP at 14:37:12. This validates observation/classification for two samples and clean timer stop; it does not validate enabled interception, disabled hook release, repeat/modifier handling under actual conversion or title/search discrimination. Task 1.7 stays partial.

### Prepared distinct block conversion fixture

To avoid repeating already-proven inline undo, added a fixed four-paragraph fixture (heading, bullet, quote, inline code), whole-paragraph/contiguity/text guards and fixed HTML payload. Six synthetic boundary tests pass along with all offline suites; check-only compilation sends no input. Test is deliberately not a Markdown parser. It allows one explicit trailing empty selected paragraph for next-line-start selection, rejects visible unselected prefixes/suffixes, multiple parents and gaps. Actual multi-paragraph paste/render/undo and surrounding structure need a user-prepared fixture. No block paste has yet been executed. No full task marked complete.

### Four-paragraph actual paste submitted; restoration logged

User twice reported fixture ready; executed only one paste attempt (PID 6444), not one per message. Exact block/whole-paragraph guards and repeated evidence checks passed. Four keyboard events submitted at 14:42:16.592 UTC; original disposable clipboard formats restored at 14:42:16.807 UTC (about 215 ms), helper exited. Await distinct block visual/structure and single-undo confirmation. This is first controlled generated block sample, not a repeat of inline undo. No claim that HTML task/code-block/other syntax is covered by this four-line fixture.

User reports heading/list/quote formatted but code appears unformatted; screenshot supplied. Performed read-only fixture export, without edits or clipboard access: exact code text appears once, no explicit monospace in its T fragment, but enclosing OE style explicitly declares Consolas. This confirms OneNote moved the font to paragraph-level style, not that it discarded the monospace font. No background/border was requested in this fixed payload or the current minimal code requirement. User visual distinction may require presentation clarification; do not dismiss the feedback or silently add styling. No additional paste/retry. Multi-paragraph single-undo and paragraph restoration remain unreported.

User confirmed single undo restores all four source lines and paragraph structure; this distinct block sample's undo is established and must not be repeated unchanged. User clarified inline code requires monospace plus light-gray background. Updated spec/design accordingly; no claim about fenced-block background. Prepared a code-only fixed span payload (#eeeeee) and exact `code` token mode to test shade import and same-paragraph isolation without rerunning heading/list/quote tests. Actual shade preservation not yet verified.

Code-only attempt PID 4304 refused initial guard before resetting clipboard or sending input. Read-only follow-up shows full OE selected=all and a 12-character selected T, not the required six-character backtick/code/backtick token. Focus and same-state double-read pass. Do not expand replacement to the whole line just to make the test run; request the narrow token selection and retain guards. No writes in this attempt.

User selected six-character code token. Single attempt PID 27620 passed exact token and repeated context guards, submitted keyboard paste at 14:54:26.125 UTC and restored disposable clipboard formats at 14:54:26.344 UTC (about 219 ms). Process exited. Fixed payload requests Consolas plus #eeeeee background on the inline span. Await visual shade/locality result; no automatic retry/undo and no claim of CSS preservation solely from payload submission.

User supplied screenshot and confirmed background is visible and requested keeping the present format. Treat the exact tested inline style (Consolas + #eeeeee, no extra border/padding requirement) as accepted. Surrounding visible words remain outside the shaded inline region in the supplied result. Do not repeat this visual fixture unchanged. This does not establish fenced-code-block whitespace preservation or broader CSS import compatibility.

### Fenced-code whitespace/Unicode fixture prepared

Added a fixed fenced sample containing four leading spaces, an empty line, Chinese and emoji. Payload is pre with Consolas/white-space:pre, preserving exact internal text; no promise of block background. Exact source, complete contiguous paragraph boundaries and unexpected prefix/suffix checks remain in place. Nine block/whitespace cases pass, including missing indentation/emoji rejection; all offline suites and input-layout compile pass. Actual OneNote pre whitespace import remains untested. Prepared fixture only, no new paste executed. Next user operation is this distinct code-block sample, not repeated inline rendering/undo.

Fenced fixture actual attempt PID 15960 passed exact source/paragraph checks and submitted once at 00:19:21 UTC; 200 ms recovery logged approximately 206 ms later, helper exited. Read-only post-export: alpha paragraph has four-space prefix and nine decoded characters; Chinese+emoji paragraph has four decoded UTF-16 characters and no preceding U+0020/NBSP spaces after normalization. No matching outline reconstructs exact expected alpha + empty line + indented Unicode sequence. This is failure to establish required whitespace preservation for plain pre import, not a Unicode loss claim (Chinese/emoji are present). Paragraph-level indentation may contribute visually but is not equivalent to preserved literal spaces. Need user visual result and one undo only to restore this distinct code-block sample; do not claim code-block support or silently accept whitespace loss. No automatic retry/undo.

User screenshot confirms second line visually lacks indentation; user reports one undo restored the original. Plain pre import fails the whitespace requirement; undo is established for this sample.

Prepared an alternative fixed fixture using explicit br and NBSP entities inside a Consolas span. Thirteen offline block/whitespace cases pass, including visual-space-normalized recovery, entity escaping and explicit rejection of tabs pending a separate policy. This candidate intentionally changes U+0020 spaces to U+00A0, so it is NOT exact text preservation and cannot be silently used under current code-whitespace spec. No actual NBSP paste executed. Need user decision on visual-code rendering versus strict copy-back whitespace before pursuing this candidate; a strict requirement means continue investigating another transport or reject affected code blocks, not label NBSP exact.

User chose strict original whitespace. Added refusal scenario to spec and disabled NBSP paste switches; code-block subset remains unapproved until exact transport passes. No semantic relaxation to visual equivalence.

Prepared alternative RTF code transport: literal U+0020 spaces, explicit line/tab controls, escaped braces/backslashes and signed UTF-16 Unicode controls including surrogate pairs. No HTML is simultaneously offered in RTF fixture, avoiding selection of the known failing HTML pre path. Nine Windows RichEdit text roundtrips pass (indentation, blank lines, tabs, trailing spaces, Unicode/emoji, escaping, combining marks, empty text). First tests exposed loss of terminal line breaks without an RTF terminal paragraph; adding the explicit document terminal paragraph fixed the fixtures. This is RichEdit evidence only, not OneNote acceptance. Updated guarded paste fixture to permit exact same fenced source with -RtfCode; check-only build/input layout passes. No actual RTF paste has been executed. Need user selection for one transport-compatibility experiment, not a rerun of plain pre.

User selected RTF sample; attempt PID 1652 refused initial exact fixture guard before clipboard reset or input. Follow-up read-only inspection shows five selected OEs: fences length 3, alpha line length 9, blank length 0, Unicode line decoded length 4, closing fence length 3. Unicode line lacks the expected four source spaces (expected decoded length 8); wrappers contain styles. This means current source does not match the controlled whitespace fixture, not an RTF transport failure. Do not relax exact guard or inject missing spaces into selected source automatically. Ask user to restore four ordinary spaces before Unicode on the source line, then reselect. No RTF paste occurred, no undo needed.

### Correct-source RTF import — whitespace validation FAILED

User restored missing source indentation. PID 24560 passed exact five-line source/paragraph guards, submitted a single RTF-only-rich clipboard paste at 00:35:28.854 UTC, and restored disposable clipboard at 00:35:29.074 UTC (about 220 ms); no helper remains. Read-only export still shows alpha line 9 decoded characters with four leading spaces but Unicode line only four UTF-16 characters (Chinese + surrogate-pair emoji), with no preceding U+0020/NBSP spaces; expected reconstructed fenced payload not found. Thus RTF transport also fails required literal-space preservation in this target sample, despite RichEdit offline roundtrips passing. Do not generalize RichEdit behavior to OneNote or mark code blocks supported. No automatic retry/undo. User should undo this changed experiment once to restore source, not to prove already-established undo behavior again. Current accepted policy is to reject affected fenced blocks pending a proven exact representation; no NBSP/image/layout workaround authorized.

User reports RTF result looks normal and one undo restores; screenshot shows visual indentation. Record visual acceptance and undo, but do not equate visual indentation/layout with literal U+0020 preservation: strict copy-back code-whitespace evidence is still absent and conflicts with decoded export result. Current gated behavior remains reject unproven fenced blocks; no further repeat requested. May later verify actual copied-out text on a designated sample if needed, rather than repeatedly importing.

Prepared one combined distinct fixture for remaining styles: italic, strikethrough, ordered list and plain-character unchecked/checked task boxes. Sixteen block/style offline checks and full suites/check-only build pass; no actual paste performed. This remains fixed payload validation, not the Markdown renderer. User action needed to select exact four-line plain source; only one import attempt should be scheduled per readiness message sequence.

User selected remaining-style fixture. PID 12328 passed exact four-paragraph and repeated evidence checks, submitted one keyboard paste at 00:40:37.990 UTC and restored disposable clipboard at 00:40:38.204 UTC (about 214 ms). Helper exited. Await combined visual result/any losses; do not infer acceptance from SendInput or restoration. No automatic undo/retry.

User supplied screenshot of italic/strike/ordered/task results and reports one undo restored the source. Record distinct fixed-payload style sample and undo as passed by screenshot/report; no claim about parser (not implemented). User requests all remaining tests together rather than serial requests. From now on prepare automated fixtures and batch manual acceptance instructions; do not ask isolated selection/source-copy exercises. No new manual test is presently ready/required. Remaining gates and future batch recorded in manual-acceptance.md; tasks stay incomplete where broad scope exceeds evidence.

### Automated consolidation and gate dependency review

Added RestoreCoordinator with injected backend/clock and twelve ordering/failure/race checks: 200/500/default delay, no lock during wait, duplicate suppression, retained snapshot on busy/publication failure, explicit cleanup-only retry and in-lock newer-copy protection. Fake-backend evidence is not a product Win32 backend. Added SnapshotLimits tests for zero/negative/overflow/per-format/aggregate/count/duplicate budgets; wired the tested limiter into OLE snapshot capture before allocation to avoid the previous post-allocation total-limit check.

Run-AutomatedValidation.ps1 now executes offline suites, check-only paste compilation, optional read-only fixture export, and registry/process residue checks. Ran with -ReadOnlyHost: all suites passed, read-only comparison completed, no probe registrations or test-process residue. No new clipboard mutation or input injection. Broad phase-one tasks remain unchecked where integrated coverage is incomplete.

Identified a dependency cycle: phase 1 requires integrated behavior whose implementation is prohibited until phase 1 fully passes. Created gate-review.md recommending an explicit user-approved move of the gate to installation/release: permit offline phases 2–5, preserve unresolved blockers and deliver a single combined manual acceptance build. This is proposed only; no gate/checkbox or production scope was silently changed. Need user review before actual product implementation, not another unchanged selection test.

### Approved offline preview implementation

User approved gate-review. Updated gate requirement/task 1.8; marked only evidence synthesis approval and concretely completed contract tasks 2.2–2.4, leaving all broad host acceptance tasks incomplete. Added production source Core/Addin modules and build.ps1. .NET Framework/x64 uses installed official PIAs; Markdig 0.37.0 and fixed System.* dependencies restored from NuGet to ignored cache. No SDK install/config changes or COM registration.

Implemented actual AST subset conversion (not fixed payload), whole-selection rejection, bounded output validation/CF_HTML, operation generation cancellation, default disabled/fault state, dedicated STA dispatcher, background parsing, surrogate COM/Ribbon entry, low-level scoped hotkey, physical modifier release wait, conservative selection full-export evidence, independent common-content clipboard snapshot and 200/500 delay/ownership restoration. Snapshot recovery is retained in memory on failure; unknown/privacy-limited source formats refuse. No product installation or real-note testing executed. UI button deliberately refuses if body focus moved; shortcut path is primary until integrated focus behavior is proven.

Build initially detected advanced pipeline injecting LinkReferenceDefinitionGroup into an ordinary heading/list document; narrowed enabled extensions to emphasis extras/task lists/pipe tables (tables recognized then refused). Real parser tests now pass 39 cases including unsafe mixed syntax, unclosed markers, nested/unsupported structures, Unicode and state cancellation. PreviewAudit passes official interface identity, Ribbon signatures and dependency layout without calling OnConnection. Build.ps1 -Test completes with no warnings. Strict OpenSpec validation passes.

Readiness is OFFLINE REVIEW ONLY. Clipboard backend failure injection and host lifetime/performance integration not complete; package license/integrity/controlled installation still pending. Whole task 3/4/5 checkboxes are not ticked from partial code or conservative rejection. Produced README.md, preview-review.md and acceptance/samples.md for one combined review/acceptance. User must review documented limitations before a controlled manually loaded preview; no automatic host activation. Current probe registry keys remain absent.

### Preview review approved; manual loading prepared, not executed

User approved preview-review and requested continuation. Added current-user Register-Preview.ps1 for the new CLSID E036DE73-418B-45F7-973F-A4B09D58E515, manual LoadBehavior=0 only, requires zero ONENOTE processes, refuses existing keys, rollback removes only product registration. Added package dependency licenses and SHA256SUMS; build and -CheckOnly preflight pass (39 Core cases plus official interface audit). Dependency resolver initialized in static Addin constructor before instance initialization to support Core loading from surrogate location; restricts names/versions and avoids capturing a live Addin instance in an AppDomain event.

### First preview command refused normal paragraph QuickStyle — fixed offline

After controlled manual registration/loading, user reports Ctrl+Shift+D no response; screenshot status says 原生样式或嵌套段落尚未支持. This proves command path reached eligibility checks, not a dead shortcut. Read-only fixture inspection: all 16 selected OEs have quickStyleIndex resolving to a QuickStyleDef named p, with only ordinary font/color/spacing attributes; no List/Tag/OEChildren present. The blanket HasAttribute(quickStyleIndex) refusal incorrectly rejects ordinary exported paragraph defaults.

Added tested ParagraphPolicy: permit a uniquely resolved p style with known default-presentation attribute names; retain refusal for missing/duplicate/unknown styles, headings, native list/tags/nesting. Five regression tests added, 44 Core tests and preview contract audit pass. Read-only replay against all 16 selected actual paragraphs returns accepted=True without writing anything. Built corrected binary into artifacts/preview-next, not over the currently loaded preview. Need normal user disconnect/host exit before deploying replacement; no registry changes or automated formatting were performed for this fix. Full combined acceptance remains pending.

User reported exit. Inspection confirms preview surrogate no longer exists; ONENOTE PID 436 remains with MainWindowHandle=0 (do not claim full host exit). Rebuilt corrected sources into registered artifacts/preview path after confirming the component surrogate had unloaded; 44 tests/contract audit passed. Registration CodeBase unchanged, LoadBehavior=0. No process terminated, no re-registration/host activation, no clipboard mutation. Enabled preference file still exists, pending marker absent; do not assume user-disable persisted. Next manual reconnect may therefore enable automatically from preference, and should be treated as enabled unless UI indicates otherwise. Background host lifetime/persistence remain acceptance issues.

### STOP: second host cleanup/startup incident

User screenshot reports OneNote cleaning up previous session and unable to start, advising computer restart. Immediately removed all preview registration via tools/Register-Preview.ps1 -Remove and verified Office Addins/CLSID absent. Only ONENOTE PID 436 remained; no matching preview surrogate. No process terminated, no cache/notebook/config content deleted.

Prior decision to request reopening while an existing host process had not exited was inappropriate and must not recur. Component unload and successful offline build do not establish clean host shutdown. Root cause of retained host/cleanup failure is unresolved; do not claim it is a formatting failure, prove data safety, or re-register. Suspend all host tests/product loading until user recovery confirmed and actual COM lifetime/STA shutdown issues investigated. Any forced host termination needs fresh explicit authorization because recent edits may be unsaved/unsynced.

User explicitly authorized forced termination for this incident. Stopped ONENOTE processes and verified zero remained after two seconds; preview Addins/CLSID/AppID keys all absent. No cache/notebook files deleted. Await user normal plugin-free launch and content/sync check; recovery not yet established. No automatic OneNote restart or test reinstallation performed.

User confirmed plugin-free OneNote opens and page content is complete. Sync and subsequent normal full process exit have not been explicitly confirmed; do not infer them. New registration is now blocked in Register-Preview.ps1 (after non-mutating -CheckOnly and rollback -Remove branches) pending COM lifetime review. No additional host restart requested. Recovery observations must not be used to resume host-loading experiments or declare the second incident root cause resolved.

### Offline actual-Addin lifecycle investigation and mitigation

Source review found three concrete lifecycle hazards: IRibbonUI retained across threads with synchronous InvalidateControl callbacks into Office (potential host/proxy wait cycle); disconnect was only a queued action which could be lost before dispatcher creation; late conversion cleanup was queued through a Post method that refuses work after shutdown. These are code hazards, not a proven diagnosis of the real residual ONENOTE process or startup incident. Earlier full host failure still lacks a wait-chain/dump.

Removed the retained Ribbon COM field and all invalidation callbacks. RibbonLoaded compatibility method now does nothing; XML no longer uses onLoad/getLabel and status is a bounded local status.txt. This deliberately reduces UI capability while investigating; not a silently accepted final status design. Shutdown synchronously sets cancellation/stop flag without waiting on the STA; initialization honors it, a local stop timer covers dispatcher races, finally releases local resources. No Environment.Exit/Thread.Abort/kill-by-name. Late parser completion ends the state even if UI posting is disabled; modifier-release wait checks cancellation before creating a OneNote COM application.

Added actual LifecycleTests running Addin.OnConnection/OnDisconnection and actual STA message loop in a separate bounded test executable with an isolated configuration folder and auto-enable/hooks disabled. 108 checks passed for twenty immediate connect/disconnect races, normal loop shutdown, duplicate callbacks/connections, context/dispatcher cleanup and no restart after shutdown. Core 44 tests and official contract audit also pass. This test does NOT exercise real Office COM objects, parsing-in-flight shutdown or clipboard/hook operations, and cannot prove host exit fixed. No host activation, registration or clipboard mutation was performed. New installation remains blocked; document current status regression before requesting any future review.

Extended the actual-Addin fixture with a controlled pending Task and operation ticket: disconnect invalidates ticket immediately, busy STA stays until the task completes, late completion ends Busy even after UI posting is disabled, and worker then stops normally. Final run passes 117 lifecycle checks, 44 Core cases and contract audit. This uses a synthetic pending task, not a hung real parser or blocked COM call. No host test requested from these results. The offline mitigation is ready for review but real incident root cause is still unproven.

### Continued without premature stop: real product cross-process and negative lifetime tests

Reused registration-free COM marshaling harness under PRODUCT_TEST with the real Addin assembly and isolated test config. Client queries official lifecycle/Ribbon interfaces, passes synthetic nonnull host objects, disconnects; server explicitly checks actual product STA exits before acknowledging normal process exit. Test passes, no registry/OneNote activation. Added no-host-Ribbon-backcall fixture, reflection check that no IRibbonUI field remains, stale enabled+pending marker fault-stop checks, in addition to pending-task and early/duplicate disconnect tests. Final build passes 142 lifecycle checks, 44 Core checks, interface audit and real product cross-process lifecycle. Build errors in newly added test fixture (using order, missing official interface member) were corrected before executing; final run clean.

Read-only Application log query did not yield a relevant Office/.NET/dllhost event proving root cause. Preview registration and matching surrogate absent. Added tools/Get-RecoveryStatus.ps1 and lifecycle-investigation.md. Remaining necessary action is plugin-free normal host-exit baseline after user work/sync; no further offline synthetic tests can establish whether the actual recovered OneNote now fully exits. Do not silently lift registration block based on fake-host tests or claim incident fixed.

### Recovered host exit baseline — PASS

User closed the recovered plugin-free OneNote normally. Two Get-RecoveryStatus.ps1 observations separated by ten seconds show ONENOTE=0, preview surrogate=0, preview registration absent. Thus no-plugin full exit baseline passes. No COM activation/clipboard/input, no forced termination. Actual corrected-plugin host lifecycle remains untested; registration block remains. Next requires user review of a lifecycle-only manual-load comparison with conversion/hook auto-enable forcibly disabled, not a repeat of format/undo acceptance.

Prepared lifecycle-only config gate in source (not enabled on user machine): when marker exists, ignores enabled preference, prevents Enable/Format and hook installation. Added actual-Addin regression with stale enabled preference and button clicks; final 152 lifecycle checks, 44 Core and cross-process suite pass. No marker created in real config, no registration lifted. This allows a narrowly scoped future trial only after explicit incident-review approval.

User explicitly approved lifecycle-only host trial. Added restricted registration switch -LifecycleOnly; normal registration still blocked. Rebuilt and passed 152/44/cross-process suites; verified host and surrogate counts zero, then wrote lifecycle-only marker into existing preview config (old enabled preference preserved but ignored) and registered LoadBehavior=0. No OneNote launched. Worker-only content-free lifecycle logging records started/exited, no filesystem logging in host callbacks. Await manual load and immediate manual disconnection/host close; no conversion or hook may be enabled. Failure means rollback, never start a second host over residue.

### Lifecycle-only trial FAILED full host exit; rolled back

User reported closed. Immediate and ten-second observations: one invisible ONENOTE process (PID 9760), zero matching preview surrogate. lifecycle.log records worker start PID 16108 at 06:07:00 UTC and worker exit at 06:07:27 UTC. Thus own STA and surrogate exited, but host did not. This reproduces the retention without conversion, shortcut hook or real page/clipboard work, narrowing the suspect scope to integration/host lifetime rather than Markdown/rendering. It does not establish exact COM reference cause.

Immediately removed all preview registration; script now blocks even lifecycle-only registration. No second host startup or forced termination. Built a read-only WCT probe, inspected 48 OneNote threads: returned single thread nodes/status 2 or 3 and no reported cycles. This does not exclude COM retention or shutdown waits: COM wait callbacks were not registered, WCT does not expose all native waits, and no stacks/dump were captured. Object names, memory and note contents were not read/logged.

Current required recovery decision: authorize ending the invisible host after ensuring needed edits synced. Do not repeat host loading or treat passing offline checks as resolving it. Keep actual release gate blocked; further root-cause investigation needs appropriate native COM lifetime diagnostics or an architectural reconsideration, not another unchanged user restart loop.

### Authorized residual cleanup and real native RCW counter fixture

User authorized ending PID 9760 and continued offline repair. Ended only that ONENOTE process and verified zero host/surrogates, registrations absent. No automatic relaunch.

Added NativeReferenceTests with an allocated unmanaged IDispatch vtable/AddRef/Release counter to distinguish real RCWs from prior managed test host objects. Actual OnConnection with two RCW arguments leaves native ref count 2 vs fixture baseline 1 after callback; GC in standalone test returns both to 1, weak reference gone. This demonstrates deferred parameter wrapper release, not definite host-retention cause: actual surrogate exit was already observed, which can independently release cross-process references. No blanket FinalReleaseComObject or forced GC/finalizer wait added to product. Microsoft's shared-RCW/in-flight call risks preclude treating manual ref-count draining as a safe fix without evidence.

All 152 actual lifecycle/44 Core/cross-process and native-reference tests pass. No real host activation, registry restoration, memory dump or clipboard operation. Further causal diagnosis now needs evidence unavailable offline: native shutdown stacks/dump or a new minimal integration comparison. Do not ask repeated format tests. Real dump consent needed if chosen because it may contain note text; no dump captured yet.

### Native stack tool prepared and self-tested; approved reproduction armed

User authorized local native debugger preparation and controlled stack reproduction (not memory dumps). Downloaded Microsoft-signed SDK installer from the documented SDK downloads page, laid out only debugger packages and administratively extracted x64 cdb 10.0.26100.9457 into approved temp directory; verified Microsoft signature. An earlier guessed fwlink was actually signed ADK bootstrapper, layout failed 1001; no ADK installation performed, then corrected to documented SDK link 2382321. /? briefly opened bootstrapper help, closed normally. Final debugger layout/extraction succeeded, no SDK-wide install/reboot requested.

Added Capture-LocalStacks.ps1: cdb -pv non-invasive/no-suspend, local-only symbol path, ~*k then qd, no arguments/locals/memory-dump commands. Self-tested against a separate disposable sleeping PowerShell: actual thread stack frames recorded, target still alive after detach, test process then cleaned up. Lack of private Office symbols and non-suspended capture may limit conclusions. Logs remain local ignored artifacts; no upload.

Added a ten-minute watcher for one visible OneNote PID then hidden retained exit; after ten-second natural-exit grace captures twice three seconds apart, auto-removes preview registration in finally, never launches/kills host. Verified clean baseline, registered lifecycle-only stack diagnostic with conversion/hook locked and LoadBehavior=0 under user authorization, started watcher PID 14856. Await one manual load/disconnect/close; normal registration still blocked. No dump authorization inferred. Do not reopen over a residual process.

User reports manually killed the host after closing, invalidating that retained-exit capture; do not label it natural clean exit. On user's request restarted the observation: removed registration, verified host/surrogate zero, re-registered same locked lifecycle-only diagnostic and started watcher PID 13584. No host termination performed by this restart. Current experiment requires leaving the closed-window host untouched for stack collection.

### Retained host native stacks captured successfully

User closed normally without killing. Watcher observed visible PID 14500, window closure, retained process after ten seconds; captured stacks twice (local stacks-14500-20261018-144000.txt and -144004.txt), then automatically removed preview registration. OneNote remains invisible, preview surrogate absent. No dump or target termination performed. All registration paths reblocked after this completed reproduction; no further automatic host trial allowed.

Both samples show the same main-thread return addresses: win32u!NtUserMsgWaitForMultipleObjectsEx → mso20win32client ordinal → onmain nearest exported labels CAsyncResultBase::WaitToComplete / CAsyncResultUsingPollBase::SetThreadAffinity → ONENOTE window/entry frames. Another onmain ExeServer::Init-labelled worker waits in WaitForMultipleObjectsEx; other threads mostly Office/native thread-pool/message/condition waits. No preview/CLR frame is visible in this native host main-thread stack. WITHOUT matching private/public PDBs many names are nearest exports with large offsets; do not interpret those labels as definitive function names or infer which asynchronous operation, event, COM client or lock is responsible. No direct COM RPC-call chain is established by this capture.

This is concrete evidence of persistent host native waiting after the proxy exits, not proof of a specific code fix. Cannot safely patch an unidentified wait by forcing GC, releasing shared incoming COM refs, terminating dllhost or faking callbacks. Current host should not be reopened over residue. Need explicit recovery authorization for PID 14500; further deeper diagnostics would require separate scope for paused debugger/symbols/dump, not an unchanged repeat of user tests.

User explicitly authorized cleanup of PID 14500. Stopped that ONENOTE process, then verified zero host/preview surrogate and no preview registration. Existing stack files remain local for analysis. No automatic restart or plugin registration; content/sync status after this termination has not yet been rechecked by user.

Created acceptance/preview-session.md with one combined manual session and cleanup. Actual host inspection shows OneNote PID 19564 with visible window still running. No registration executed; user must close normally before controlled setup. No notes/clipboard touched. Remaining tasks keep their outstanding integrated coverage; do not mark installation/lifecycle success from script preflight.

Known fixture limitations: restoration failure can leave temporary clipboard; no recovery persistence yet, externally delayed GetData isolated only by standalone supervisor, first-source publication ordering not yet formalized, privacy/protection marker handling and real browser/Notepad cases pending. Task 1.5 remains partial. Treat authorization as sufficient for further non-destructive clipboard tests within the agreed scope; request user action only for a needed new fixture or a materially different decision.

## Status

Partial progress on 1.1, 1.2 and read-only prerequisite investigation for 1.3. No complete task checkbox is warranted yet. A temporary validation Add-in is now compiled and registered per-user with explicit user permission; actual OneNote loading is pending a user-controlled restart. No document or clipboard writes were performed.

## Environment evidence

- OneNote running executable: `C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE`.
- File version: `16.0.19127.20860`.
- Office x64 was established during exploration; x64 typed interop probe successfully ran in this session.
- Windows build: `26200.9457`, DisplayVersion `25H2`, 64-bit OS. Registry ProductName reports `Windows 10 Home China`; use build/version rather than that legacy label to identify the tested environment.
- .NET Framework registry Release `533509`, Version `4.8.09221`.
- Installed dotnet host `8.0.23`; no .NET SDK. Runtimes include Microsoft.NETCore.App and Microsoft.WindowsDesktop.App 6.0.5/8.0.23. Production runtime not selected.
- Available compiler: Framework64 v4.0.30319 csc.exe.
- Installed OneNote interop assembly: GAC_MSIL version 15.0.0.0.
- Actual OneNote paste settings: not yet established. Do not infer them from successful manual tests or undocumented registry values.

## Evidence table

| Check | Source | Result | Limit |
| --- | --- | --- | --- |
| Inline rich-text replacement and one undo | User report during exploration | PASS | Copy source formats and exact paste settings not recorded |
| Multi-paragraph heading/list replacement and one undo | User report during exploration | PASS | Not actual generated CF_HTML or plugin input path |
| Attach through GetActiveObject | Executed PowerShell | MK_E_UNAVAILABLE (0x800401E3) | Does not establish OneNote automation is unavailable |
| PowerShell late-bound Windows properties | Executed PowerShell | Unreliable/null properties; reflection E_FAIL | Binding issue, not evidence that window API is absent |
| Typed x64 C# metadata probe | Compiled and executed validation/OneNoteProbe/Probe.cs | PASS | Does not inspect selection or modify documents |

Probe output:

```text
COM activation: OK
Window count: 1
Window handle present: True
Current page present: True
Quick note: False
Window metadata probe: PASS (not a selection or paste test)
```

## Pending prerequisites

- User designated a disposable page; typed interop GetHierarchy(hsSelf) confirmed current page title exactly matches `Markdown插件验证`. Page body was not read or modified.
- Establish paste preference via supported UI observation and record it without changing user preferences implicitly.
- User authorized minimal test Add-in registration and will cooperate with restart. Await manual restart; do not terminate the current host to force a test.
- No SDK is installed; the existing Framework compiler suffices for the current probe. Choose any additional toolchain only after the minimal host integration investigation.

## Remaining gate matrix

| Task | Evidence still required |
| --- | --- |
| 1.1 | Disposable page established and paste settings recorded |
| 1.2 | Minimal COM Add-in load, Ribbon callback, unregister/unload and startup measurements |
| 1.3 | Exact ranges, repeated text, paragraphs, mixed objects, focus, views and multiple windows |
| 1.4 | Real generated HTML, one undo, surrounding structure, code whitespace and read-only behavior |
| 1.5 | Independent clipboard snapshots across supported formats and media; failure matrix |
| 1.6 | Verified paste consumption/completion and clipboard restoration, including races and failure |
| 1.7 | Shortcut scope, release, modifiers, repeat and Ribbon selection behavior |
| 1.8 | Consolidated go/no-go after the above checks |

Full implementation (phases 2–6) remains gated. No technical impossibility has been concluded from this metadata probe.

## Minimal Add-in registration

- Source: `validation/OneNoteProbe/ProbeAddin.cs`; compiled x64 with the existing Framework compiler.
- Per-user ProgID: `OneNoteMarkdown.ValidationProbe`.
- CLSID: `{AE4143BE-ABDB-4B51-989B-D894A7BD63D5}`.
- Registry scope: HKCU Software/Classes for this CLSID/ProgID, and HKCU Software/Microsoft/Office/OneNote/Addins/OneNoteMarkdown.ValidationProbe. LoadBehavior=3.
- Registration script refuses existing keys and rolls back newly created keys on failure. Registration completed successfully; this is not evidence of host loading.
- Expected Ribbon tab: `Markdown Probe`; button: `Verify callback`. Callback writes a content-free diagnostic record only.
- Log: `%LOCALAPPDATA%\OneNoteMarkdownValidation\probe.log`. Elapsed time is since probe type initialization, not total OneNote startup time.
- Unregister: run `validation/OneNoteProbe/Register-Probe.ps1 -Remove`, then manually restart OneNote. Actual unloading and clean startup still require verification.

## First host restart and activation correction

- User reported no Ribbon tab. Host PID changed to 17504; LoadBehavior changed from 3 to 2 and no callback log existed. Initial registration did not successfully connect the Add-in.
- In-process activation from PowerShell succeeded; QueryInterface for IDTExtensibility2 and IRibbonExtensibility both returned S_OK. A subsequent ReleaseComObject call failed because PowerShell returned an unwrapped managed object; that cleanup error is not a COM activation failure.
- Compared the public OneMore registration (`https://raw.githubusercontent.com/stevencohn/OneMore/main/OneMoreSetup/Registry.wxs`): the initial probe lacked CLSID AppID linkage and AppID DllSurrogate registration.
- Updated the per-user registration script to add those entries and remove them on rollback, then re-registered this probe only. No machine-wide permissions were changed.
- Compiled and executed `SurrogateProbe.cs`: CLSCTX_LOCAL_SERVER activation for IRibbonExtensibility returned 0x00000000, and cross-process GetCustomUI returned expected XML. Log PID 25884 recorded GetCustomUI, independently of OneNote PID 17504. This log entry is from the standalone probe, NOT evidence of OneNote connection.
- Actual OneNote loading and button callback remain pending a second manual restart. Successful surrogate loading may revise the assumed location of the plugin entry in the architecture; inspect actual callback PIDs before finalizing that decision.

## Surrogate-exit mitigation (offline evidence and implementation)

Cross-reference evidence collected offline:

- `register-onemore.reg` proves OneMore uses the identical architecture: AppID `{88AB88AB-CDFB-4C68-9C3A-F10B75A5BC61}` with empty `DllSurrogate`, `InprocServer32=mscoree.dll`, ThreadingModel=Both, LoadBehavior=3. Not in-proc.
- OneMore `TechNote - COM Surrogate` documents the MTA surrogate constraint and their dedicated STA helper; our dedicated STA worker follows the same practice.
- OneMore `AddIn.cs` `OnDisconnection` unconditionally runs `GC.Collect()` then `Environment.Exit(0)` with the comment "this is a hack, modeless dialogs seem to keep OneNote open :-(" — a production practice since 2016 that terminates the surrogate process on disconnect.
- OneNoteDev/VanillaAddIn issue #2: Microsoft's official sample add-in also leaves an invisible ONENOTE.EXE after close and breaks the next start with "OneNote is cleaning up"; never fixed. The symptom class is not specific to this project.

Evidence gap identified: earlier records of "proxy exited but host still waits" were never verified while suspended — the watcher previously observed only ONENOTE, and `Get-RecoveryStatus.ps1` ran only after cleanup. `Watch-ExitStacks.ps1` now records the preview dllhost count at window-close, PASS and retained checkpoints.

Implementation:

- `src/Addin/Addin.cs`: `OnDisconnection` calls `Shutdown()` then `ForceExitSurrogate(mode)`. Real-host mode only: waits (up to 5s) only while an operation is busy so an in-flight clipboard restore can finish, writes a `surrogate-exit mode=...` lifecycle line, then `GC.Collect()` + `Environment.Exit(0)` to end this surrogate process. The host process is never touched. Offline constructions (lifecycle tests, probes) return immediately, preserving the 500ms disconnect-callback bound.
- `tools/Register-Preview.ps1`: loading now requires `-FixTrial` and references the reviewed plan in `lifecycle-investigation.md`; `-Remove`/`-CheckOnly` unchanged; the trial still writes the `lifecycle-only` marker (conversion and hotkey locked).
- `tools/Watch-ExitStacks.ps1`: records preview dllhost presence at the three decision checkpoints.

Reviewed fix-verification trial (pending one user session): round 1 load-then-close-directly (host-shutdown path, expect `mode=ext_dm_HostShutdown`), round 2 load-then-uncheck-then-close (user-closed path, expect `mode=ext_dm_UserClosed`). Expected outcome both rounds: ONENOTE and preview dllhost fully exit within seconds; watcher records PASS and removes registration. Decision rules and constraints are recorded in `lifecycle-investigation.md`. Installation/release remain gated separately.

## Surrogate-exit fix-verification results (PASS)

Two sessions, both disconnect paths each: total 4/4 clean.

| Session | Round | Path | Lifecycle evidence | Watcher verdict |
| --- | --- | --- | --- | --- |
| A (original) | 1 | load then close directly | `surrogate-exit mode=ext_dm_HostShutdown idle` at 07:26:10 | `PASS host process exited naturally; preview dllhost=0` at 07:26:10 |
| A | 2 | load then uncheck then close | `surrogate-exit mode=ext_dm_UserClosed idle` at 07:30:09 | `PASS ... dllhost=0` at 07:30:23 |
| B (fresh cold boot) | 1 | load then close directly | `surrogate-exit mode=ext_dm_HostShutdown idle` at 10:23:24 | `PASS ... dllhost=0` at 10:23:27 |
| B | 2 | load then uncheck then close | `surrogate-exit mode=ext_dm_UserClosed idle` at 10:24:41 | `PASS ... dllhost=0` at 10:24:46 |

All four: preview dllhost=0 at every checkpoint, registration auto-removed by the watcher finally block, post-trial `Get-RecoveryStatus.ps1` all zero (host, surrogate, registration). The host fully exited within roughly one to three seconds of window close; most rounds never entered the "window closed but process still alive" intermediate state. Contrast: before the fix the same actions retained the host for 10+ seconds and produced the captured native wait stacks.

Conclusion scope: this removes the reproducible host-retention failure and corrects the earlier "proxy exited but host still waited" record (proxy liveness was never verified at suspension time; the fix has the proxy exit before host close and the host then exits). Root cause remains an internal OneNote lifecycle behavior, not proven with private symbols. The trial exercised only load/unload/close: conversion, hotkey and clipboard stayed locked by design, so clipboard safety relies on the bounded wait (≤5s only while an operation is in flight) and was not exercised mid-hang. Adopting this behavior requires revising the previously approved "no Environment.Exit" rule in design.md and specs/addin-lifecycle/spec.md.

## Real-host functional acceptance (PASS) and sibling-run defect fix

First full-functional acceptance run (registration `-Acceptance`, conversion and Ctrl+Shift+D enabled) exposed two distinct outcomes that were both initially perceived as "hotkey dead":

1. Defect (fixed): `Selection.Read` parsed **every** `T` sibling in the selected paragraph with the strict fragment reader. `TextFragment.Parse` throws on `font-weight:bold`/italic/`text-decoration`, so after a first conversion inserted a bold span, a second conversion in the same paragraph threw on the **non-selected** bold sibling and was swallowed by the generic catch as "操作未支持，已跳过". Root cause and analysis: the strict reader (which extracts text **and** enforces "selected text must be plain") was reused only to compute the `whole` flag from sibling length; the spec's own "plain-text **selection**" + "选区外内容保持不变" (`specs/safe-selection-formatting/spec.md:4,63,67`) never required the paragraph to be unformatted. Fix: only the selected run is parsed strictly; sibling runs use `TextFragment.HasContent` for the `whole` flag and are otherwise ignored (they survive the paste). Selecting already-formatted text still refuses, now with a surfaced message; the Addin catch filter also surfaces Core-assembly messages instead of a generic string. `ParagraphPolicy.IsPlain` intentionally never inspected inline spans, confirming the restriction was over-applied, not a required invariant.

2. Not a regression: a run refused with "剪贴板含未支持或保护格式，已跳过" because the clipboard held browser-private formats (`Chromium internal source RFH token` / `Chromium internal source URL`). Per `specs/safe-selection-formatting/spec.md:32` unknown/protected formats must be refused rather than silently dropped; this is expected behavior of the clipboard policy, unrelated to the hotkey. It only blocks when the clipboard source is a Chromium app.

Offline suites after the fix: 47 core (3 new fragment cases) + all lifecycle/contract/cross-process/native suites PASS. Real-host re-run with a disposable plain-text clipboard: two `**测试文字**` tokens in one paragraph both converted in sequence, two Ctrl+Z restored, `status.txt` ended "已提交；剪贴板已恢复", `surrogate-exit mode=ext_dm_HostShutdown idle` at 10:57:53 with host PASS and clean recovery status. Specs (`safe-selection-formatting`) and `implementation-contract.md` were updated to lock the "strict parse applies only to the selected run" rule and prevent regression.

## Clipboard feedback (C) and classified metadata policy (A)

Follow-up work approved by the user after the browser-clipboard refusal was misread as a hotkey failure. Both items were implemented and offline-verified; real-host visual verification is pending.

- C (feedback): skip/failure statuses now also raise a non-focus-stealing tray balloon (`NotifyIcon`, `BalloonTipIcon.None`), gated by `!offlineLifecycle && !shuttingDown` so offline fixtures never show UI. The balloon is hidden on `Shutdown`/disable and disposed in the worker `finally`; the `Environment.Exit(0)` surrogate-exit also ends it. Message text is metadata only (reason plus the offending clipboard format name) and never note text, HTML or clipboard bytes. Clipboard refusals now name the format (`剪贴板含未支持格式：<name|#id>`) instead of a generic string. Alerts fire for every skip/refusal/error path (conversion timeout, selection change, `no-markdown`, unsupported syntax, unsupported clipboard, restore-not-clean, etc.); clean successes stay file-only. `addin-lifecycle/spec.md` gained a "Skip or failure feedback" scenario; `implementation-contract.md` records the rule; added task 3.5.
- A (classification): `Chromium internal source RFH token` and `Chromium internal source URL` are classified as auxiliary metadata and skipped in `ClipboardSession.Capture` (not saved, not restored, no refusal), preserving the promised HTML/text content. `safe-selection-formatting/spec.md` "Classified auxiliary metadata" scenario names the initial list; unknown formats are still refused and the list is extended per class as encountered. Option B (blindly replaying all HGLOBAL private formats) was explicitly declined: `HGLOBAL` does not guarantee self-contained bytes (handle/pointer and synthesized-GDI formats exist), fan-out `GetData` on all formats adds STA stall risk, and it still would not cover OLE/embedded objects provided via `TYMED_ISTORAGE`.

Offline after C+A: addin recompiles with `System.Drawing` reference; 154 lifecycle + 47 core + contract/cross-process/native suites PASS; `openspec validate --strict` valid. Not yet done: real-host observation of the balloon and of a browser-origin clipboard actually proceeding end-to-end.

### Real-host result and feedback channel correction

Real-host run: browser-origin clipboard (Chromium) **passed** — a valid `**测试文字**` selection converted normally, confirming A works end-to-end. The skip path reached `Alert` (`status.txt` = `跳过：no-markdown`) but **no tray balloon appeared**: `NotifyIcon.ShowBalloonTip` is suppressed in this environment (Win11 notification settings / no AppUserModelID shortcut). The tray balloon is therefore rejected as the sole channel per the spec rule "MUST NOT depend on a mechanism that OS notification settings can suppress".

Replaced with a self-drawn `Toast` form: borderless, `ShowInTaskbar=false`, `TopMost`, positioned bottom-right of the working area, `WS_EX_NOACTIVATE|WS_EX_TOOLWINDOW` plus `ShowWithoutActivation` so it never takes focus, 6-second auto-dismiss, single instance (a new alert closes the previous), closed via the STA dispatcher on disconnect and in the worker `finally`. No `NotifyIcon` remains. Offline suites after the change: 153 lifecycle (field-enumeration count shifted because the alert field changed; coverage unchanged) + 47 core + contract/cross-process/native PASS; `openspec validate --strict` valid. Real-host popup **confirmed** (skip path showed the bottom-right `Toast` without taking focus; host exited naturally with `preview dllhost=0`).

### Real-host popup confirmed; readable clipboard messages

The `Toast` popup was confirmed on the real host (`跳过：no-markdown`). The Explorer-file-copy refusal also fired and reached the popup, but the raw message (`剪贴板含未支持格式：Shell IDList Array，已跳过`) was judged unintelligible to a user. `ClipboardSession.Describe` now maps known format names to a plain-Chinese label while keeping the raw name in parentheses for diagnostics, e.g. `剪贴板含未支持的内容：资源管理器复制的文件或文件夹列表（Shell IDList Array），已跳过`. Mappings cover Explorer file lists, `FileNameW`/`FileName`, `Preferred DropEffect`, Excel `Biff*`/`Csv`/`XML Spreadsheet`, `Hyperlink`, embedded `Object Descriptor`/`Ole Private Data`, and Word source. The image/privacy/native-fallback refusals were reworded to plain Chinese as well. Message text remains metadata only (reason plus format label/name); no note text, HTML or clipboard bytes. Offline suites and `openspec validate --strict` pass after the change; real-host re-observation of the friendlier text is **confirmed** (Explorer file-copy refusal showed the plain-Chinese label with the raw name in parentheses; host exited naturally with `preview dllhost=0`).

## Progress checkpoint (2026-10-18)

Paused at the user's request to archive current progress. The change is intentionally **not** moved to `archive/`: installation/release remain gated (see `gate-review.md`) and tasks 1.1–1.7, 2.1, 3.1–3.4, 4.x, 5.x and most of 6.x are still open. This checkpoint records only what is verified.

Environment left clean: `OneNoteProcesses=0`, `PreviewSurrogates=0`, `PreviewRegistered=False`, `CleanForFutureControlledLoad=True`; watcher ended with `PASS host process exited naturally; preview dllhost=0` and auto-removed the registration. `artifacts/preview` holds the latest offline build.

Verified on the real host this session:
- Surrogate-retention fix: 4/4 exit trials (two cold sessions × {`ext_dm_HostShutdown`, `ext_dm_UserClosed`}); host always exited naturally with `preview dllhost=0`; `Get-RecoveryStatus` clear each time.
- Selection strict-parse fix: two `**测试文字**` tokens converted back to back in one paragraph; two Ctrl+Z restored; `status.txt` ended `已提交；剪贴板已恢复`.
- C (feedback): non-activating `Toast` popup for every skip/failure, confirmed on the host without stealing focus; the OS-suppressed tray balloon was removed.
- A (classification): Chromium-origin clipboard accepted end-to-end (conversion succeeded); Explorer file-copy refused with a readable message; friendly format labels confirmed.

Implemented and offline-green but not yet host-verified: offline lifecycle/race fixtures for the popup lifecycle; forbidden single-format clipboard snapshot paths beyond browser/Explorer/Excel; restore-failure persistence; fenced code block whole-paragraph rejection on the host.

Resume point: `& .\tools\Register-Preview.ps1 -Remove; & .\tools\Start-ExitStackWatch.ps1; if($?){ & .\tools\Register-Preview.ps1 -Acceptance }`, then continue the remaining acceptance scenarios and gate items.

## Hard-blocker work (A) — offline parser/refusal boundaries

Extended `tests/CoreTests.cs` by 18 cases (core 47 → 65) to close the gate-review gap on "malformed/unsupported Markdown rules and supported-syntax nesting boundaries" (gate-review.md:22, tasks 4.1/4.4) and to lock the fenced/indented code whole-paragraph rule (gate-review.md:21):
- Refused: star `*` and plus `+` bullets; setext heading; level-four heading; NUL character; unclosed fence; tilde fence; indented code combined with a list; fenced code mixed with other syntax (whole selection skipped); dangling marker across inline types; dangling marker with emoji.
- Accepted: heading with inline bold; multiline quote with inline; bold wrapping inline code.
- Reason strings asserted for traceability: `fenced-code-fidelity-unverified`, `whole-paragraphs-required`.
- Code content not reinterpreted (`<b> & *literal*` inside code stays literal).

All offline suites re-run green (65 core + 153 lifecycle + contract/cross-process/native). `Watch-ExitStacks.ps1` now records `host process started pid=…` so launch-to-window latency can be derived from the watcher log for the pending startup-overhead measurements (tasks 1.2/6.2).

### Two OneNote editions and the "multiple hosts" false abort

Root cause of the intermittent `STOP multiple hosts` aborts (which triggered the safety rollback and made the add-in disappear from COM add-ins): **two OneNote editions are installed and both use the process name `ONENOTE.EXE`**:
- Desktop C2R (the target; supports COM add-ins): `C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE`, 16.0.19127.20860.
- Store/UWP (`Microsoft.Office.OneNote_16.14326.22589.0_x64`, under `WindowsApps`): a UWP app that **never** loads this COM add-in.

Launching "OneNote" from the Start menu / a pinned (Store) tile could start the Store app alongside or instead of the desktop app, producing 2–3 `ONENOTE.EXE` processes. Behaviour when launched via the explicit desktop path was clean: one fresh host, `parent=explorer`.

Tooling updated to disambiguate by executable path (`Path -notlike '*\WindowsApps\*'`): `Watch-ExitStacks.ps1`, `Measure-HostStartup.ps1` and `Get-RecoveryStatus.ps1` now ignore the Store edition (reported separately) instead of aborting.

### Startup-overhead samples (tasks 1.2 / 6.2)

`tools/Measure-HostStartup.ps1` measures process start → first non-zero main window and → exit, desktop edition only. Samples (n=1 each, indicative; OneNote warm):
- No registration (baseline-3): launch→window 272 ms, launch→exit 2853 ms.
- With registration, LoadBehavior=0 (registered-3): launch→window 264 ms, launch→exit 4397 ms.

Because the preview is manual-load (`LoadBehavior=0`), the add-in does not load at host startup, so registry presence adds no measurable startup cost; the metric reflects the host, not the add-in. Caveat: the number is window-handle creation, not "ready", and n=1 per arm — a proper n≥3 run is still owed before this closes task 1.2/6.2. The add-in's own load cost (surrogate spawn + Markdig init) is not yet isolated with a dedicated trace.

### Startup-overhead n≥3 and add-in load cost (tasks 1.2 / 6.2)

`TraceLifecycle` now records `+Nms-since-surrogate-start`, and `Measure-HostStartup.ps1` supports `-Samples N`; both launch-to-window and launch-to-exit are logged per sample. All samples used the desktop edition via the explicit path.

| Arm | launch→window (ms) | notes |
|---|---|---|
| No registration | 272, 259, 208 | mean ≈246; earlier `baseline-3` is the 272 sample |
| Registered, LoadBehavior=0 | 264, 259, 300 | mean ≈274 |

The ranges overlap; with manual load the add-in never loads at startup, so there is no measurable startup-cost difference. launch→exit is dominated by user action and is not a plugin metric. Add-in's own load cost from a load session (user enabled the add-in): `addin-connected` at **+53 ms** and `worker-started` at **+69 ms** since surrogate start — i.e. the STA thread + message loop + hooks are up in ~69 ms; Markdig initialization is lazy (per first conversion) and is not part of this number.

### Double `worker-exited` trace and idempotent finalizer

One load session produced two byte-identical `worker-exited` lines (same pid, same millisecond) with a single `addin-connected`/`worker-started`. `Run()` has exactly one `finally`, so the likely cause is the `Environment.Exit(0)` in `ForceExitSurrogate` re-entering the worker teardown during CLR shutdown. Guarded it: the teardown now runs under `Interlocked.Exchange(ref workerFinalized,1)` so `Unhook`/dispose and the trace happen at most once. Offline suites green (154 lifecycle — field-count shift, coverage unchanged). Host re-observation still owed.

### User-confirmed host scenarios (2026-10-18)

The user reports these pass on the real host (user-observed, not machine-logged here), which closes the remaining group-A acceptance scope:
- Read-only page: hotkey action refused, no write.
- Hotkey scoping: `Ctrl+Shift+D` does not act outside the target editing environment (consistent with the hook being installed only inside the OneNote process, and with the offline "hook is zero on disable" check).
- Word/Excel (and Notepad, browser, Explorer) clipboard sources: refuse or accept as classified, with a readable message.

Task bookkeeping: 1.7 (hotkey scoping/disabled passthrough/single-flight/ribbon-button focus) is marked done. 1.4 and 1.5 stay open on purpose because two sub-clauses are genuinely unmet regardless of these confirmations — 1.4's "代码空白" (code whitespace; fenced code is refused, task 4.5) and 1.5's "外部延迟渲染" (see gate-review.md:20).

Real-host fence test: the fenced-code refusal fired correctly but the popup showed the raw code `跳过：fenced-code-fidelity-unverified`. Added `Addin.Notice`/`DescribeSkip`: the popup maps known reason codes to plain Chinese (`代码块暂不支持（空白保真未验证）`, `需要选中完整段落`, `选中内容不含 Markdown`, `含不支持的语法`, `选中内容过长`, …) while `status.txt` keeps the raw code for diagnostics. Matches the earlier clipboard message readability change.

Operational note: a run aborted with watcher `STOP multiple hosts` (two `ONENOTE.EXE` instances) and the safety rollback then removed the registration, so the add-in vanished from the COM add-ins list mid-session. Ensure exactly one host is running before launching for acceptance.

### User guide and residual risks (task 6.4)

Wrote `usage.md` (root): the end-user guide covering tested environment, hotkey/ribbon usage and single-undo, selection requirements and refusals, the supported syntax subset, the skip categories with their Chinese popup reasons (raw codes stay in `status.txt`), clipboard snapshot/restore rules (200 ms default, 500 ms opt-in, ownership/sequence guard, privacy flags, unsupported-format refusal), manual fault recovery (`operation.pending` → re-enable, "retry clipboard restore"), the tested support table, and the residual-risk/unsupported list. Cross-links from `README.md`.

Also refreshed the stale `README.md` banner, which still claimed loading was paused, the install path fenced, and safe host exit unverified — all superseded by the verified surrogate-exit fix and host acceptance. README now states: manual-load preview, host exit fix verified, functional acceptance passed, install/release still gated; it lists the still-open items (6.3 scripts, 1.6/5.5 fault injection and restore-failure persistence, `worker-exited` re-observation, external delayed rendering, idle CPU/memory and conversion timing) and points at `usage.md`.

Task 6.4 marked done. 1.4/1.5 remain open (unmet sub-clauses as noted above); this is documentation only and does not mark any unverified behaviour as supported.

### Install / uninstall / rollback (task 6.3)

Added `tools/Install-Addin.ps1` — the formal per-user (HKCU) install path with `-DryRun`, `-Verify`, `-Uninstall`, integrity-checked build input, host-must-be-closed guard, read-back verification, and a transactional rollback. It deliberately does **not** implement `LoadBehavior=3`: enabling/disabling stays in OneNote's COM Add-ins dialog, consistent with the gate.

Rollback/isolation is proven offline by `tools/Test-InstallRollback.ps1`, which runs the real installer against `HKCU\Software\OneNoteMarkdownInstallTest` seeded with two foreign add-in keys and one foreign CLSID key. 20 assertions passed:

- dry run creates nothing and reports planned keys;
- install creates exactly the four preview keys, `LoadBehavior=0`, foreign keys byte-identical (SHA-256 of `reg export`);
- `-Verify` reports `STATE=REGISTERED` then `STATE=NOT_REGISTERED` after uninstall;
- uninstall removes only the four keys; foreign keys unchanged;
- injected mid-install failure (stage `app`) exits non-zero, removes all four keys, leaves foreign keys unchanged;
- injected failure (`stage class`) with a pre-existing preview key restores that key byte-for-byte and removes the other three.

Real-root `-Verify` was run (read-only): `STATE=NOT_REGISTERED`, no siblings. The sandbox is deleted at the end of every run; no real registration, host launch or clipboard access occurs. Steps documented in `acceptance/install.md`. A live install still requires the gate to open; this task records scripts + failure-isolation evidence only.

PowerShell detail worth keeping: `2>&1` on a native command under `$ErrorActionPreference='Stop'` turns the child's stderr into a terminating `NativeCommandError` in 5.1, which aborted the harness on the intentional-failure case; the harness now runs those calls under `Continue` and reads `$LASTEXITCODE`.

### Fault injection — offline module matrix (tasks 5.5 / 1.6)

Extended `tests/CoreTests.cs` (65 → 76) with the operation-state fault matrix, all deterministic and host-free:

- duplicate request rejected (single flight);
- late response after `Disable()` rejected (`Valid(ticket)=false`);
- no new op while busy (no overlapping restart);
- `End()` then `Enable()` is the manual re-enable path after cancel;
- generation advances across restart, so a stale ticket stays invalid;
- `Fault()` stops and invalidates in-flight work; recovery again needs `End()` + `Enable()`;
- restore delay policy accepts only 200/500 (any other value, including 0/-1/300, falls back to 200).

These cover the state-machine half of 5.5 (duplicate/late/disable/uncertain commit is distinguished by `sent==4` in `Addin.cs:271`) and the 200/500 half of 1.6. `Addin.cs:262/264` re-validate `state.Valid` **and** `selection.Same(...)` before and after publish, which is the code path the page-change/selection-change/typing cases exercise on the host.

The host half (page change, selection change, typing, second window, oversized input, third-party clipboard race during 200 ms, in-flight disable, engine crash, paste-result-unknown) is enumerated in `acceptance/fault-injection.md` with the exact expected popup/`status.txt` strings; it needs a supervised host session. The engine-crash case requires killing the add-in's `dllhost` and is explicitly marked as needing separate authorization. Tasks 1.6 and 5.5 remain open pending that session; the offline evidence does not mark them supported.

### Fault injection — supervised host session (tasks 5.5 / 1.6)

Ran the supervised session (`Register-Preview.ps1 -Acceptance`, desktop edition only, exit-stack watcher live) with the user driving the manual steps and the surrogate kill pre-authorized.

| Case | Result |
|---|---|
| page change / selection change / typing during op | user-reported pass (pre-commit cancel path) |
| second OneNote window | pass (refused) — but see watcher note below |
| oversized input (`32769` chars from Notepad) | pass (`跳过：选中内容过长`, no hang) |
| clipboard race inside the restore window (500 ms) | pass (third-party content preserved) |
| in-flight disable | pass (cancel, no write) |
| paste result message (`sent==4`) | confirmed via `status.txt` = `已提交；剪贴板已恢复` |
| engine crash (authorized) | pass — see below |

Engine crash: killed only the add-in's surrogate `dllhost` (`pid=9964`, `CommandLine ... /Processid:{E036DE73-...}`); the host (`ONENOTE 6160`) was untouched. `lifecycle.log` for 9964 has `addin-connected` + `worker-started` and **no** `surrogate-exit`/`worker-exited`, i.e. abrupt termination with no graceful trace. Post-crash UI observed by the user: the Ribbon tab is still present (it is a static definition the host already holds) but `Ctrl+Shift+D` and the 启用 button are inert and no false success is shown — the intended "crash requires reload, UI inert, never lies" behaviour. Recovery: normal exit → re-register → relaunch → re-enable → one conversion at the default 200 ms → `status.txt` `已提交；剪贴板已恢复` → clean natural exit (watcher `PASS host process exited naturally; preview dllhost=0`).

Operational note (recurring): opening a **second desktop window** spawns a second `ONENOTE.EXE`; the exit-stack watcher's safety rule aborts with `STOP multiple desktop hosts` and rolls the registration back, so the add-in disappears from the COM Add-ins list until re-registered. This is the same class of event as the earlier "two editions / multiple hosts" note; it is the watcher behaving as designed, not an add-in fault.

Still not directly observed: the `提交结果不确定` branch (`Native.Paste()` not returning 4) and the restore-failure path — both remain code-path-only; a deterministic restore-failure injection (holding the clipboard open across the restore window) is the planned check before 1.6 closes.

### `worker-exited` trace is now deterministic

Follow-up on the double-trace anomaly. The line is written from two independent places on two threads: the worker thread's `finally` (`Addin.cs:90`) and, on graceful shutdown, `ForceExitSurrogate` calling `Environment.Exit(0)` (`Addin.cs:328`). Historically the worker `finally` won the race and logged once (every pid in the log has exactly one line) except `pid=11112`, which logged twice. The `Interlocked` guard fixed the duplicate but the count became racy in the other direction: the post-fix host run (`pid=15448`) logged **zero** `worker-exited`, because `Environment.Exit(0)` won.

Fixed by making the trace itself once-only via a shared flag: `TraceExited()` (guarded by `Interlocked.Exchange(ref exitTraced,1)`) is now called by both the worker `finally` and `ForceExitSurrogate`, so exactly one line is recorded whichever path runs first, and the process ends immediately after. Added an offline assertion in `LifecycleTests` (now 156 checks) that a normal disconnect produces exactly one `worker-exited` line. Points at a deterministic record for the 6.2 exit-cleanup evidence.

Host confirmation after the fix: a fresh supervised session (`pid=21828`) logged `addin-connected`, `worker-started`, `surrogate-exit`, `worker-exited` — `surrogate-exit=1 worker-exited=1`, exactly one each — with a successful conversion (`status.txt` = `已提交；剪贴板已恢复`) and a natural host exit (watcher `PASS host process exited naturally; preview dllhost=0`). The double-trace anomaly is closed.

### Gate decisions and consolidated acceptance (2026-10-18)

The user closed the remaining blockers with explicit decisions, recorded in `gate-review.md` ("安装授权" and "用户决策"):

- **B1 install approved** — installation via `tools/Install-Addin.ps1` is authorized, limited to `HKCU` per-user with `LoadBehavior=0` (manual load). `LoadBehavior=3` auto-load remains unauthorized and is still not implemented.
- **B2 acceptance passed** — the supervised host acceptance (lifecycle/exit fix, conversion + single undo, strict-selection fix, clipboard classification, non-activating feedback, hotkey scope, read-only refusal, fault-injection matrix, deterministic single `worker-exited`) is recorded as passed.
- **B3 external delayed rendering — ignored** (whitelist + default refuse covers the risk).
- **B4 restore failure / paste-result-unknown — ignored** (restore failure is covered by the offline coordination tests; the `提交结果不确定` branch stays code-path-only and is not claimed as verified).
- **B5 performance — to be measured** (idle CPU/memory, conversion time); still open.
- **B6 full-syntax host retest (6.1) — user-confirmed passed**.
- **B7 spec amended** — `addin-lifecycle/spec.md` now states the recent result is kept as a content-free local record (`status.txt`) with a non-activating hint on skip/failure, matching the implementation (the Ribbon keeps only a static pointer; no `IRibbonUI` is retained, by design).
- **B8 fenced/indented code blocks — known limitation, planned for a later version** — `markdown-subset/spec.md` now lists them as out of this version's scope, keeps whole-selection refusal, and the friendly skip text is `代码块不在本版本支持范围（计划后续版本）` (raw code `fenced-code-fidelity-unverified` unchanged for tests/diagnostics).
- **B9 private clipboard format table — ignored** (unknown formats are refused by default).
- **B10 signing and packaging — required**, positioned as an **open-source project** (not personal/local use).

Consequently tasks 1.1–1.6, 2.1, 3.1–3.4, 4.1–4.5, 5.1–5.4 and 6.1 were closed with evidence notes; 6.2 (B5) and 6.5 (final delivery check) remain open, along with the B10 packaging work.

### Host footprint and conversion timing (B5, task 6.2)

Measured with `tools/Measure-HostFootprint.ps1` (read-only sampler; the add-in now traces `conversion-ms=<n> chars=<m>` in `lifecycle.log`), single desktop host with the add-in loaded and enabled:

| Metric | Value |
|---|---|
| Idle CPU (20 s, 16 logical cores) | ~0.00% (single 0.09% sample) |
| Host `ONENOTE.EXE` (pid 7508) | working set 273 MB, private 185 MB |
| **Add-in surrogate `dllhost` (pid 23068)** | **working set 39.2 MB, private 28.2 MB**, cumulative CPU 0.4 s since start |
| Conversion, first (14 chars) | 68 ms (includes Markdig init/JIT) |
| Conversion, warm (23 chars) | 3 ms |

The host figures are the OneNote baseline; the add-in's own cost is the surrogate row (~39 MB WS, ~28 MB private, no measurable idle CPU). Against the phase-2 budgets (32768-char input, 2 s parse wait) both conversion samples are far inside the budget; the first-conversion cost is one-off initialization. B5 closed; task 6.2 marked done.

### Open-source packaging and signing (B10, task 6.5)

Positioned the project as open source: added `LICENSE` (MIT), `THIRD-PARTY-NOTICES.md` (Markdig BSD-2-Clause; System.Memory/Buffers/Numerics.Vectors/Runtime.CompilerServices.Unsafe MIT; Office PIAs are referenced from the local GAC and not redistributed), and `VERSION` (`0.1.0`).

`build.ps1` now stamps `AssemblyVersion`/`AssemblyFileVersion` from `VERSION` via a generated `Version.g.cs`. The legacy `csc.exe` does not accept `/version:` (error CS2007), and PowerShell variable names are case-insensitive so the `-Version` parameter collided with the dependency loop's `$version`; both are handled (generated attributes, and a separate `$versionStamp`). Verified: both assemblies report `0.1.0.0`.

Added `tools/Package-Release.ps1`: runs `build.ps1 -Test`, re-verifies `SHA256SUMS.txt`, stages `bin/` (DLLs + per-package license files), `tools/`, `docs/`, `usage.md`, `README.md`, `LICENSE`, `THIRD-PARTY-NOTICES.md`, `VERSION`, and a generated `VERSION-INFO.txt` + `SHA256SUMS.txt`, then produces `artifacts/release/OneNoteMarkdown-0.1.0-x64.zip` (22 files) plus `OneNoteMarkdown-0.1.0-x64.zip.sha256` (`3FDF87A9…740E`). Optional Authenticode signing via `-Thumbprint` uses `signtool` (sign + verify with SHA-256 + timestamp); without a certificate the archive is unsigned, which is the normal open-source posture, with checksums as the integrity evidence. No registry, host or clipboard access.

With B5 and B10 done and tasks 6.2/6.5 closed, every task and every gate item is now closed or explicitly recorded as an accepted limitation (B3/B4/B8/B9 and the two Rejected-with-record clipboard sub-clauses). `openspec validate --strict` passes.

### Formal installation (B1 authorization, executed)

Ran `tools/Install-Addin.ps1` against the real root (no sandbox) with no host running; `-Verify` reports `STATE=REGISTERED` and the five keys resolve exactly as designed: `HKLM`-free HKCU `CLSID\{E036DE73-…}` (`AppID`), `InprocServer32` (`mscoree.dll`, `Assembly=OneNoteMarkdown.Addin, Version=0.1.0.0`, `ThreadingModel=Both`, `CodeBase=file:///…/artifacts/preview/OneNoteMarkdown.Addin.dll`), `AppID\{E036DE73-…}` (`DllSurrogate=""`), `Classes\OneNoteMarkdown.Preview`, and `…\OneNote\Addins\OneNoteMarkdown.Preview` (`LoadBehavior=0`). The `0.1.0.0` assembly version confirms the B10 version stamping reaches the registration. LoadBehavior=3 remains unset and unauthorized. The release is now permanently registered for this user; launching the desktop host and checking the add-in in the COM Add-ins dialog, then clicking **启用** on the **Markdown 预览** tab, is the remaining manual step.
