# Clipboard support matrix

Current user-approved goal: preserve common user content, not every application-private paste capability. Default restoration delay 200 ms, optional 500 ms; owner/sequence checks remain mandatory. Classification does not imply snapshot implementation.

| Representation | Goal / treatment | Evidence today |
| --- | --- | --- |
| Unicode/ANSI/OEM text, locale | Independent copies, restore supported formats together | Synthetic byte roundtrip passed; default 200 ms OneNote inline sample passed |
| HTML Format / Rich Text Format | Preserve each existing representation, not only Unicode fallback | Controlled HTML/RTF/Unicode simultaneous byte roundtrip passed at 200 ms; real app-origin cases pending |
| DIB/DIBV5, PNG | Preserve common image representation, dimensions and pixels | Synthetic PNG bytes and DIB bytes/pixels/dimensions passed; synthesized DIBV5 representation retained when present; complex images pending |
| Bitmap / enhanced metafile / metafile handles | Clone with correct medium-specific ownership or refuse | Not implemented; never GlobalCopy arbitrary handles |
| DataObject / Ole Private Data | OLE transport; investigate enumeration/materialization and reconstruction rather than copying private bytes | Observed in actual Notepad copy; no approved omission yet |
| EnterpriseDataProtectionId | Determine protection semantics; do not silently discard | Observed in Notepad copy; currently refuses |
| History/cloud exclusion flags | Preserve privacy meaning or refuse | Classified, not roundtrip-tested |
| Chromium source token / URL / custom MIME | App-specific source/custom data; classify payload role before deciding omission | Observed; unknown policy means refuses today |
| Files, virtual-file descriptors/content | Outside initial common-content scope; refuse | Classification only |
| Unknown format / medium | Refuse before changing clipboard | Strict format-gate refusal observed |

## Validation order

1. Keep synthetic plain-text baseline as a regression; do not ask the user to repeat inline undo.
2. Validate independent HTML/RTF snapshots with controlled fixtures; never reconstruct only text while claiming rich preservation.
3. Validate image snapshots with controlled small images and medium-specific ownership; no general bitmap decoding in a host callback.
4. Investigate OLE/protection metadata with read-only inspection and source documentation; approve omission only with evidence, not because it is invisible to the user.
5. Test competing-process copying, clipboard contention and restoration failures. Same-process different-owner simulation is not full coverage.

No broad support claim until these results are recorded. Timeout, resource limits and crash recovery are still required for product integration. Existing disposable tests intentionally reset only under explicit permission; they do not represent normal user clipboard backup.

Real OneNote bold-text sample: OLE source enumeration exposes OneNote 2016 Internal, HTML, Unicode/ANSI text, ENHMF and DIB. Read-only independent HGLOBAL snapshots passed for HTML/text/DIB with unchanged clipboard sequence. Native private data and enhanced metafile were not read, and no roundtrip was executed for this source. OLE bookkeeping wrappers visible to Win32 enumeration are absent from the source's OLE representation list; source representation enumeration is preferable to replaying Ole Private Data bytes. Common-content omission of private/native/vector representations requires tested fallback and explicit documented fidelity limitations.

Subsequent authorized actual-source roundtrip PASSED: HTML/Unicode/ANSI/DIB independent bytes restored after temporary text and 200 ms delay, compared representation by representation. OneNote native format omission requires both rich-text and Unicode fallback; vector/bitmap omission requires DIB/DIBV5 fallback. No unknown representation is omitted. Scope does not preserve internal OneNote metadata or enhanced-metafile vector fidelity. OLE source refs/media are released before publication. This is tested with one small sample, not all OneNote copies.

## Documented format semantics

Microsoft Clipboard Formats documentation confirms synthesized text and bitmap representations may appear during enumeration; enumerated formats are not necessarily all explicitly published by the source. CF_DIB/CF_DIBV5 are preferred to device-dependent bitmap storage; ordinary handle copying is not sufficient for all media.

Documented privacy flags: CanIncludeInClipboardHistory and CanUploadToCloudClipboard contain a DWORD (zero prohibits, one requests). ExcludeClipboardContentFromMonitorProcessing presence with any data excludes history and cloud processing. Preserve these semantics; do not infer that an invisible format is disposable. EnterpriseDataProtectionId is not specified by that page and remains unvalidated.

Source: https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats

## Executed common-content fixtures

Test-CommonClipboard.ps1 published authorized disposable fixtures and verified independent snapshots, temporary replacement and 200 ms restoration. All supported snapshot representation bytes matched. The DIB case skips synthesized CF_BITMAP only if DIB/DIBV5 exists, then verifies restored 2x2 pixels through a separate WinForms image read. This establishes the fixed fixture, not all color spaces/alpha/palettes. A separate child process copied different text; restoration guard rejected old ownership/sequence and verified newer text retained. Final clipboard contains the generated small DIB image, as authorized. No OneNote access or original application-private data preservation is claimed.
