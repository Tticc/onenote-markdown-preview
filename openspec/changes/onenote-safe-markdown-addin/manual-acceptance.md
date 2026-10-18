# Consolidated manual acceptance

## Execution status

Plan only: complete product entry, Markdown parser and automatic harness are not yet implemented. Do not instruct the user to run these against the current fixed-payload probes. Prepare one acceptance build and fixture page before requesting the consolidated batch. No requirement to repeat unchanged tests already passed. Phase-one gates remain binding; preparing this plan does not authorize bypass.

## Already observed — do not repeat unchanged

- Fixed inline bold paste and one undo; generated paragraph heading/list/quote import and one undo.
- Inline code Consolas/#eeeeee visual acceptance; italic/strike/ordered/task-symbol fixture and one undo.
- Title/image/cross-container/cross-cell/multi-window/empty-selection read-only rejection samples; search stale-selection evidence; normal/full-page focus samples.
- 200 ms inline and multi-paragraph restoration samples; 500 ms standalone clipboard recovery; synthetic common-content and competing-copy tests; real OneNote common-content snapshot/restoration.
- Official COM manual-load callback/disconnect and independent surrogate tests. Full host exit/startup cost remains unresolved.

## Future single batch requiring user actions

1. Actual parser-to-plugin mixed syntax: one prepared sample with all approved syntax, Unicode, paragraph boundaries and plain surrounding sentinels. Run once; inspect rendering and one native undo. Exclude unproven fenced blocks from supported sample.
2. Rejection matrix: unsupported/malformed Markdown mixed with supported syntax, no selection, title/search, images, multiple containers/cells/windows; same batch runner records expected refusal and no changes. Only integrated behavior, not repeat of metadata probing.
3. Clipboard: text, formatted text and a small image copied from real apps; format once and paste saved content elsewhere. Verify restoration and defined common-content fidelity. Unknown/protection formats skip. Runner generates synthetic fixtures and failures without requiring manual copies per format.
4. Timing/context: during controlled delayed conversion change selection/page, type, copy newer content; expected cancel before commit or preservation of newer clipboard. No automatic undo/retry.
5. 500 ms configuration: one normal integrated sample; no broad latency guarantee from success.
6. Lifecycle/shortcut: enable/disable, foreground other app, repeated/held shortcut, restart preference, manual disconnect/reconnect; ordinary host exit and residue checked by harness. Never force-close user OneNote as a normal test step.
7. Read-only and operation failure: use a designated read-only fixture if available. Exceptions, oversized input, delayed rendering, unavailable clipboard, timeout and late results are primarily injected by automation, not manual host-crash exercises.

## Conditional isolated investigation

Strict fenced-code whitespace is not supported by current evidence. Visual RTF result and one undo passed, but exported literal-space check failed. A future read/copy-back experiment may resolve layout-vs-character discrepancy; do not repeatedly reimport code, substitute NBSP or claim strict support. Need user decision only if scope/semantics change.

## Automated responsibility

Agent owns parser/rendering tests, Unicode/CF_HTML checks, snapshot bytes/pixels, resource limits, fault/copy races, logging privacy, registration cleanup, startup/CPU/memory measurements and evidence categorization. Record passes/failures/skips separately; do not call all work completed from manual screenshots.
