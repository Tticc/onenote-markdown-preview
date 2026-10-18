# Reference implementation findings

## OneMore — source verified

Inspected repository commit `b0e4a32fd4cb54195c7bf0b05410b7ff17fd1ef9`. Findings describe this commit, not all released versions.

- `OneMore/Commands/Edit/ConvertMarkdownCommand.cs`: SelectionRange → PageEditor.ExtractSelectedContent per Outline → PageReader → OneMoreDig.ConvertMarkdownToHtml → insert `one:HTMLBlock/one:Data` CDATA at anchor → one.Update(page). Reload page and rewrite newly generated heading/todo/code/list/spacing structures, then optionally call Update again. This command is not a clipboard/Ctrl+V conversion path.
- `OneMore/Commands/File/Markdown/OneMoreDig.cs`: Markdig pipeline plus custom rendering; explicitly removes task-list extension to prevent OneNote stripping HTML input checkboxes before native todo postprocessing. HTML import can therefore require semantic cleanup.
- `OneMore/OneNote.cs` Update: OptimizeForSave, text-color stabilization, schema validation, then UpdatePageContent(xml, DateTime.MinValue, xs2013, true) inside retry wrapper. Does not use expected modification timestamp in this method. This is reference evidence, not a concurrency/safety guarantee and not a recipe to copy force/retry into our design.
- `OneMore/Models/PageEditor.cs`: detailed selection extraction, anchor placement, nested hierarchy/table handling and image CallbackID materialization illustrate significant edge cases. Source specifically notes cursor text-run splitting and image data loss risk during moves.
- `OneMore/Helpers/ClipboardProvider.cs`: separate clipboard helper for other features uses STA access, stashes supported text formats and an image, reconstructs a DataObject on restore, and Paste sends Ctrl+V then Task.Delay(200). StashState has TODO for other formats such as files. RestoreState does not show a newer-copy ownership check. This helper is not proof that ConvertMarkdownCommand uses it, nor an arbitrary-format snapshot or paste-completion contract.
- Source does not establish that its two-pass conversion is a single native undo unit on our target environment. Needs testing if considered.

Sources: https://github.com/stevencohn/OneMore/tree/b0e4a32fd4cb54195c7bf0b05410b7ff17fd1ef9 (paths above).

## OneMark — public documentation only

- https://onemark.neux.studio/docs/en/welcome/ documents Enter-triggered rendering, per-line hidden Markdown source, bidirectional editing, F5 rendering and F8 clipboard export.
- https://onemark.neux.studio/docs/en/tray/ documents a tray executable providing the same functionality and explicit connection to OneNote.
- https://onemark.neux.studio/docs/en/faq/ documents F5 scanning continuous plain lines above cursor and stopping at rendered lines/tables/blocks, and refusing the same note open in two windows.
- No implementation source was located in this investigation. These docs do NOT establish whether rendering uses UpdatePageContent, clipboard, or another mechanism. F8 export to clipboard is not evidence for clipboard-based rendering. No decompilation/installation performed.

## Implication for this change

OneMore provides a concrete third candidate: HTMLBlock embedded in OneNote XML, letting OneNote import HTML via COM without system clipboard, rather than mapping all Markdown manually to native OE structures. It still updates Outline scope, requires transformation/postprocessing, and has unverified single-undo/concurrency characteristics for our requirements. Current user choice remains clipboard; do not silently switch architecture. ClipboardProvider's fixed delay/partial stash does not satisfy the stronger preservation/completion requirements and cannot be presented as their solution.
