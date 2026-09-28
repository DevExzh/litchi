# Next non-iWork coverage step

The independent corpus audit recommends fresh public OOXML ordinary-save
measurements on three tracked artifacts after this provider/scaling batch:

| Format | Input | Source SHA-256 |
| --- | --- | --- |
| DOCX | `test-data/ooxml/docx/documentProperties.docx` | `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5` |
| XLSX | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4` |
| PPTX | `test-data/ooxml/pptx/shapes.pptx` | `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` |

`test-data/office-interop/PROVENANCE.md` binds these inputs to edit/resave/readback
history. The DOCX and PPTX rows do not identify a Microsoft producer application;
no such provenance should be inferred. Existing ordinary-save selectors in
`tools/perf-baseline/src/ordinary_save.rs` accept caller-named `--ooxml-file`
inputs and cover lifecycle, edit, atomic publication, and counting publication.
Their actual selected edits differ from some provenance edits, so fresh current
selector admission and independent output oracles remain required before timing.

A following OLE2 corpus batch can use DOC paragraph replacement, XLS numeric
replacement, and PPT slide removal over tracked POI/LibreOffice artifacts.
Do not substitute a shape-text success for PPT: the prior native shape-text
census retained refusals. All public semantic, untouched-stream/member, source,
refusal, partial-sink and reopen checks must be scoped to the actual operation.

This is a queued evidence direction, not completed coverage or a timing claim.
The current 0816 packet remains low-level category-15 provider/scaling evidence.
