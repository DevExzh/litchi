# 0508 ADR review

All 29 accepted numbered ADRs and the README were read in the ongoing program.
The fresh `adr-manifest.json` verifies that every input remains unchanged from
0507. This batch changes the benchmark's default selection, catalog metadata,
coverage contract and evidence. It changes no format implementation.

| ADRs | Disposition |
|---|---|
| 0001, 0002, 0024 | Correctness and preservation retain priority. No public API, dependency, crate ownership or topology change. Boundary checker passes. |
| 0003 | Existing immutable snapshots are read through text-output APIs; no edit, commit, patch or conflict behavior changes. |
| 0004, 0007 | Format-owned semantic plain text remains distinct from rendering and archive serialization. Paragraph, row and slide object semantics are documented. |
| 0005 | Timers and sink costs are explicit. Source/binary/build/host identities and all sample vectors are retained. The baseline makes no causal speedup or whole-process memory-efficiency claim. |
| 0006 | Output/object ceilings and exact semantic/digest checks remain. Synthetic provenance is truthful; unknown security/producer properties remain unknown, including the checked-in RTF watermark exception. |
| 0008 | Existing201 identities are preserved; default additions require actual report/catalog validation. Representative category coverage is not full support certification. Incorrect XML-compaction references are replaced by scope/evidence references. |
| 0009, 0010, 0011, 0023 | Detection, facade physical ownership and ODF family ownership are unchanged. ODF text export calls existing format owners. |
| 0012, 0016, 0027 | BIFF8 formulas, writer locations and XLS anchors are untouched. |
| 0013, 0017–0022 | PPTX notes, producer templates, calculation chains, web settings, table styles, glossary and fonts are untouched. |
| 0014, 0015 | Property ownership and lossless CRUD are untouched. |
| 0025, 0026 | Chart-area transactions and OLE directory bindings are untouched. |
| 0028, 0029 | iWork work is excluded by the user's scope. No iWork implementation changes. |

The independent `source-review.md` details successful output oracles, RTF's
retained pre-reserved sink, ODF's hashing discard sink, timing exclusions, and
remaining producer/provider/failure-path gaps. No exception or proposed ADR
is needed for this harness coverage expansion.
