# XLSX x14ac fallback probe

This is a standalone public API probe for the change-0516 fallback decision.
It is evidence-only: it does not modify the performance harness or any
production crate, and it uses the same independent Cargo workspace, frozen
lockfile, shared allocation observer, corpus shapes, update coordinates, and
timer boundaries as the unchanged `guard-probe`.

The probe creates one deterministic two-sheet numeric workbook for each shape.
Before the initial commit it applies the public
`WorksheetEdit::defaults().height(15).descent(0.2)` operation to both sheets.
The generated source is then inspected through the public
`litchi_opc::{OpcPackage, Part, PackURI}` APIs (`OpcPackage::from_bytes`,
`get_part`, and `Part::blob`).  Every worksheet part must
contain the x14ac namespace URI, the markup-compatibility namespace URI, an
`mc:Ignorable` token for the x14ac prefix, and the serialized
`dyDescent="0.2"` attribute.  The public `Worksheet::defaults()` view must
read back the same descent on both sheets before any timing starts.

Only `warm-changed-one-cell` and `warm-changed-one-percent` are measured.
Each iteration opens the frozen source, enumerates the complete public cell
view for every sheet touched by the update, and prepares the edit before the
`Edit::commit` timer.  Serialization, source-marker checks, default-descent
readback, complete numeric cell validation (including untouched cells), and
changed-cell validation happen after the timer.  The preceding commit remains
retained until replacement after the next measured region, matching the plain
guard's lifetime.  Allocation samples use the unchanged shared observer and
are unavailable in the normal binary.

The report keeps chronological elapsed and allocation vectors with contiguous
zero-based sample indices, percentile statistics, allocator identity, corpus
variant metadata, source-marker counts, and the initial/output descent
readback proof.  `corpus_sha256` binds every shape's serialized source.  The
binary intentionally keeps the plain guard's package name because the parent
driver copies each role's executable into its own evidence namespace.

No build, benchmark, or commit is part of this probe-authoring step.  The
parent driver owns those actions and records their receipts.
