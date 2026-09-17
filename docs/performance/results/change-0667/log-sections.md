# Log sections for change 0667

These are the four coordinator-ready paragraphs for the shared rollup files.
The branch itself does not edit those files.

## For `HOTSPOTS.md`

## 0667 — the XLSX value-only editor closes the shared-string D2 gap with stable-index admission

Record: [0667](../../0667-xlsx-value-editor-shared-strings.md). The follow-on
to 0657 implements 0602 D2: one internal shared-string relationship is retained
and its table is materialized once, only when a worksheet references `t="s"`.
Adjacent numeric and inline-string edits preserve the source table byte-for-byte
and shared cells read back as text; a mutation at a shared-string target refuses
before staging because the editor does not author or renumber entries. The
relationship cardinality, content type, internal-target, no-outbound-relation,
index, malformed-table, and 64 MiB retained-part gates are covered by focused
witnesses. No speedup is claimed; the producer harness evidence and its
after-only timing are retained in the packet with the unstable dense-open case
excluded from the p95/p50 floor.

## For `GOAL_AUDIT.md`

## 0667 — shared-string reads now preserve the workbook closure without moving the safety boundary

Record: [0667](../../0667-xlsx-value-editor-shared-strings.md). This work
removes the unnecessary whole-workbook refusal caused by a valid shared-string
relationship while keeping the editor value-only: the source table stays
unchanged, indexed cells resolve through one lazy cached table, and any edit
that could add, remove, or renumber an index is refused before a partial
candidate exists. The 66,935-entry witness, malformed/unreferenced-table test,
index error test, and exact publication test cover bounded resources and
lossless preservation. Remaining pivot/table/query dependencies and shared
table authoring remain deferred.

## For `REPORT.md`

## 0667 — shared strings admitted for read and preserved on value-only edits

Record: [0667](../../0667-xlsx-value-editor-shared-strings.md). The change
retains one valid internal `sharedStrings.xml` part, resolves `t="s"` indexes
through a lazy `OnceLock` table, and publishes an adjacent numeric edit while
transferring the shared part unchanged. Medium and dense producer-shaped
evidence contain 40% shared cells and 64 unique entries; both read variants
record the shared-string fact as admitted. The 12-case release run is
after-only evidence with no baseline or registered performance claim, and its
stability floor excludes the dense open case.

## For `ADR_COMPLIANCE.md`

## 0667 — D2 applies preservation-by-default and bounded retained state

Record: [0667](../../0667-xlsx-value-editor-shared-strings.md). The record
implements 0602 D2 under 0652 decision 7's dependency rule and the 0651 queue:
shared-string bytes and relationship identity are captured as source state,
the table is lazy and charged to the existing aggregate bound, and reduced
readback keeps the original index identity. ECMA-376 Part 1 §12.3.15's single
internal part rule and §§18.3.1.96, 18.4, 18.4.9 and 18.18.11's index/type/count
rules are cited in the record. Mutation paths refuse before staging whenever a
shared target could change the table; no public API, OPC contract, or accepted
ADR wording changes.
