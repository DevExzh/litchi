# 0818 — preserve untouched DOCX relationship XML

Ordinary DOCX paragraph edits now retain the original main-document relationship
XML when the relationship state is unchanged. The real `documentProperties.docx`
regression failed on base `7cdaba587b` because publication reordered `rId4` and
removed declaration whitespace; the same exact-byte assertion passes after the
fix. A second regression verifies that a newly added external hyperlink is
serialized and present in the reopened relationship graph.

DOCX previously published its rebuilt document Part through `OpcPackage::add_part`.
That general replacement API discards source relationship provenance. The
format-owned publication path now updates the existing Part's payload in place
and swaps in the rebuilt relationship map. This retains the package-level
source relationship XML without adding a relationship-map clone. OPC's existing
binding comparison decides whether that XML remains usable: IDs, types, target
spellings, and target modes must match. Changed graphs still serialize. Missing
main Parts retain the old insertion path.

The change is confined to `litchi-docx` publication and a two-case regression
file. It adds no public API or dependency and changes no OPC authoring contract.
Mutable access still revokes exact whole-source authorization and tracks the
signature graph; the existing rollback guard restores state after failure.
[Independent source review](results/change-0818/final-review.md) covers the
binding, ownership, signature, and rollback boundaries. This is a preservation
repair, with no measured latency, allocation, or throughput claim.

The independent auditor also corrects two mistakes exposed by
[0817](0817-real-file-ordinary-save.md). Generated corpus counters describe
paragraphs, cells, shapes, and logical payload bytes rather than physical ZIP
members. The new reader recomputes those semantic projections separately from
archive inventory. XLSX cell edits intentionally invalidate workbook calculation
properties. The reader accepts only the exact dirty `calcPr` flags and compares
all remaining workbook structure, root attributes, and namespace declarations;
unsupported source `calcPr` markup is refused.

Neither correction relaxes the DOCX relationship preservation requirement.
Preflight against the immutable 0817 artifacts still rejects all three DOCX
relationship-byte/order observations while the other five cases pass. The
[preflight report](results/change-0818/preflight/artifact-audit.json) and
[audit review](results/change-0818/audit-review.md) retain that boundary.

The new regression first compiled and failed on unchanged production at its
exact relationship-byte assertion. The repaired all-feature DOCX test run has
1,938 passes and 32 ignored tests across 66 suite summaries, including doctests
and both new regressions. The failed baseline run is retained separately; it
is not counted among passing tests. All 35 previously read normative inputs
retain their hashes. The three unrelated workspace files remain excluded.

The following artifact export is an untimed development-profile correctness
check using the existing locked standalone harness. Root-workspace tests and
the standalone tool use their existing different lockfiles, whose identities
are recorded without dependency updates. No native benchmark, allocation lane,
process-counter lane, or performance comparison is introduced by this batch.
The broader OLE2/OOXML performance goal remains active; iWork is excluded.

Fresh independent admission passes for all six corpora and thirty policy
outputs. The real DOCX output is 23,535 bytes with SHA-256
`de9e163ac26e170ee3881c7d7e53ac29836efd6caea6f72db7ebd2cd31d6c774`.
Its unchanged relationship member retains all 817 original decoded bytes.
Only `word/document.xml` differs from the source. Compared with 0817's output,
the only decoded-member difference is the repaired relationship XML; the other
five corpus outputs are byte-identical to their 0817 counterparts.

The separate [ZIP preservation replay](results/change-0818/zip-preservation.json)
checks all six default outputs and confirms identical member order and archive
comments. Every untouched member retains its compressed payload and fifteen
metadata fields, including timestamps, flags, methods, versions, attributes,
extra data, comments, CRC, and sizes. Five-policy byte equality extends these
checks to all thirty outputs. Changed-member semantics remain covered by the
[independent XML audit](results/change-0818/admission-0/artifact-audit.json).

Formatting, all-feature/all-target checking, tests, warning-denied Clippy,
warning-denied rustdoc, and the crate-boundary checker all pass. The checker
covers 65 workspace packages and 244 internal dependency declarations with
11 existing debt items. This does not expand the batch's scope to iWork.
The [replayable outcome](results/change-0818/outcome.json) binds commands,
source snapshots, corpus, binary, audit, and quality results. The
[next step](results/change-0818/next-step.md) is the complete current-source
ordinary-save timing matrix, followed by separately admitted larger corpora.

Root verified the captured exporter before removing its owned build and scratch
directories: 8,894 files / 21,390,560,337 logical bytes. Post-cleanup replay passes:
`python3 -B docs/performance/results/change-0818/validate.py --final`.
[Results review](results/change-0818/results-review.md) confirms the admission
and preservation scope. The previous 0817 packet and the three unrelated
workspace files remain unchanged.
