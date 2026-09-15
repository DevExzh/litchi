# Log sections for change 0593

Four paragraphs for the coordinator to merge, one per document, in the style of
each document's newest section. Each is written to stand alone. Their links are
relative to `docs/performance/`, where the four log documents live, not to this
packet directory.

## For `docs/performance/HOTSPOTS.md`

## 0593 — OPC publication reuses the open-time relationship proof

Retained implementation of 0587's SAVE-1 and SAVE-2. The publication plan
reserialized and XML-audited the `.rels` of every part with relationships and
rebuilt and audited `[Content_Types].xml` on every save, then byte-compared the
result and raw-copied the source member, so the audited bytes never reached the
output: 40 of 42 audits and all 41 reserializations on the 132-member hide save
were for unchanged members. `Relationships` now carries the canonical
serialization the preservation provenance captured at open, behind an `Arc`
cleared by every mutation and matched by pointer identity, so an unchanged
collection is copied without being serialized or audited; the content-types
manifest is decided from provenance before it is built; identical payloads
settle on the pointer before `memcmp`; and `atomic::replace_with_impl` stages
the temporary file through a 64 KiB buffer. Per publish on the 132-member
workbook with every member unchanged: `verify_authored` 42 → 0,
`try_to_xml_bytes` 78 → 0, native cycles 992,747 → 342,567 (−65.49%), callgrind
Ir 4,914,478 → 2,400,514 (−51.15%); on a 103-member PPTX, cycles −71.11%. The
real `tabs … hide` save falls from 14.00% to 9.43% of the operation
(−35.82% of the save) with its one regenerated part still audited. Write
syscalls on `save(path)` fall 531 → 14 for the same 654,681 bytes and the same
digest, but `save(path)` is fsync-bound: SAVE-2's modelled 0.8–1.6 ms is
**falsified** here at a measured p50 of −2.33%, inside the host's 4% p50 floor.
Publish-only p50 improves 61.52–68.68% across four fixture/scenario pairs with a
p50 A/A floor of −0.57% to +0.66%; the p99 A/A floor in that window was −38.54%
to +99.40% and the tails are not interpreted. `performance_claim: none`;
`claim_authorized: false`. Remaining in this area: SAVE-3 (PPTX eager save
regenerates every slide, frozen design needed), SAVE-4 (fresh deflate state per
regenerated member), SAVE-5 (C2′ lazy part decode, proposed ADR needed). OLE2
and OOXML remain active; ODF is deferred until that goal completes and iWork is
excluded. [Change and limitations](0593-opc-publication-pristine-members.md);
[retained evidence](results/change-0593/README.md).

## For `docs/performance/GOAL_AUDIT.md`

## 0593 — unchanged-member preservation stops paying for discarded work

Retained implementation against the audit's standing "close ZIP64/CFB and
output-source preservation intersections" row, on its unchanged-member
passthrough clause. Every OOXML save through `PackageWriter` previously paid a
whole-package XML serialization and audit pass for members it then copied
verbatim; the open-time canonical capture now proves the unchanged case in one
pointer comparison, and the content-types manifest is decided from provenance
before it is built. Correctness is established by whole-corpus byte identity,
not by inspection: 336 OOXML fixtures × 4 mutation scenarios = 1,344 published
digests and typed errors, before and after, with an empty diff, including 27
preserved save refusals and 8 preserved open refusals; plus 12 of 12 identical
editor-level rows, among them the two pre-existing `NotCompact` refusals 0587
reported. Measured per publish: native cycles −62.20% to −71.11% on four
fixture/scenario pairs, callgrind Ir −27.48% to −61.37% across ten, publish-only
p50 −61.52% to −68.68% against a p50 A/A floor of ±0.7%. The audit row this does
**not** close: no DOCX or PPTX semantic-editor save was measured through the
ordinary save path, because no example opens a real `.docx` or `.pptx`, edits
through the model and saves — 0587 §4 recorded that gap and it remains open, as
does the absence of an ordinary-save selector in `tools/perf-baseline`. Cold
cache, peak RSS and real-device fsync distributions remain unmeasured, and
`save(path)` remains fsync-bound at ~7.5 ms of an 8.3 ms median on this host.
`performance_claim: none`. OLE2 and OOXML remain active; ODF is deferred until
that goal completes and iWork is excluded.
[Change](0593-opc-publication-pristine-members.md);
[evidence](results/change-0593/README.md).

## For `docs/performance/REPORT.md`

## 0593 — pristine-member proof reuse and a buffered atomic tempfile

`crates/litchi-opc` now carries the canonical `.rels` serialization captured at
open inside each `Relationships`, behind an `Arc` that every mutating method
clears and that publication matches by pointer identity against the preservation
provenance for that exact part. A match means the source member is copied
without reserializing, auditing or deriving its URI; anything else takes the
unchanged serialize-audit-compare route, so every *changed* member is still
audited exactly as before (ADR 0006, record 0528). `[Content_Types].xml` is
decided from provenance before it is built, identical payloads settle on
`std::ptr::eq` before `memcmp`, and `atomic::replace_with_impl` stages its
temporary file through a 64 KiB `AtomicSink` flushed before `sync_all` — the
one API change, replacing `&mut File` with `&mut AtomicSink<'_>` in the closure
of the public `atomic::replace` and `replace_with`; all six in-tree call sites
compile unchanged and caller-supplied streaming sinks are never wrapped, so
`IncompleteOutput.written` still counts bytes the caller's own sink accepted.
Validation passed `litchi-opc` 422 library tests plus its integration suites,
and `litchi-xlsx`, `litchi-docx` and `litchi-pptx` tests and clippy, with
`cargo fmt --all --check` and `cargo doc -p litchi-opc --no-deps` clean; nine
tests were added, including a differential that publishes one package through
both routes and requires identical bytes. One intentional behaviour change is
recorded: the audit of generated XML that is then discarded no longer runs for
unchanged members, so a package whose unchanged `.rels` would exceed the
authored-XML auditor's aggregate attribute ceiling (about 62,500 relationships
on one part) now saves instead of refusing — the published bytes are the source
member's, and the same package already saved without refusing when no mutation
occurred. No latency, RSS, cold-cache or OOM claim follows; SAVE-2's modelled
millisecond-scale gain is reported as falsified on this host. See
[Change 0593](0593-opc-publication-pristine-members.md);
`performance_claim: none`.

## For `docs/performance/ADR_COMPLIANCE.md`

## 0593 — preservation provenance stays planning evidence

ADR 0005's 2026-08-21 amendment holds: the open-time relationship capture is used
exactly as the byte comparison it replaces was used, to choose `Copy` for one
member inside an already-proven preservation plan. It does not touch
`exact_source_authorized`, does not widen who may take whole-archive exact
passthrough, does not authorize a normalizing full-writer fallback, and does not
remove the retained source archive ADR 0005 requires. The proof is conservative
in one direction only — a pointer match proves "unchanged", a mismatch proves
nothing and costs the previous path — and cannot be borrowed across parts or
packages, because each part's provenance owns its own `Arc` allocation; a
collection moved wholesale into another slot therefore falls back to
serialize-and-compare, which a dedicated test asserts. ADR 0006 and record 0528
hold: both the original and the replacement OPC audits remain for every changed
member, and the XML-part audit is untouched. `ReadLimits` and the authored-XML
`Limits` are unchanged; no malformed-input defence over published bytes is
removed; no new `unsafe`, global cache, executor, lock, hidden Rayon pool or
ambient I/O is introduced, and no archive type, raw lock or executor is leaked —
the new public `AtomicSink` is a plain `Write` adapter over a temporary file the
module owns. ADR 0005's output rules hold: atomic replacement keeps its sibling
temporary artifact, permission preservation, symbolic-link refusal, flush,
`sync_all`, rename, parent-directory sync and `Committed` error identity, and
the added staging buffer is flushed before synchronization. One typed behaviour
boundary moves and is recorded rather than hidden: the authored-XML audit of
bytes that are generated and then discarded no longer runs for unchanged
members. Focused validation passed 9 new tests, `litchi-opc` 422 library tests
and its integration suites, and the `litchi-xlsx`, `litchi-docx` and
`litchi-pptx` suites. See [Change 0593](0593-opc-publication-pristine-members.md);
`performance_claim: none`.
