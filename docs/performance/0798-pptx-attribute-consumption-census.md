# 0798 — checked-attribute consumption in PPTX workflows

This diagnostic measures attribute counts and caller consumption before choosing
another iterator prototype. It makes no latency, allocation, RSS, instruction,
or production optimization claim. The 0794 and 0797 rejections remain in force.

## Protocol and scope

The [packet](results/change-0798/README.md) starts at `5f6ec5ca31`. A temporary
hook instruments only `litchi-opc::xml_attributes::CheckedAttributes` on the
calling thread. OOXML common re-exports this helper, including its PPTX notes
callers. Other helper copies, unchecked and lenient iterators, skipped empty
attribute tails, and other threads are outside the census. Element names are
retained losslessly but do not identify Rust call sites.

The probe inherits the deterministic public PPTX fixtures and verification from
0793. Tiny, medium, and large shapes contain 3×4, 12×8, and 100×100 slide/shape
combinations. ASCII and Unicode vendor fixtures each use 12×8. Capture covers
`Package::opened_presentation`; commit covers `commit` after edit staging;
lifecycle covers capture, edit, commit, publication, and serialization. Ingress
and output verification stay outside each census region.

Fifteen plain controls use one sample without warmup. Two instrumented repeats
run the same 15 cases in forward then reverse order, also with one sample and
no warmup. Root runs builds and all 45 captures serially; captures use CPU 12.
Production source is restored before captures. Separate copied binaries retain
the exact plain and instrumented build identities.

The hook records each iterator instance, inherited clone prefix, local successful,
error and end yields, and its drop or live-at-finish state. A separate unchecked
scan counts lexical attributes through the first lexical error. Duplicate names
do not stop this lexical scan. `early_drop` means that no error or end result was
observed; `partial_consumption` means the successful prefix is shorter than the
lexical attribute count. These are different conditions.

Before aggregation, the probe validates instance and lineage identities. Rows
then group by every field except instance and lineage IDs, retaining frequencies
and exact source/name bytes. Qualification requires conservation, no saturated
counters, zero live instances at finish, exact repeat agreement, and semantic,
fixture, source and output parity with plain controls and sealed 0794 results.
An independent reader lexes every retained well-formed tag body and reconstructs
the count distributions.

Source copies, event bookkeeping, diagnostic allocation, and drop-time lexical
scans perturb execution. Recorded elapsed fields are excluded from comparisons.
The census cannot estimate workflow speedup by multiplying counts by 0797 micro
ratios. It adds no real-producer, cold/range, concurrent, or CRUD coverage.

## Validation and results

All 45 reports and samples pass semantic/source/output parity. The two census
repeats agree exactly, including grouped raw source bytes and consumption fields.
Across both repeats, 295,600 iterator instances yield 438,800 attributes. Every
instance reaches `None` and is dropped before finish. There are zero clones,
errors, early drops, partial consumptions, never-advanced instances, live-at-finish
instances, or saturated counters. This describes these fixtures and operation
regions; it does not establish that all production callers consume fully.

Counts below are per operation in one repeat; the second is identical. The last
column combines attribute counts three and above; exact histograms remain in
both independent audit outputs.

| Shape | Operation | Instances | Zero attrs | One attr | Two attrs | Three or more |
|---|---|---:|---:|---:|---:|---:|
| tiny | capture | 429 | 2 | 216 | 173 | 38 |
| tiny | commit | 539 | 112 | 216 | 173 | 38 |
| tiny | lifecycle | 1,016 | 114 | 432 | 394 | 76 |
| medium | capture | 1,023 | 2 | 486 | 488 | 47 |
| medium | commit | 770 | 208 | 261 | 263 | 38 |
| medium | lifecycle | 1,889 | 210 | 747 | 847 | 85 |
| large | capture | 61,327 | 2 | 30,374 | 30,816 | 135 |
| large | commit | 5,250 | 2,416 | 1,177 | 1,619 | 38 |
| large | lifecycle | 67,777 | 2,418 | 31,551 | 33,635 | 173 |
| vendor | capture | 1,119 | 2 | 486 | 488 | 143 |
| vendor | commit | 778 | 192 | 261 | 263 | 62 |
| vendor | lifecycle | 1,993 | 194 | 747 | 847 | 205 |
| unicode-vendor | capture | 1,119 | 2 | 486 | 488 | 143 |
| unicode-vendor | commit | 778 | 192 | 261 | 263 | 62 |
| unicode-vendor | lifecycle | 1,993 | 194 | 747 | 847 | 205 |

Large capture has 30,374 one-attribute instances (49.528%) and 30,816
two-attribute instances (50.249%); together these cover 99.777% of observed
instances. The 0797 replay penalty on two-attribute tags therefore affects a
common observed class, rather than an obscure boundary. This count observation
supports seeking a design that avoids replay while preserving early duplicate
refusal and bounded hostile-input behavior. It does not quantify that design's
benefit. Zero-attribute instances remain visible, especially during commit;
skipped iterator construction is outside the census and cannot be inferred from
these zero counts. Vendor fixtures retain six- and nine-attribute classes.

The isolated instrumented-helper suite passes 22 tests: 12 existing helper
checks and 10 census checks. One canonical-copy filesystem test is deliberately
filtered because only OPC is instrumented. The suite includes parser differential,
error, clone, and bounded-handoff checks plus census lifetime, thread, stale-session,
and saturation checks. This is an isolated helper test, not a fresh full-workspace
regression run. Plain and instrumented probes pass formatting, locked release
build/check, and Clippy with warnings denied.

Retained failed preparation attempts include an invalid optional-feature mapping,
a test helper shadowing error, a needless-borrow Clippy error, and unused inherited
allocation APIs/doc formatting in the plain probe. Corrections precede the final
successful builds; no captured workload is retried or omitted. The independent
reader initially included `metrics.elapsed_ns` in historical identity comparison;
it was corrected to exclude that one timing field as required by the protocol.
The original reader and correction note are retained. Other semantic metrics
remain exact comparisons.

Production remains byte-for-byte unchanged. The next candidate should target
fully consumed short tags without replaying the first attribute, and must pass
fresh correctness, public-workflow, resource, and cross-format gates before
adoption. No baseline or coverage is promoted by this census.

After successful capture and replay, the owned target was removed (990,773,730
logical bytes). Both binary hashes and sizes remain in the cleanup witness;
post-cleanup replay passes. All 9,196 production files, 35 architecture inputs,
unrelated working-tree files, and existing worktrees remain intact. The retained
failures, final gates, raw captures, independent audits, and reviews are sealed
with this packet.
