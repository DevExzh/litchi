# 0968 — staged CFB commit census and writer-proof handoff

**Status: completed diagnostic; no production candidate or adoption.** The
corrected v6 48-child census and independent raw audit pass. This record
selects a source review and route census for a possible CFB writer-proof
handoff. It contains no timing, allocator, or speedup claim. Production remains
unchanged. [Packet cleanup](results/change-0968/receipts/build-cache-cleanup-v1.json)
is recorded in the receipt (8,359 bytes, SHA-256
`0c45dde744f4f0488c0e737a6bdd46e1321db8f800b8d506ce63a0b956b8d5ee`) and
freed 8,896,917,504 allocated bytes while preserving the exact executable,
source, and raw evidence.

## Opportunity and source boundary

The highest-ranked current opportunity is an operation-local proof after a
successful source-layout `ReusePlan`. A staged CFB commit renders
`self.package`, reopens the emitted bytes, and calls `Package::capture`, which
opens every stream into new payload allocations before
`reuse_stream_allocations` can recover equal staged `Arc`s. A future handoff
could skip that second payload materialization only when a proof is tied to the
exact model, validated plan, emitted vector, limits, and sink. Final physical
open, metadata projection, limits, protection, and object discovery would
remain required.

The current source revision is
`6493e26155ee05e554f1371dd2fad707cd4bca0e`. The relevant path is
`ObjectEditor::commit_candidate_with_rendered` in
`crates/litchi-ole-common/src/object/editor.rs:977-1000`, through
`render_copy_through` or `render_with_layout` in
`crates/litchi-ole-common/src/object/codec.rs`, the writer's
`ReusePlan::validate`/`emit`, the final `OleFile::open` and `codec::open`,
`Package::capture` and `reuse_stream_allocations`, and finally
`discovery::from_package`. The census observer is external and feature-gated;
the main checkout has no runtime change.

Three routes remain separate:

| route | census meaning | handoff disposition |
| --- | --- | --- |
| Source-bound copy-through | `render_copy_through` composes a source-backed output and reopens/checks it. | No `ReusePlan` proof is implied. Keep the current capture path until a separate overlay receipt exists. |
| Source-layout writer with successful `ReusePlan` | `render_with_layout` reports a successful `last_sector_layout`; that route report implies the reuse plan was validated and emitted. | Initial handoff target, subject to the proof and metadata gates below. |
| Fallback/from-scratch writer | A declined plan, rewrite policy, or validation failure uses ordinary serialization. | No handoff proof exists; retain full capture. |

This distinction preserves the rejected 0921 source-bound recapture result.
0921 compared staged streams against a source-bound rendered view while
retaining complete recapture, metadata checks, final reopen, and discovery. It
is not evidence that the prospective writer-proof handoff is safe or common.
The new handoff must project final directory metadata without stream-payload
copies, then perform the ordinary final physical reopen and discovery before
publishing the package. A blind `self.package` assignment would lose or stale
SID, kind, class ID, links, state bits, FILETIMEs, start-sector, stream-size,
or MiniFAT metadata, especially for new streams and changed stream sizes.

## Retained attribution

The canonical v2 attribution is a read-only reduction of retained 0942
frame-pointer samples. It is bound to the current six source files and makes
no new build, workload, latency, allocation, or adoption claim. The
`FloatingPictures/e32` commit-root denominator is 1,477 samples and
6,302,069,415 period units. Inclusive ancestry assigns:

| inclusive target | samples | period share |
| --- | ---: | ---: |
| `Package::capture` | 370 | 25.054588% |
| nested `OleFile::open_stream` | 205 | 13.878805% |
| `OleWriter::write_to` | 558 | 37.789601% |

The `open_stream` rows are nested within `Package::capture`; inclusive rows
overlap and cannot be added. The nearest-frame exclusive buckets conserve the
denominator and assign `cfb_writer_write_to` 37.789601%, while missing named
frames do not prove zero CPU. These observations rank a diagnostic question;
they do not quantify removable work or predict a speedup. The historical 0919
requested-byte figures for staged recapture are allocation-owner evidence under
a different method and are not a current bound.

The draft packet retains the canonical attribution JSON
[`results/change-0968/analysis/v2/cfb-commit-attribution-v1.json`](results/change-0968/analysis/v2/cfb-commit-attribution-v1.json),
SHA-256
`e6f937fee96738866b0a19e8d0356bbe97397b346f21ecd02f5a397491863d49`, and
its byte-matched Markdown reduction
[`results/change-0968/analysis/v2/cfb-commit-attribution-v1.md`](results/change-0968/analysis/v2/cfb-commit-attribution-v1.md),
SHA-256
`1e612750df98818e2a94e7bd4a5e9b7294a06d11ed519add06f3ba60bb1f3c7e`.
The independent CPU review is
[`results/change-0968/review/independent-review-0968-v1.json`](results/change-0968/review/independent-review-0968-v1.json),
SHA-256
`71b475363d4b62558115cfe3a152508e2c6fe724dd354836c08a18576a1553ed`, and
the utility approval is
[`results/change-0968/review/perf-decision-independent-approval-v1.json`](results/change-0968/review/perf-decision-independent-approval-v1.json),
SHA-256
`a1af06aa16eaa79bba2a6a683567855aec64d662a0d961c491dfd0229ffb3b98`.
The census descriptor
[`results/change-0968/provenance/pending-census-descriptor-0968-v1.json`](results/change-0968/provenance/pending-census-descriptor-0968-v1.json)
records the external implementation and capture roots, SHA-256
`12892125e9013b23a382485e08a8a2a1080c020e48a2aeb5564d405be4fd5a94`; it is
not a census result.

The maintained [`tools/perf_decision.py`](../../tools/perf_decision.py) utility is
the approved facet-decision aggregator for this batch. It distinguishes missing
or incomplete evidence from a complete facet failure. Its independent review
passed the 2,048-case facet-precedence matrix and 14 focused tests. A replay of
0967 demonstrates that it treats complete failed gates as `reject_candidate`
rather than the old reducer's `pending_incomplete`; that replay does not
qualify this 0968 census, and no threshold or performance claim is carried
forward from it.

## Census design

The external diagnostic uses the current source baseline and one
feature-gated observer. It covers the existing 24-cell matrix—`NoHeadFoot` and
`FloatingPictures`, 1/2/8/32 edits, and zero/default/release retention—with
two fresh serial children per cell: 48 children total. Each child uses zero
warmups and one counter sample. Fixture preparation, correctness routes,
output hashing, JSON serialization, and independent oracles are outside the
commit census. `timing_claim` is false; the observer has no clock or allocator
instrumentation.

The main render-phase gate covers only the main render expression inside
`commit_candidate_with_rendered` and ends before physical reopen. The surrounding
commit scope continues through physical reopen, codec admission, package
capture, `reuse_stream_allocations`, and discovery serialization. The minimal
packet schema records bounded `commit`, `overlay`, `layout`, and `capture` route
observations. A successful `last_sector_layout` report is the layout route
evidence that the reuse plan validated and emitted, so the census does not
claim separate validation or emission counters. Capture records the successful
stream count and payload bytes. Aggregate `commit.errors` includes final-open,
codec, stream-reuse, and discovery failures; these are failure accounting, not
separate timing endpoints. Capture payload bytes mean only the exact stream
bytes read by `Package::capture` in the active commit scope; they are a
diagnostic recapture workload signal, not allocator traffic.

The independent reducer must reject missing or extra fields, saturated or
negative counters, duplicate/missing cells, unequal repetitions, output
mismatches, route-conservation failures, and any false existing correctness
gate. It requires the two repetitions in each cell to have byte-identical route
and counter vectors, output bytes and SHA-256, and correctness results. It may
assert observed conservation, but must not infer that every overlay decline
reaches layout or that capture bytes equal total allocator traffic.

## Census result, custody, and future proof requirements

The first 48-child epoch completed with all children exiting 0 and all 20
correctness gates true, but its protocol admission is invalid. The checker
incorrectly required one staged CFB commit per sample; the fixed matrix has
1/2/8/32-edit cells with 12 children per edit count, so 36 reports were
rejected by that checker. The epoch and its rejection evidence are preserved,
but none of its route counts or payload bytes qualify as census results. The
protocol and helpers were corrected as v3; the fresh v6 epoch used the new
corrected approval and an explicit root GO. That correction does not change the
source or executable.

The corrected v6 reduction is
[`results/change-0968/analysis/census-summary-v1.json`](results/change-0968/analysis/census-summary-v1.json),
50,259 bytes, SHA-256
`48fc947ac31ce136ea45f3c2b55143e00543e65ad6accd02f054424ab4304035`.
It reports `status: pass`, zero reducer errors, 48 children, 24 cells, and
two equal repetitions per cell. Across both repetitions and all edit counts,
the fixed diagnostic records 516 successful commits, 516 overlay declines,
516 successful layout-reuse routes, 516 successful captures, 4,902 captured
stream reads, and 130,646,838 captured payload bytes. There are no commit,
overlay, layout, or capture errors; these are route and workload counters, not
timing or allocator measurements.

At 32 edits, the summary's representative child is identical across zero,
default, and release retention and across both repetitions:

| fixture | commits | captured streams | captured payload bytes |
| --- | ---: | ---: | ---: |
| `NoHeadFoot` | 32 | 192 | 1,354,864 |
| `FloatingPictures` | 32 | 416 | 16,121,304 |

For one retention cell across both repetitions, those rows are 64 commits,
384 streams, and 2,709,728 bytes for `NoHeadFoot`, and 64 commits, 832
streams, and 32,242,608 bytes for `FloatingPictures`. These values describe
the current capture route and do not estimate removable work.

The independent raw audit is
[`results/change-0968/review/independent-raw-census-audit-v2-final-v1.json`](results/change-0968/review/independent-raw-census-audit-v2-final-v1.json),
27,912 bytes, SHA-256
`1fb1c5b5aaad32d8b55770d3e06917927b7cc6e4f7b7100ae02b2d58f6bb0333`.
It passes custody, fixed schedule, packet/schema conservation, all 20
correctness gates, all 24 repetition comparisons, and no-retry checks. Its
decision is `diagnostic_only_no_adoption`; no candidate is enabled.

The restored protocol approval is
[`results/change-0968/review/protocol-approval-v2.json`](results/change-0968/review/protocol-approval-v2.json),
SHA-256
`2d8468789f99de61042d74183ce8e680e1326d4d213e43f13995ae7d54150225`.
After GO, a reviewer overwrote that receipt with amended v3 bytes. The
[custody incident receipt](results/change-0968/review/protocol-approval-custody-incident-v1.json),
SHA-256
`ed51390508d8795432637c954d2eeb58a9d9536dc8ebd865be646725056cd949`, records
that the exact original v2 bytes were restored and root-verified; the amended
v3 bytes remain preserved separately at SHA-256
`dbd51169792a19350b2c308b6eb0d21746a72be08645bfab3ac9161a618bc70c`.
The incident changed no source, executable, plan, helper, raw capture, build,
workload, or production artifact and does not invalidate the v6 capture.
Cleanup is recorded by
[the build-cache receipt](results/change-0968/receipts/build-cache-cleanup-v1.json)
(8,359 bytes, SHA-256
`0c45dde744f4f0488c0e737a6bdd46e1321db8f800b8d506ce63a0b956b8d5ee`), which
records 8,896,917,504 freed allocated bytes while preserving the exact
executable, source, and raw evidence.

Only a later source/design review may implement a typed receipt. It must bind
the exact operation-local package model, source-layout identity, limits, stream
allocations, successful `ReusePlan::validate`, successful `emit`, and the
fresh output vector. It must validate the complete final directory projection
for root, storages, and streams, retain current error ordering and failure
atomicity, and run the ordinary final `OleFile::open`, `codec::open`, limits,
protection, and `discovery::from_package` checks. Unknown routes, proof
mismatch, sink failure, metadata failure, or limit failure must retain the
current full capture path or return the existing error.

No source optimization, late-CFB deferral, public API change, global cache, or
source-bound 0921 mechanism follows from this diagnostic. Any implementation would
need a fresh correctness review and a separately authorized native/allocator
comparison after the census establishes that the successful `ReusePlan` route
is common and its recapture payload is material.

## ADR disposition

ADR 0001 and 0005 keep correctness, preservation, and bounded resource
evidence ahead of the historical allocation-owner signal. ADR 0003 and 0006
require final output identity, metadata, reopen, error ordering, and semantic
controls to remain intact. ADR 0008 keeps the invalid epoch, corrected census,
and custody incident visible and prevents promotion without a candidate-specific
bounded qualification.
