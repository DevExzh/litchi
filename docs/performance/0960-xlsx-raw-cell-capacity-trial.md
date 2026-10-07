# 0960 — XLSX temporary cell capacity trial

Status: rejected by the frozen screening rule after a complete independent second epoch. Production is unchanged.

## Measured motivation

[0959](0959-xlsx-number-ownership-candidate.md) rejects shared numeric storage
and identifies `Parser::finish_cell` as the largest concrete requested-byte
owner in the six 0958 edit/commit captures: 102,826,752 bytes. This is cumulative
allocation-request traffic, including growth requests. It does not establish
physical copies, resident memory, or a latency bottleneck of the same size.
`finish_parse` separately reserves final `Stored` cells; this trial does not
remove that semantic materialization.

The source-backed shared traversal delivers every namespace-resolved event to
its validator/facts observer and then the ordinary parser. `finish_cell` grows
its temporary `Vec<RawCell>` incrementally. The ordinary materializing parser
already reserves its strict-lane vector from `lane::Summary.cells`; the shared
traversal does not use that lane. Existing source facts are completed too late
to pre-size the parser. Worksheet dimensions are extents, not sparse counts.

## Candidate and costs

The isolated trial uses the existing strict body recognizer after the observer
and parser accept the worksheet's `sheetData` start and namespace. An admitted,
bounded body supplies an exact cell-count capacity hint. The hint uses a
fallible reservation; a failed optional reservation keeps the existing per-cell
allocation path. Declined syntax keeps the same parser path. Every event still
passes through the existing observer and parser, with complete validation,
source freshness, finalization and authoritative fallback unchanged.

The recognizer adds a body scan. That cost is part of measured edit/commit and
total latency, not excluded setup. A reduced allocation-request total alone
cannot justify the candidate. Reserving early also changes transient allocation
timing on a later-invalid body; refusal-order and resource tests are required.
There is no dimension heuristic, retained memo, public API, executor, unsafe
code, source-parse removal or iWork change.

Historical [0544](changes/0544-xlsx-event-preflight-rejected.md) rejected an
extra event preflight after valid-boundary and refusal regressions.
[0779](0779-opc-bounded-input-growth.md) rejected a different buffer-growth
strategy despite much lower cumulative requested bytes, because retained
memory/RSS increased. Those results remain constraints on interpretation;
they are not measurements of this candidate.

## Custody and verification plan

The current tracked runtime source matches all 10,932 files of retained 0957
source byte-for-byte. Performance/report evidence is excluded from both source
exports. Retained baseline executables may therefore be reused with their
original build commands and SHA-256 bindings. Candidate source, build products
and large evidence live outside the checkout under
`/home/zhuhe/litchi-performance-0960/`.

All 35 previously read normative inputs, including GOAL, the scenario checklist
and accepted ADRs, were rehashed and are unchanged. The prior turn made progress:
0959's measured rejection and attribution correction were committed as
`a6223afa08`. This batch does not claim the broader GOAL is complete.

The screening plan uses the same three managed serial cases as 0959, with
three alternating ABBA/BAAB blocks, ten warmups and sixty measured native
iterations per process. A separate allocator lane uses three paired repeats,
two warmups and ten measured iterations. Affinity is CPUs 12–15, workers one;
processes are fresh and serial with fail-stop execution and no sample retries.
Screening cannot authorize production adoption. A promising result requires
full six-block confirmation and relevant fallback controls; an unhelpful
result can terminate the candidate without treating allocation savings as
latency evidence.

## ADR constraints

| Requirement | Trial boundary |
| --- | --- |
| 0001/0005 measured work elimination | Existing allocation attribution motivates a trial; extra scan cost, native latency, tails and memory remain decision inputs. |
| 0002/0024 ownership | Private worksheet parser implementation only; no crate coupling or public surface. |
| 0003 immutable source and publication | No early callbacks, partial stores or publication of provisional facts. |
| 0006 validation and limits | Full event validation, namespace checks, bounded admission, existing typed errors and authoritative fallback remain. |
| 0008 verification | Differential semantic/error tests, full XLSX tests and warning-denied checks precede native capture. |
| 0031 execution | No new runtime or concurrency; benchmark execution remains explicit and bounded. |

## Candidate checks so far

The BOM-corrected production candidate passed all 1,404 XLSX library tests.
Independent review caught a draft offset error: content offsets had been
passed to `Entry::locate`, which applies the BOM origin itself. The corrected
code passes raw reader positions to that helper while preserving content
positions in observer spans. A focused BOM test proves hint admission.

An explicit external-input census then read all eight worksheets from the
retained medium and dense-sparse source archives. Their ZIP SHA-256 values
match the benchmark source oracles. Each worksheet admitted exactly one hint;
all eight reservations succeeded, totaling 27,008 cells. Medium contributes
four times 2,304 cells; dense-sparse contributes 16,384, 1,024, 256 and 128.
This is route qualification, not a timing result. The census test is ignored
in the ordinary suite because it requires an explicit external manifest, and
was run separately with that manifest.

Focused tests cover unsupported syntax, observer-before-parser refusal,
malformed tails, duplicate sheetData, event limits and injected optional
reservation failure. The first Clippy run flagged a needless reference;
removing the redundant borrow passed warning-denied library Clippy. Earlier
failed draft checks remain external evidence, rather than being replaced by
the passing result.


## First-epoch interruption

On continuation, the previous tool session `57382` returned `Unknown process
id`. The packet contains 34 child start records and 33 validated children of
36 planned. The unvalidated `native__medium-one-edit__b02__p1` directory has a
start record but no launch or terminal receipt. A start record alone does not
prove a process launched. The lane has no terminal receipt, and the allocator
packet does not exist. Current process inspection finds no capture process
visible in this environment; it does not establish how the prior runner ended.

The first packet is incomplete. For that packet, no reduction, selector decision, timing comparison,
allocator result, rejection on performance grounds or adoption is claimed.
The frozen no-retry plan must not be completed by inserting replacement samples.
Preserve the partial packet and qualify any future fresh experiment separately,
with a new plan and explicit account of this interruption. Full confirmation
and fallback controls remain required before production adoption.

The final source differs from the 1,404-test suite source only by a redundant
borrow removal and an ignored external-fixture census test. The final source
passed seven focused tests, warning-denied library Clippy and scoped rustfmt;
the full suite was not rerun on that exact final hash. The external census was
explicitly executed and passed.

At the interrupted continuation, sandbox permissions permitted checkout and `/tmp` writes, but exposed
`.git` and `/home/zhuhe/litchi-performance-0960/` read-only. Completion of the original external capture
and removal of its owned build caches therefore were unavailable in
that environment. A separate writable scratch epoch is recorded below. Canonical sources, executables and partial evidence are
retained. The bounded [evidence packet](results/change-0960/README.md) records
custody and verification scope. The broader GOAL remains active.

Normal staging was attempted and failed with `Unable to create
'/home/zhuhe/code/litchi/.git/index.lock': Read-only file system`. The batch
remained uncommitted at that point; no push was attempted. External cache cleanup
was pending until the owning directory became writable.


## Independent second epoch and decision

A new plan was frozen before any second-epoch measurement, after independent
review of restart feasibility. Both 10,933-file source trees were rechecked for
membership, modes and SHA-256 without differences. The same retained binaries,
runner, reducer, selector and thresholds were reused with a fresh host receipt.
Outputs and workload working directories moved to
`/tmp/litchi-performance-0960-epoch2/` on tmpfs. Baseline and candidate use the
same environment within this epoch. It is not a cross-epoch comparison.

The original interrupted packet remains entirely excluded: no timing inspection
for selection, pooled samples, repair or replacement. The new plan authorizes
one complete fresh experiment, with no child retries or automatic further epoch.
Its native 36/36 and allocator 18/18 children completed and validated, including
semantic readback, source/output/inverse oracles, sink behavior and resource
checks. Native and allocator processes ran separately and serially on CPUs
12–15, with one worker. There are 2,160 measured native iterations and 180
allocator rows. The [second-epoch packet](results/change-0960/epoch2/README.md)
retains the plan, terminal receipts, comparison, selector and external custody.

Ratios below are candidate/baseline, summarized as the median of six native
paired legs or three allocator pairs. Values are descriptive screening results,
without confidence intervals or a statistically qualified effect claim.
Allocation calls and requested bytes are per-child means over ten drained-core
rows; region peaks are the maximum over those rows. They are distinct from
whole-process RSS and do not measure physical bytes copied.

| Metric | Medium 1% | Dense-sparse 1% | Medium one edit |
| --- | ---: | ---: | ---: |
| Edit/commit p50 | 1.0218 | 1.0346 | 1.0215 |
| Edit/commit p95 | 1.0258 | 1.0337 | 1.0204 |
| Total p50 | 1.0347 | 1.0214 | 1.0160 |
| Total p95 | 1.0344 | 1.0377 | 1.0217 |
| Native peak RSS | 0.9905 | 0.9992 | 1.0028 |
| Allocator requested bytes | 0.7360 | 0.8561 | 0.7492 |
| Allocator allocation calls | 0.9996 | 0.9999 | 0.9996 |
| Allocator region peak live bytes | 1.0000 | 1.0000 | 1.0000 |
| Allocator-process peak RSS | 0.9318 | 1.0065 | 1.0111 |

The frozen root selector returns `reject_screen`: no native primary latency,
native RSS, or allocator region-peak median reaches the qualifying ratio of
0.95. Allocation-request traffic alone was explicitly insufficient. The
allocator-process RSS measurement is kept separate from native RSS and the
live-byte observer peak; it does not satisfy that predeclared signal rule.
No primary median exceeds the 1.05 stop threshold, but all individual
phase/tail/pair review flags remain in the comparison. Examples include medium
publication p50/p95 ratios of 1.0523, one-edit ingress p95 of 1.0657 and
one-edit ingress p99 of 1.1703. These descriptive flags are not hidden by the
primary-median summary.

The reducer's `followup_required_for_confirmation` field indicates that
screening alone cannot qualify an effect. It is not an instruction to override
the separately frozen root selector. Because the candidate fails the selection
rule, it does not advance to full confirmation or fallback performance trials.
The extra recognizer scan does not establish a useful end-to-end or peak-memory
benefit in these scenarios. Keep the isolated patch as rejected experiment
provenance; do not introduce it into production. Future work should target
measured semantic materialization and merge ownership while preserving parser
validation, rather than assuming lower requested bytes means lower peak memory.


Independent review rechecked custody and oracles for all 54 children and
recomputed all 63 native/allocator metric summaries directly from the raw
reports without importing the reducer's arithmetic. It found zero mismatches
and independently reproduced `reject_screen`. Native p50/p95/p99 phase results,
RSS, paired ranges and individual >1.05 flags remain available in the bounded
comparison and audit.

No new build cache or disposable bytecode was created for the second epoch.
Its raw packets, plans, source checks and replay helpers are retained as
measurement evidence. Cleanup of the earlier external build caches remained
pending because those directories were read-only. A repeated staging attempt
also failed on the read-only Git index; the reviewed batch was still uncommitted
at the end of that continuation.
No push was attempted.


## Next measured hypothesis

A separate source review identifies a narrower work-elimination candidate:
`MultiSourceEdit::apply_batch` in `cell_values/source.rs` reserves all remaining
batch slots before consuming requests, including the single-item array passed
by `set`. The corrected 0959 owner ranking attributes 28,570,176 requested
bytes across 1,632 calls to this owner. This is allocation-request attribution,
not proof of latency, RSS or physical-copy cost.

A future isolated trial can size temporary staging from the iterator hint,
with the existing bound and fallible growth preserved. Iterator hints are
untrusted optimization inputs: underreported hints must not permit infallible
growth or bypass batch limits. Unknown hints, empty input, over-limit batches,
atomic rollback, cancellation and refusal ordering need explicit controls.
Measure ordinary and managed single-edit, repeated 1% edits, dense and full
batch cases before adopting anything. This batch implements no such change.


## Writable continuation closure

The next continuation classified the previous turn as progress: epoch2 had
completed and independently rejected the candidate, while staging and cache
cleanup were blocked. With writable Git and external storage restored, the
coordinator replayed the raw audit: all 54 children passed, zero mismatches,
and the output was byte-identical to the retained audit. A separate reviewer
replayed reduction and selection, verified all 511 external manifest entries,
679 absolute artifact descriptors and 35 normative input hashes. The selector
still returns `reject_screen`; no workload was rerun.

The two obsolete 0960 build caches were removed after process/open-file and
retained-binary checks, releasing 2,806,865,920 allocated disk bytes. Sources,
raw evidence, build commands and hash-verified executables remain external.
Temporary audit replay files were removed. Closure receipts are retained under
`results/change-0960/review/closure-{audit,review,cleanup}-v1.json`.

Git staging now succeeds. The bounded documentation/evidence batch is ready
for commit; unrelated work is excluded. The repository size and binary-path
checks pass. Whitespace checks pass outside the preserved candidate patch;
its three blank unified-diff context lines are intentionally unchanged to
retain the recorded patch hash. No production code changed, so no new runtime
suite is claimed. The broader performance goal remains active.
