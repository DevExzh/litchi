# Change 0467: scan eager XLSX cell attributes once

`performance_claim: scoped owned dense XLSX one-percent commit/save latency`

`claim_authorized: true; primary dense XLSX latency only`

This batch evaluates a private XLSX parser change motivated by the 0466 dense
commit/save profile. `Parser::start_cell` previously scanned each cell's XML
attributes separately for `r`, `s`, `cm`, `vm`, and `t`. The candidate performs
one checked scan, decodes the coordinate at encounter time, and retains the
other four modeled attributes in fixed borrowed slots until their existing
validation points. It is intended to remove repeated parser work while
preserving malformed-input rejection, duplicate checks, entity normalization,
qualified-extension handling, numeric bounds, and the existing error order.

The production candidate is committed as
`87733cf3b86c5500ee7aee4cf6a4cb0f3cabf6a7` (`perf(xlsx): scan eager cell
attributes once`), based on control revision
`33243bc2e6bd63ab056eb45493b0717818501f4b`. Both revisions have clean build
bindings and completed descriptive comparisons. A further 500-sample
qualification uses one shared build path and passes all four primary latency
statistics. The claim is limited to the dense one-percent commit/save row.

## Mechanism

The old path invoked the shared checked attribute helper five times. Every
invocation traversed the full start-tag list and performed quick-xml duplicate
and malformed-input checks. The candidate's private `CellAttributeView` does
one `with_checks(true)` traversal. `r` is decoded and owned immediately;
`s`, `cm`, `vm`, and `t` remain raw `Attribute` values borrowed from the
current start tag. The numeric fields are then decoded and parsed by a small
shared helper in the same old order, and `t` is decoded when the
`PendingCell` is built.

The view has a fixed number of fields and does not collect unknown attributes.
It therefore removes repeated per-cell iterator setup and scanning without
creating an unbounded index or changing the source-backed preservation model.
The measured observations and their limitations are recorded below.

## Semantic boundary

The scan still consumes the complete attribute list with duplicate checking
enabled. Ignored and qualified attributes are not projected into the typed
cell state, but malformed syntax and duplicate names remain errors. Entity and
XML 1.0 normalization use the same quick-xml operation as the old helper.

The candidate completes the scan before parsing `r`, just as the old first
attribute lookup completed its checked iteration before calling `parse_a1`.
It decodes `r` during the scan so its malformed entity error retains priority
over later errors. Style parsing, metadata parsing and type decoding remain
after coordinate validation in their previous order. Existing style and
metadata limits, inferred-column behavior, formula handling, and publication
validation are unchanged.

The focused candidate tests cover entity decoding, valid duplicate rejection
for all five fields, ignored and qualified duplicate names, invalid-coordinate
and invalid-style precedence, metadata bounds, and a prefixed-field success
case. The retained `cargo test --locked -p litchi-xlsx --all-features` receipt
passes all 1,242 tests across 59 targets: 935 library tests plus the
integration and doctest targets in the same log. XLSX formatting, the
all-feature workspace check, XLSX `-D warnings` Clippy with `--lib --no-deps`,
rustdoc with warnings denied, and the crate-boundary checker also pass.
Workspace-wide formatting retains one unchanged excluded iWork difference at
`crates/litchi-keynote/src/document.rs:592`; `validation/format-scope.json`
records why it is excluded. Both separate-path and shared-path release builds
pass with the same Rust 1.98.1, frame-pointer, unwind-table and debug-level-1
settings. The boundary audit covers 64 packages and 238 internal dependency
declarations, with 11 existing explicit debt items.

## Evidence basis

0466 measured the ordinary dense `xlsx_one_percent_commit_save` path on two
256-by-256 sheets (131,072 cells, 1,311 updates) using source SHA
`5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714`. Its
frame-pointer profile contained 15,483 samples and identified inclusive
`unqualified_attribute_value` context at 11,458,403,980 weighted event
periods within exact commit ancestors. Heaptrack reported 56,349,806
allocation calls and identified repeated quick-xml checked-attribute growth as
a concrete lead. Those observations include setup, warmups, expected-output
construction, verification and teardown where applicable; they are not an
operation-local before measurement for this code change.

The 0467 protocol freezes clean, matched release debug-level-1/frame-pointer
builds for A1/B1/B2/A2 normal lanes, with tiny, medium and dense-wide shapes,
100 samples and five warmups per regular lane. These are descriptive results,
below the registry's 500-sample minimum. The complete mean and tail results,
paired ratios and same-role drift are retained in `analysis.json`.

| Scenario | Shape | A1 p50 ms | B1 p50 ms | B2 p50 ms | A2 p50 ms |
|---|---|---:|---:|---:|---:|
| One cell commit/save | Tiny | 0.248871 | 0.233306 | 0.232020 | 0.248561 |
| One percent commit/save | Tiny | 0.457072 | 0.425026 | 0.423211 | 0.455612 |
| One cell commit/save | Medium | 2.879741 | 2.633446 | 2.611868 | 2.877936 |
| One percent commit/save | Medium | 11.614700 | 10.439234 | 10.533683 | 11.413336 |
| One cell commit/save | Dense-wide | 196.195093 | 181.919869 | 179.583598 | 195.862543 |
| One percent commit/save | Dense-wide | 395.150801 | 362.532066 | 372.430696 | 394.173261 |

Process maximum RSS for these four lanes is respectively 120,532, 111,072,
119,992 and 110,344 KiB. The closing A2-to-B2 pairing increases 8.74%, while
both same-role RSS comparisons drift by about 8%. These process-wide
observations do not establish a reduced memory footprint.

The original full 201-row, 15-sample guards retain 67 canonical policy
regressions, including 64 latency regressions. The separate five-percent
latency review flags 111 statistic cells across 46 scenario/corpus rows.
Full-process RSS rises from 146,372 to 154,816 KiB (5.77%). All individual
comparisons remain available in `full-guard.json`; no aggregate hides them.

A supplemental 29-row, 100-sample ABBA follows every scenario whose mean or
median rose more than five percent with a baseline above 0.1 ms. It retains
all CFB/DOC shapes and the flagged medium XLSX source-backed shapes.
The source-backed XLSX median regressions did not repeat. DOC payload-heavy
write medians did: 3.8310 / 4.2590 / 4.3282 / 3.8325 ms in ABBA order.
One CFB concurrent-read source counter also varies between legs; the canonical
summary rejects that source-identity mismatch. Raw vectors and the rejection
are preserved, and those rows cannot support a strict latency claim.

## Build-path follow-up and allocation observations

The earlier reused executable, built from the same control source in the root
checkout, measured DOC payload-heavy at 4.275977 ms. Its separately rebuilt
clean control measured 3.889006 ms. This suggested a build/layout effect and
motivated a new comparison using the same absolute checkout path for both
builds. It does not identify the exact cause of the earlier difference.

Both fixed-path builds use `/tmp/litchi-goal-0467/shared-tree`. Before every
capture, the driver checks that tree is clean, switches it to the exact source
revision of the immutable binary, and checks it again. The executable's
compiled-in manifest path therefore reports the matching clean revision.
The initial fixed-path full guard measures DOC payload-heavy at
4.268618 / 4.290148 ms and dense XLSX one-percent commit/save at
407.163791 / 371.279934 ms. The frozen follow-up retains three DOC guard rows
alongside the primary XLSX row in four fresh 500-sample processes.

The fixed-path qualification completed all four lanes with 500 samples and
five warmups per row. Dense XLSX p50 in A1/B1/B2/A2 order is
409.517775 / 373.989532 / 365.149511 / 396.278436 ms. The two paired
reductions are 8.68% and 7.86% for p50, 8.66% and 7.47% for mean,
8.68% and 6.07% for p95, and 8.82% and 5.86% for p99. All four statistics
pass the unchanged `latency-abba-v1` sample and drift policy; the largest
same-role primary drift is 3.40%.

All three DOC guard rows stay within the five-percent pairing and drift review
threshold. Payload-heavy p50 is 5.899478 / 5.868253 / 5.864874 / 5.902334 ms.
These values use the matched 500-sample workload and are not compared across
the different full-matrix configuration. Fixed normal process RSS is
115,648 / 115,300 / 115,552 / 115,796 KiB; paired differences are under 0.31%
and same-role drift under 0.22%. No memory reduction is claimed.

The fixed-path full guard still retains 53 canonical latency flags and 98
five-percent statistic flags across 41 rows. Only one mean or median flag has
a baseline above 0.1 ms: `opc_source_concurrent_same_part` on
`few-large-compressible`, whose mean rises 6.80% from 0.416072 to 0.444374 ms;
its median is not flagged. These 15-sample tail, short-operation and concurrency
observations remain unresolved. They do not support a blanket regression-free
claim. Full process RSS is 166,824 / 149,452 KiB; its different matrix and the
earlier RSS variability preclude a memory-benefit claim.
`acceptance-review.json` records the scoped retention decision and all inputs.

Separate five-sample/one-warmup Heaptrack exports from the original clean
build pair record 56,349,804 / 44,815,468 allocation calls and
26,103,886 / 14,569,547 temporary allocations. Both rounded peak-heap values
are `104.38M`. These are whole-process observations, including generation,
expected-output construction, setup, warmup, verification and teardown.
They are not counts per commit, exact peak-byte reductions, or normal RSS.
`heap-summary.json` recomputes the totals from authenticated text exports.

## Limitations and acceptance state

This candidate changes only the cell-start attribute path. It does not claim
to reduce complete Store parses, writer work, package I/O, decompression,
reopen cost, cache footprint, remote reads, cold-cache behavior, parallel
scaling, or unrelated XLSX operations. The whole non-iWork performance goal
remains open.

The fixed-path primary passes qualification for mean, p50, p95 and p99.
The retained production change removes measured repeated work with preserved
semantics. Registry entry `claim-0467-xlsx-cell-attributes` is limited to this one
scenario/corpus and four accepted latency statistics, with all DOC guards and
original adverse observations retained outside that claim scope. Its standard
ABBA package preserves explicit primary projections of the complete raw reports.

See the [source review](../results/change-0467/source-review.md), the retained
[0466 profile record](0466-xlsx-dense-commit-profile.md), and the
[0467 evidence bundle](../results/change-0467/) for reproducibility details.
