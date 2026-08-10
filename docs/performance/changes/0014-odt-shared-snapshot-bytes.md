# Change 0014: ODT shared snapshot bytes

Date: 2026-08-11

## Decision

Accept a narrow ODT transaction-snapshot ownership handoff. When an already
validated `Document` creates a transaction `Snapshot`, the snapshot now clones
the package's private `Arc<Vec<u8>>` handle instead of allocating and copying
the complete archive. Direct `Snapshot::from_bytes` ingress keeps its existing
independent validation and copy behavior.

The 64 MiB transaction package bound still runs before publication. Exact
source bytes, immutable edit ownership, source-checked reversible patches,
durable operations, signed/encrypted exact no-ops, changed-envelope refusal,
compact-XML audits, candidate parsing, and complete semantic readback are
unchanged. The shared handle is crate-private; no archive type, dependency,
runtime, cache, lock, unsafe code, or public API is introduced.

A private unit test proves `Arc::ptr_eq` between the document's package bytes
and the captured snapshot, in addition to byte equality. This makes the removed
allocation an enforced ownership contract rather than an inferred compiler
optimization.

## Matched latency result

The existing ODT semantic harness already isolates this handoff in
`odt_semantic_noop_edit_save`: corpus construction and `Document::from_bytes`
happen before timing, while the timed interval starts at `Document::edit` and
ends after exact no-op publication. `odt_semantic_one_edit_save` and
`odt_semantic_open` are changed-publication and unrelated-open guardrails.

The completed harness was frozen on production base `56dfde4fd` before the
production source changed:

- before SHA-256:
  `793d01bc572a56ad88e5ba80bbe754621c557f6780d8529d1159d9c20d1d5f2f`
- after SHA-256:
  `c32f9bac228d089d707ed37dcc0f17d246c8428db89933980dadbe3882ddedc1`

Both states ran pinned to CPU 2 in before-A, after-A, after-B, before-B order,
with three warmups and 30 measured samples per leg. The table pools both legs
(60 samples per state). Mean intervals are two-sided Student's-t 95%
intervals. Times are microseconds except the changed edit/save rows, which are
milliseconds.

| Case | Before p50 / p95 / p99 | After p50 / p95 / p99 | p50 delta | Before mean (95% CI) | After mean (95% CI) | Mean delta |
|---|---:|---:|---:|---:|---:|---:|
| No-op edit/save, medium | 0.440 / 0.581 / 0.851 us | 0.321 / 0.470 / 0.551 us | **-27.05%** | 0.451 (0.429-0.473) us | 0.342 (0.326-0.358) us | **-24.19%** |
| No-op edit/save, large | 3.950 / 5.698 / 34.729 us | 3.219 / 3.926 / 4.396 us | **-18.51%** | 4.478 (3.431-5.526) us | 3.154 (3.017-3.290) us | **-29.58%** |
| One edit/save, medium | 0.715 / 0.772 / 0.826 ms | 0.734 / 0.791 / 0.824 ms | +2.77% | 0.718 (0.710-0.725) ms | 0.734 (0.726-0.743) ms | +2.36% |
| One edit/save, large | 19.589 / 21.382 / 22.537 ms | 19.866 / 21.314 / 23.218 ms | +1.41% | 19.754 (19.534-19.975) ms | 20.004 (19.782-20.227) ms | +1.27% |
| Open guard, medium | 29.646 / 46.464 / 52.703 us | 30.432 / 42.930 / 47.535 us | +2.65% | 32.163 (30.705-33.622) us | 32.680 (31.428-33.931) us | +1.60% |
| Open guard, large | 758.614 / 876.978 / 1079.537 us | 771.837 / 852.174 / 955.706 us | +1.74% | 777.305 (760.195-794.415) us | 781.334 (768.792-793.877) us | +0.52% |

One interrupted before-B large no-op sample accounts for the 34.729 us p99;
the pooled p50 and the non-overlapping mean intervals still support the
targeted result. All guardrail p50 changes remain below 3%, and their mean
intervals overlap. Changed publication is intentionally unaffected because
package rewrite and complete reopen dominate that path.

Raw samples:
[`before A`](../results/abba-odt-shared-before-a.json),
[`after A`](../results/abba-odt-shared-after-a.json),
[`after B`](../results/abba-odt-shared-after-b.json), and
[`before B`](../results/abba-odt-shared-before-b.json).

## Allocation and memory result

Valgrind Memcheck over 50 large no-op samples reports exactly 100 fewer
allocation calls (9,330,163 to 9,330,063) and 1,423,002 fewer allocated bytes
(2,712,437,270 to 2,711,014,268). The byte reduction is exactly 50 archive
copies plus their vector storage overhead; whole-process allocated bytes fall
0.052%, because complete correctness reopening after each timed sample remains
in the profiler process.

Heaptrack over 20 matched large no-op samples reports exactly 40 fewer
allocation calls. Before the change it attributes 20 28.42 KiB vector
allocations and 20 `Arc` control-block allocations directly to
`Snapshot::from_document`; after the change it attributes zero allocations to
that function. Peak heap moves from 18.00 to 17.97 MB (-0.17%), while profiler
RSS moves from 29.48 to 29.84 MB (+1.22%). A reverse-order uninstrumented GNU
Time run records 30,976 versus 30,848 KiB maximum RSS (-0.41%). Peak memory is
therefore flat within process/profiler noise while the targeted per-snapshot
copy and allocations are demonstrably removed.

## Correctness and contract gates

- The complete all-feature `litchi-odt` suite passes: 522 unit tests, every
  integration suite, and 55 doctests. This includes exact no-op, source-checked
  patch, inverse, durable replay, signed/encrypted refusal, malformed input,
  bounded history, and genuine LibreOffice preservation coverage.
- The all-feature library warning-denied Clippy gate passes. The all-target gate
  remains blocked by four pre-existing `cloned_ref_to_slice_refs` findings in
  the unchanged `packaged_transactions` integration test; it is not reported
  as passing.
- The performance harness passes 22 tests, formatting, warning-denied Clippy,
  and a release 21-record tiny ODF smoke. The established 88 selectable / 36
  default case counts and ODF CI matrix are unchanged.
- All four ABBA reports parse as JSON, repository formatting passes, and no
  manifest, dependency edge, CI topology, OOXML/OLE2/RTF behavior, or iWork/IWA
  file changed.

## Remaining limitations

The optimization applies only when a validated ODT `Document` creates its
transaction snapshot. Direct byte snapshot ingress still validates and owns an
independent allocation. Ordinary ODT open and repeated semantic queries remain
eager, and changed publication still rebuilds and reopens the package. The
generated text corpus does not replace real-producer, media, unknown-extension,
malformed, signed, encrypted, cold-source, bulk-edit, or conversion/export
matrices. Native DOC/XLS/PPT semantic baselines are the next non-iWork
measurement prerequisite before another CFB ownership experiment.
