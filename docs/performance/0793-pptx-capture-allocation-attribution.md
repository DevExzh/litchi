# 0793 — PPTX capture allocation attribution

**Diagnostic only; frozen qualification fails.** All 50 reports and 130 measured
samples retain source/output identity and semantic readback. The owner-filtered
stacks reproduce the exact capture-wrapper subset in all ten traces, but the
frozen counter formula double-counts reallocations. The failure is retained;
no owner allocation fractions, production change, or speedup is claimed.

## Scope and controls

The base is `1d17b5b0e7`, including the retained 0792 empty-tail optimization.
All 9,196 production files and 35 architecture/goal/taxonomy inputs match the
captured source. The five generated PPTX shapes and publication/readback oracle
are inherited from the sealed 0792 packet. See the
[0793 packet](results/change-0793/README.md) for source, host, commands, raw
traces, controls, and offline replay.

Four freshly built binaries distinguish direct and wrapped capture with and
without operation allocation counters. Two alternating repetitions, five
shapes, and three samples without warmup give 40 control children and 120
samples. Ten separate heaptrack children use one sample each. CPU 12 is used;
all builds, captures, and twenty heaptrack decodes run serially. The non-inlined
wrapper encloses only `Package::opened_presentation`, including a black-box
reference to the result before returning it. Fixture construction, output
verification, and result drop stay outside that boundary.

Operation count/byte deltas, net live bytes, and peak above entry agree across
direct and wrapped counter controls. Absolute live boundaries and process
high-water counters are not operation deltas. Every whole-process allocation
stack sum equals its print-summary and size-histogram count. The filtered
flamegraph equals the exact wrapper subset of the unfiltered flamegraph.
Filtered print summaries and histograms remain whole-process observations.

## Retained failure and supplementary explanation

The frozen rule compares wrapper-owned allocations with
`allocation_calls + reallocation_calls`. Source inspection after capture shows
that `Counters::reallocation_locked` already increments both counters: the
second is a subset of the first. The frozen rule therefore fails in all ten
traces. Neither the plan nor its captured inputs were amended.

The separate [counter-semantics note](results/change-0793/counter-semantics.md)
explains a post-capture comparison against `allocation_calls` alone. This
comparison matches all ten observations but does not replace the failed gate.
The following raw counts are identical in both repetitions; nested rows overlap
and are not disjoint categories or authorized fractions.

| Shape | Whole-process calls | Wrapper-owned calls | Frozen expected calls | Nested duplicate check | Nested notes inspector |
|---|---:|---:|---:|---:|---:|
| Tiny 3×4 | 13,634 | 1,338 | 1,488 | 444 | 455 |
| Medium 12×8 | 24,071 | 2,869 | 3,254 | 1,038 | 1,046 |
| Large 100×100 | 608,290 | 72,106 | 74,709 | 61,342 | 61,274 |
| ASCII vendor 12×8 | 33,002 | 3,157 | 3,734 | 1,254 | 1,262 |
| Unicode vendor 12×8 | 33,038 | 3,157 | 3,734 | 1,254 | 1,262 |

For large capture, the 2,603-call discrepancy exactly equals the reallocation
subset counter. Whole-process calls include fixture creation and readback;
they must not be described as capture calls. No profile elapsed time is used
as native latency evidence, and no intercepted process heap peak is interpreted
as an operation peak or RSS.

## Next bounded experiment

The [source review](results/change-0793/source-review.md) identifies quick-xml's
checked-attribute duplicate-key `Vec<Range<usize>>` as a plausible source of
short-lived allocations. First-key insertion allocates; the retained empty-tail
shortcut does not affect attribute-bearing elements. The raw stack counts
support investigating this path, subject to a correctly predeclared allocation
qualification in the next experiment.

A candidate should use safe bounded inline duplicate-key state with an ordered
fallback, preserving attribute error positions and precedence, duplicate checks,
malformed-value behavior, XML budgets, and fail-fast semantics. It still needs
focused differential tests and the full public capture/commit/lifecycle paired
matrix before retention. No such candidate is introduced here. The secondary
presentation scan remains lower priority under the 0791 native profile.

## Validation and limits

All four builds succeed. Probe formatting and 35 fresh probe tests pass (28
ordinary helper/counter tests and seven real-allocator tests). Existing unused
probe-helper warnings are retained in build logs. Production correctness gates
are inherited from source-identical 0792, not rerun or counted as fresh tests.
The main replay and independent raw audit verify captured evidence while
reporting the frozen qualification failure explicitly.

The owned build target is removed after recording all four binary identities.
The packet and this report plus five performance indexes are sealed for replay.
No CRUD coverage row, native-producer evidence, cold/range behavior, concurrency
claim, or broad goal completion is promoted. OLE2/OOXML remain active, ODF is
deferred, and iWork is excluded. The rejected 0787 candidate remains rejected.
