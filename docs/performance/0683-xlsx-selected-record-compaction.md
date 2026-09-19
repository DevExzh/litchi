# 0683 — compact XLSX selections and skip impossible text escapes

Status: retained after final correctness, allocation, paired timing and independent review. `performance_claim: none`; no registry or
CRUD coverage promotion. OLE2 and OOXML remain active; iWork is excluded.

## Mechanism and public boundary

This batch implements the compact-record reduction proposed by
[0679](0679-xlsx-scanner-publication-design.md). `SelectedRecord` now holds an
address and `SelectedPayload::{Cell, SharedString}` rather than two independent
optional payloads. Both scanner constructors already produced exactly one
payload; the type now expresses that invariant. The source-backed resolver no
longer walks every selected record to reject impossible both/neither states.

The temporary requested-SST vector reserves for the number of actual deferred
shared-string records K, rather than all N selected physical cells. Numeric,
inline, formula and explicit-empty records do not need that staging capacity.
Indexes remain `usize` at the existing dependency-reader boundary; the optional
narrow-index design remains separate work.

The low-level public `raw::selected_worksheet::SelectedRecord` changes: callers
match `record.payload` instead of reading `record.cell` and
`record.shared_string_index`. `SelectedPayload` is non-exhaustive. Accepted
ADRs 0001 and 0008 permit this breaking refactor without compatibility shims.
The ordinary `cell`, `cells`, and `visit_cells` signatures remain unchanged.

XML/MCE validation, selected/unselected dependency maxima, complete archive
verification, fallback, and source/execution fences stay in their existing
order. The worksheet and dependency readers are closed before callbacks begin.
Single-cell compatibility still distinguishes missing records, explicit empty
records and deferred shared-string references. No replay pass or early callback
is introduced.

## Measured regression and decoder extension

The initial compact-record candidate reduced allocations but regressed the
owning long-inline selected query by 10.50–10.87% p50 in the paired windows.
A separate 1,000-sample native ABBA diagnostic reproduced approximately 10.4%
more cycles despite 0.2% fewer instructions. Sampled profiles placed the extra
cycles in the unchanged `raw::strings::decode_spreadsheet_text` loop: estimated
self cycles rose from 12.5 to 15.8 billion. This locates the affected work;
it does not establish a specific allocator or cache cause.

The final candidate also uses the existing `memchr` dependency to find possible
underscore starts before invoking the SpreadsheetML escape parser. It skips
plain runs, while preserving literal escapes, surrogate adjacency, UTF-8 slicing,
and exact error offsets. The separate materialized shared-string decoder is
unchanged. The final measurements therefore apply to the combined patch.
Initial source, tests, measurements and diagnostic reports remain archived in
`initial-candidate/`; the initial latency regression is not an accepted result.

## Cost and remaining scope

The selected record vector is still O(N), and shared-string dependency values
remain retained until publication. These semantic vectors did not charge the
managed `Memory` budget before this batch and still do not; allocation
measurements do not establish hierarchical admission or an RSS bound.
The change adds no cache, executor, ambient source or production unsafe code.

## Allocation observations

On the pinned 64-bit toolchain, `SelectedRecord` is 88 → 80 bytes while `Cell`
remains 72 bytes. The same standalone allocator measured three identical runs
per workload before and after. These figures are operation deltas with package
setup outside the counted interval, retaining the returned `cells` vector
through the final gauge where applicable.

| generated selection | requested bytes before → after | peak logical live bytes before → after |
| --- | ---: | ---: |
| dense numeric, `visit_cells`, 65,536 cells | 72,501,836 → 70,929,004 | 14,671,449 → 14,147,161 |
| dense numeric, `cells`, 65,536 cells | 77,744,716 → 76,171,884 | 14,671,449 → 14,147,161 |
| shared strings, `visit_cells`, 65,536 references | 113,225,608 → 112,177,064 | 14,792,407 → 14,268,119 |

The dense requested-byte reduction is 1,572,832 bytes: smaller growth requests
for the record vector plus removal of the unused 524,288-byte index buffer.
The shared-string case still needs its index buffer; its reduction is 1,048,544
requested bytes. Both lose 524,288 peak logical live bytes (3.57% and 3.54%).
Returned-result retention is unchanged. Smaller sparse, formula and long-inline
cases save fewer bytes; the real POI fallback control's allocation figures are
identical. A late fallback still saves scan allocation bytes but has unchanged
peak memory because materialization sets the peak.

These are allocator gauges, not RSS, managed admission, physical cold behavior,
or general XLSX memory guarantees. The selected route remains O(N).

## Final paired timing and native diagnostic

These are scoped engineering observations for the combined patch. Each ABBA
leg contains 20 native timing samples after three warmups, pinned to CPU 12.
Negative deltas mean lower p50 time; comparisons are B1/A1 and B2/A2.

| generated selected query | paired p50 deltas |
| --- | ---: |
| dense numeric, visitor | −0.57%, +0.10% |
| dense numeric, owning cells | −1.31%, −0.45% |
| shared strings, visitor | −1.33%, −0.29% |
| long inline text, visitor | −54.94%, −55.09% |
| long inline text, owning cells | −54.82%, −55.05% |
| escape-heavy Unicode, visitor | −2.94%, −3.64% |
| escape-heavy Unicode, owning cells | −1.18%, −3.68% |

The long-inline owning query falls from approximately 2.13 ms to 0.96 ms.
A separate 1,000-sample-per-leg native ABBA diagnostic confirms p50s
2,136,416 / 965,159.5 / 962,160 / 2,128,870 ns. Whole-process instructions
fall from 41.12/41.11 billion to 14.93/14.90 billion; cycles fall from
9.64/9.61 billion to 4.36/4.35 billion. These counters include process setup
and warmups. They support eliminating the scalar plain-byte search; they do
not prove a general XLSX speedup or a cache-locality mechanism.

All 56 final groups are retained in `comparison.json`, including mean and
sample p95/p99. None has a p50 regression above 5% in both paired windows.
Small changes and materialized/fallback controls are not promoted as speed
claims. For example, real-fallback selected visitor candidate drift is 11.27%,
so its unequal apparent improvements are not stable evidence. The 20-sample
windows and shared host do not establish tail-latency or concurrency guarantees.

## Validation and evidence

Final quality checks pass: formatting, all-feature/all-target owner compilation,
warning-denied Clippy and rustdoc, 1,380 owner tests and 97 XLSX/ODS facade tests.
The five new integration tests cover mixed payloads, repeated shared strings,
unselected dependency maxima, cold/warm parity, late zero-callback refusal and
callback re-entry. Two new decoder tests compare the scalar reference and exact
UTF-8 byte offsets. Existing scanner tests were migrated to the tagged type;
CRC, cancellation, source-change, ordering and fallback tests remain green.
Initial test-only unused-import/cast lints and a generated-fixture relationship
namespace error were corrected; their failed logs remain in the packet.

Ten generated/real differential cases agree across five public read routes
and across both binaries, including escape-heavy Unicode and a late unpaired
surrogate at byte 4096 with zero callbacks. The 56 workload groups have 336
allocation samples and 6,720 timing samples. The final baseline A/A maximum
absolute p50 drift is 10.44% (long-inline cold owning query); the next largest
is 3.67%. This noise limits small latency interpretations. Final paired results
and hardware diagnostics are retained in the [packet](results/change-0683/README.md).
