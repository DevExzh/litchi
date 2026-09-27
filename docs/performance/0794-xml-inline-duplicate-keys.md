# 0794 — bounded inline XML duplicate keys

**Rejected; production restored exactly to the baseline.** The candidate removes
85.048% of allocation calls in large PPTX capture, but none of the fifteen primary
cases shows the required timing benefit. Tiny XLSX full-cell scanning regresses
5.511%, with its ratio confidence interval entirely above one, independently
triggering the cross-format veto. Allocation savings alone do not justify retention.

The experiment compares the five shared fail-fast XML attribute helpers with
four borrowed names stored inline, followed by a bounded vector and the existing
ordered-map fallback. The base is `008bea923a`; the separate formula lenient
iterator remains unchanged. See the [evidence packet](results/change-0794/README.md).

## Method and correctness

The primary generated PPTX matrix covers capture, staged commit, and lifecycle
for tiny, medium, large, ASCII-vendor, and Unicode-vendor shapes. Six alternating
paired blocks use 30 samples after three warmups. Two separate allocation blocks
use three samples without warmup. Baseline qualification contributes fifteen
one-sample reports. This gives 255 primary reports and 5,595 samples.

A separately built ordinary harness covers DOCX open/full-text for tiny and large
documents, and XLSX open/full-cell-scan for tiny and dense-wide workbooks. Six
paired blocks plus qualification give thirteen reports and 2,888 case samples.
The full-cell scan visits Sheet1; open checks all generated sheets. These timings
are not pooled with the PPTX probe. Both lanes use CPU 12 and serial processes.
The recorded host is an AMD EPYC 9R45 Linux x86-64 machine with Rust 1.95.0
and LLVM 22.1.2. The PPTX release probe uses opt-level 3, thin LTO, one
codegen unit, debug level 1, and unwind panics; the cross-format harness uses
its ordinary release profile. Exact host and build environments are retained
in the packet.

The predeclared primary rule requires a capture/lifecycle improvement of at least
3% with bootstrap upper bound below one. Any significant paired-p50 regression
above 5% rejects; paired allocation block medians may not increase in calls,
allocated bytes, net live bytes, or peak above entry. The cross-format lane adds
a significant 5% regression veto. Seeds 794079/794080, 10,000 resamples, and
endpoints 250/9749 were fixed before measurement. PPTX uses nearest-rank p50;
the ordinary cross-format harness uses the integer midpoint for its 30 samples.

All six final quality gates pass across fourteen packages: formatting,
all-features/all-targets compilation, 12,865 tests (89 ignored, 500 suites),
Clippy and rustdoc with warnings denied, and dependency boundaries. The latter
checks 65 packages and 244 declarations with eleven existing debt entries.
The separately retained source-identical baseline has 12,850 passing tests and
89 ignored. Existing probe tests are inherited from 0793 rather than counted as
fresh. Source review verifies original public visibility and aligned helper
bodies; no unsafe code or dependency is added.

Three failed candidate validation attempts are retained: accidental OLE common
visibility restriction, constructor shadowing in new tests, then formatting of
the renamed bindings. All were corrected before candidate performance capture.
An earlier failed patch-path check applied no source change but inadvertently
started baseline quality checks; that successful baseline run remains distinct.
Exact archives, logs, and transition checks are in the packet.

## Mechanism and tradeoffs

The first four names avoid a heap allocation. The fifth spills into capacity
eight; the vector grows through 32 names, then the 33rd distinct name moves into
an ordered map. Duplicate names retain the original byte positions and error
precedence; first error and exhaustion fuse the iterator. The long-value and
clone boundary tests cover inline, spill, and ordered states.

Quick-xml performs lexical parsing with duplicate checking disabled; local code
performs all duplicate checks. The trait documentation's reference to quick-xml's
linear check describes comparable cost, not ownership of the check. For an early
duplicate, lexical parsing may now scan a long quoted or unterminated value, or
whitespace after `=`, before returning the same duplicate error. An unquoted
value still rejects at the first non-whitespace byte. This is bounded by tag
bytes but is not equivalent fail-fast work. Caller limits remain unchanged.

The count-only helper confirms the iterator grows from 120 to 192 bytes. Inline
storage therefore trades a larger iterator for fewer short-lived heap allocations;
it is not a claim of uniformly lower stack or process memory. Its 84 iterations
and two reports are separate from public-workflow samples.

## Measured decision

Positive changes mean slower candidate execution. Intervals are bootstrap
intervals of the median paired process-p50 ratio, not pooled sample intervals.

| PPTX case | p50 change | Ratio interval |
|---|---:|---:|
| large/capture | +1.004% | 0.992390–1.016934 |
| large/commit | +2.122% | 1.011487–1.024562 |
| large/lifecycle | +1.322% | 1.009677–1.015075 |
| medium/capture | +0.880% | 0.999242–1.012941 |
| medium/commit | +0.750% | 1.006624–1.009021 |
| medium/lifecycle | +0.375% | 1.001574–1.006802 |
| tiny/capture | +0.325% | 0.998714–1.008161 |
| tiny/commit | +1.219% | 1.007403–1.013573 |
| tiny/lifecycle | +0.420% | 1.000503–1.005516 |
| unicode-vendor/capture | +2.558% | 1.020974–1.031271 |
| unicode-vendor/commit | +1.226% | 1.008420–1.016516 |
| unicode-vendor/lifecycle | +0.901% | 1.003329–1.013694 |
| vendor/capture | +2.028% | 1.015033–1.025702 |
| vendor/commit | +1.271% | 1.008746–1.017250 |
| vendor/lifecycle | +1.023% | 1.008927–1.011287 |

No primary row violates the 5% latency limit, and every paired allocation
resource guard passes. However, there is no eligible benefit: every primary
paired p50 median is slower. No threshold or workload was changed after capture.
The diagnostic inventory retains ten tail/RSS metric groups with at least one
paired block above 5%, and 26 between-process spread flags. These are distinct
from the frozen median-p50 decision; no tail or RSS improvement is claimed.

| Cross-format case | Shape | p50 change | Ratio interval |
|---|---|---:|---:|
| docx_semantic_full_text | large | -2.445% | 0.960779–0.996907 |
| docx_semantic_full_text | tiny | -3.945% | 0.945358–0.970282 |
| docx_semantic_open | large | -1.015% | 0.965642–1.001852 |
| docx_semantic_open | tiny | -1.113% | 0.979907–0.997013 |
| xlsx_full_cell_scan | tiny | +5.511% | 1.050283–1.081791 |
| xlsx_full_cell_scan | dense-wide | +0.536% | 0.998137–1.006945 |
| xlsx_open_owned | tiny | +0.801% | 1.002853–1.011804 |
| xlsx_open_owned | dense-wide | +0.203% | 0.977090–1.042027 |

The tiny XLSX scan exceeds both veto conditions: median ratio >1.05 and
bootstrap lower bound >1.0. DOCX improvements do not substitute for the primary
PPTX benefit requirement or override this veto.

## Allocation and profile evidence

| Large PPTX capture metric | Before | Candidate |
|---|---:|---:|
| Operation allocation calls | 72,106 | 10,781 |
| Operation allocated bytes | 4,767,939 | 843,139 |
| Net live bytes | 278,201 | 278,201 |
| Peak above entry bytes | 338,955 | 338,955 |

The two allocation blocks agree. Allocated bytes fall 82.316%; this measures
requested operation allocation traffic, not physical bytes copied or process RSS.
All four heap profiles satisfy the predeclared owner/counter qualification:
whole-process stack/histogram/summary totals agree, the filtered stacks exactly
equal the owner subset, and owner calls equal operation allocation_calls.
Both repeats yield 72,106 before and 10,781 after. Whole-process calls are
608,290 and 241,608 respectively, including fixture creation and verification.

The raw nested quick-xml duplicate-check count falls from 61,342 to zero;
the candidate local checked-attribute frames account for 17 calls. Nested notes
inspection counts fall from 61,274 to 449. These overlapping counts corroborate
the allocation mechanism without establishing a timing gain. The profile
qualification does not retroactively repair the failed 0793 formula.

Across all lanes, 274 reports retain 8,487 measured samples plus 84 count-only
helper iterations. The primary and cross-format numerical decisions are also
recomputed independently from raw sample arrays. No production change remains.

## Next bounded question

The result demonstrates that eliminating these heap calls is insufficient for
latency improvement with this iterator design. The 72-byte iterator growth and
additional local duplicate bookkeeping are candidate costs, not established
causes. Any follow-up should compare instruction and branch costs on these exact
public cases before proposing another layout. The retained 0792 empty-tail
optimization and earlier cached-Part rejection are unchanged.

## Scope

Process RSS is a high-water observation from `/usr/bin/time`, not scoped operation
memory or causal proof. Heaptrack elapsed time does not establish native latency.
Owner profile qualification uses allocation_calls alone because reallocations
are already included. Nested frame counts overlap and must not be summed as
separate components. No cold-cache, device-floor, remote-range, concurrency,
native-producer coverage, or universal speedup claim follows. OLE2/OOXML work
remains active; ODF is deferred and iWork excluded.

The owned build target is removed with all six primary binary identities
recorded; two cross-format binaries and two helper copies have separate cleanup
witnesses. The candidate archive, raw reports, logs, four lockfiles, and review
notes are retained. The final seal covers every packet payload and this report
plus five indexes, with no production file in the committed change. Replay:

```sh
python3 -B docs/performance/results/change-0794/validate.py --require-final-seal --check-workspace
```
