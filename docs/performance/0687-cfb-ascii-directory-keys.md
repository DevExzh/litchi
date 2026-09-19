# 0687 — inline ASCII CFB directory keys

Status: final validation passed; scoped retention independently reviewed. `performance_claim: none`.
OLE2/OOXML remain active; iWork is excluded.

## Measured hypothesis

After [0686](0686-xls-oversized-index-admission.md), a fresh two-million-query
profile of the 54016 owned-source selected-cell route attributes 12.87% self
samples to shared CFB directory-name key construction. Its generic UTF-16
collection and comparison-key construction account for 4.41% and 5.00% of
total samples. Chain traversal remains larger at 49.47%; this batch does not
change its identity, validation or retention design.

CFB resolves each requested stream name through the same private helper used
by readers and writers. Short ASCII names such as `Workbook` do not need
UTF-16 scalar iteration or incremental SmallVec construction. The hypothesis
is that building both bounded keys directly removes useful repeated work
without a new cache, lifetime rule or memory reservation.

## Implementation and proof

The existing empty-name, NUL and first-forbidden-character checks run first.
Only ASCII names of at most 31 bytes enter the new branch. Each byte equals
its UTF-16 code unit, and ASCII uppercasing equals the existing simple mapping.
Two `[u16; 32]` arrays become inline SmallVec values through safe
`from_buf_and_len`; admitted length is below their capacity. Longer ASCII
names and every Unicode name retain the general path and its UTF-16 length
refusal. There is no unsafe code, dependency, public API or retained state.

Four new differential tests compare exact original/comparison keys and errors
against a frozen general reference. They cover every one- and two-byte ASCII
name, every byte across lengths 0–33 and positions around the 31-unit boundary,
NUL/forbidden/length precedence, Unicode mappings and supplementary-plane
length boundaries. The existing full-Unicode-scalar uppercase differential
continues to run. Independent source review found no correctness blocker.

## Evidence and limits

The [packet](results/change-0687/README.md) binds baseline `175b6c454`, final
sources, unchanged probes and fixtures, binaries, commands and raw results.
Native A/A and A/B/B/A cover 24 case/source groups with 14,400 fresh owners
and 115,200 queries. Separate same-owner controls, allocation, source-I/O,
hardware-counter and native-child RSS captures preserve their different scopes.
The tiny, missing, numeric, late-target and formula-refusal routes are explicit
controls, rather than being omitted from an aggregate speedup.

All six owner/facade quality gates pass, plus DOC/PPT consumer tests:
4,392 tests passed, zero failed, 27 existing ignored. This does not establish
cross-platform or native Office compatibility. Physical cold-cache, remote
latency and concurrent scaling remain outside this experiment.

## Final measured results

The table gives the two paired median changes for the eighth query. These are
fresh-owner samples with earlier queries building the optional index; longer
same-owner loop controls are reported separately below.

| Stored target | Owned | File, warm OS cache |
|---|---:|---:|
| 54016 first | −13.61% / −12.93% | −6.75% / −8.63% |
| 54016 late | −12.38% / −12.19% | −9.57% / −8.95% |
| Plan1 first | −16.77% / −16.87% | −4.10% / −6.15% |
| Plan1 late | −16.48% / −16.85% | −6.10% / −5.49% |
| Simple first | −21.67% / −21.67% | −7.57% / −7.05% |
| 45365-2 first | −16.90% / −18.06% | −5.86% / −7.30% |
| Generated 70,001-cell fixture | −19.38% / −18.75% | −6.54% / −7.86% |

For example, owned 54016 first falls from 1,470 ns to 1,270/1,280 ns,
and Simple from 600 to 470 ns. Missing indexed queries, which do not resolve
a CFB name, are effectively unchanged; tiny timer quantization and single-leg
variation are retained in the raw results.

The sum of open and eight query timers improves 16.48–16.56% for Simple
owned and 8.03–8.07% for file input. Simple missing improves 13.85–14.06%
owned and 6.69–7.45% file. Plan1 first improves about 3.5% owned and 2.9%
file. Formula-refusal windows improve 3.99–4.83% owned and 4.58–5.28% file;
errors remain uncached. Large 54016/default windows are approximately flat:
+0.08%/+0.27% owned and −0.45%/−0.67% file. No general large-workflow
speedup follows from the warm-query result.

Nine separate processes per leg, each with 50,000 queries after two warmups,
confirm the repeated-work reduction. Owned loop means improve 11.90–12.19%
for 54016 first, 22.04–22.70% for Plan1, 27.53–28.64% for Simple and
25.92–27.21% for 45365-2 first. The 54016 late target improves only
4.08–4.53% owned and 2.23–2.26% file in this shape; remaining chain work
limits the gain. These are distributions of process loop means, not individual
query latency distributions. All long-loop cases use the default 2 MiB limit.

## Regressions and rejected candidate

The final 45365-2 build query regresses 7.23–9.19% for its first owned
target, 6.73–8.51% for file, and 5.12–8.44% for the late owned target.
Its first-target open-plus-eight window regresses 5.39%/+3.36% owned and
2.40%/+3.68% file; other 45365-2 windows rise 1.05–4.42%. Independent
review accepts these explicit costs for scoped retention; no gain is claimed
for every opening or indexed workload.

[Median triggers](results/change-0687/regressions.md) and
[mean/tail triggers](results/change-0687/tail-regressions.md) include build
mean/p95/p99 regressions and noisy Simple/file open p99 changes. Bootstrap
intervals do not remove shared-host drift. In particular, the baseline
54016-late/file open-plus-eight A/A changes −5.37%; its sub-percent final
workflow movement is not treated as a gain. One hundred samples do not
establish a stable population p99 or a tail-latency guarantee.

The first candidate checked ASCII before the length bound. It improved warm
queries but regressed 54016 open by 11–16%, with unchanged counted I/O and
query allocation gauges. Cold profiles exposed changes in unrelated XLS
parsing hot spots and same-sized functions shifted in the binary; they did
not establish a source-level cause. Review also found an avoidable full ASCII
classification of overlong input. The final candidate checks the bound first,
removing that work and reducing code size. Its 54016 open medians return to
approximately baseline (owned default −0.45%/−0.89%). The exact cause of the
initial cold regression is not isolated; the final paired measurements, rather
than a causal claim about check order, establish its absence here. The initial
source, checks, measurements and diagnostic profiles remain archived.

## Work and memory costs

All 96 query-allocation groups match exactly before/after over three repeats
per binary, including calls, requested bytes, peak-live and retained deltas.
All 12 counted routes preserve opening and all eight queries' read counts,
bytes, freshness observations and outcomes. No new cache or persistent object
is introduced. These query allocator captures exclude opening and do not
measure stack usage.

Separate whole-process counters subtract N=10 from N=100,010. Representative
owned extra-query instructions/cycles are:

| Target | Instructions before → after | Cycles before → after |
|---|---:|---:|
| 54016 first | 27,038 → 25,580 | 5,580 → 4,875 |
| 54016 late | 46,179 → 45,449 | 7,994 → 7,664 |
| Plan1 first | 13,029 → 11,571 | 3,335 → 2,587 |
| Simple first | 7,322 → 5,863 | 1,964 → 1,413 |

Instructions fall 5.39%, 1.58%, 11.19% and 19.93%, respectively. Both source
modes, branches, branch misses, cache misses and page faults remain in the
counter comparison. These estimates include whole-process setup differences
and the timing wrapper; native phase timers are the latency evidence.

Disassembly retains a real tradeoff: helper code grows 1,633 → 2,440 bytes,
and its stack reservation grows 248 → 376 bytes. The ASCII branch zeroes two
64-byte arrays and copies their complete inline storage into the result while
bypassing generic SmallVec growth/extension. The compiler vectorizes part of
this safe scalar code; no hand-written SIMD or unsafe code was added. The
initial candidate was larger still at 2,672 bytes. Matched warm profiles lower
name-construction self share from 12.87% to 7.76%; chain traversal remains
51.52% of final self samples. Profile percentages are not absolute work counts.

Native-child peak RSS has review triggers. Long-run medians rise 3,972 →
4,192 KiB for 54016 late/file (+5.54%), 2,604 → 2,744 for Simple/owned
(+5.38%) and 2,524 → 2,744 for Simple/file (+8.72%). Plan1/file rises
2,796 → 3,000 KiB at N=10 (+7.30%) and 2,832 → 2,996 at N=100,010
(+5.79%). All three repeats and unflagged cases are retained. The cause of
these process differences is not isolated; they are not equated with the
128-byte stack increment or dismissed as an allocator result. No RSS or
universal memory-reduction claim is made.

## Validation and disposition

The 126-real-fixture differential matches exactly for owned/file sources.
The generated 70,001-cell fixture retains exact full-visitor counts and
semantic digests. Formatting, owner/all-target checks, warning-denied Clippy,
tests, warning-denied rustdoc, dependency boundaries, strict/structural claims,
report classification, coverage and non-iWork gates pass. Four new focused tests
are included in the 4,392 passing tests. No new fuzz campaign or native Office
round-trip evidence is claimed.

Independent source and performance review recommend retaining the final
bound-first candidate with the disclosed costs; see the
[review](results/change-0687/review.md).
The remaining chain traversal, SST setup, 45365-2 build cost, broader source
and concurrency evidence, and full GOAL completion remain open. No registered
performance claim or CRUD coverage row is promoted.
