# 0533: inline checked CFB sector claims with cold error helpers

The 0533 campaign measures a CFB/OLE2 candidate based on revision
`0dd079b95982aa1fbef5eddd1ec3743261dcaf20`. It moves the three formatted
`claim_sector` failures into private cold, non-inlined helpers and gives the
checked method an ordinary inline hint. The candidate keeps the conversion,
bounds, ownership, error-order and role-publication behavior of the baseline;
the full physical reconciliation pass remains in place.

The candidate is accepted by the frozen native, allocation, constructor-Ir,
reviewed-regression and quality gates. Cleanup, [post-cleanup verification](../results/change-0533/verification.json)
and [evidence sealing](../results/change-0533/SHA256SUMS) are complete.

## Frozen protocol and scope

The [campaign plan](../results/change-0533/plan.json) freezes nine XLS
workflows and three direct CFB shapes (`tiny`, `many-small` and `few-large`).
Native capture uses two paired repeats in ABBA order with 20 warmups and
1,000 samples per case: 24,000 samples per stage and 48,000 matched samples overall. The four
primary XLS rows each require at least 3% lower p50 in both repeats. The
separate allocator lane uses two repeats, three warmups and 30 samples for
each of the 24 stage rows. Constructor profiles use five calls for each of
four jobs per repeat; hardware is a whole-child diagnostic lane.

The workload is synthetic, warm and in memory. It does not establish cold
I/O, range or remote providers, native Office producers, concurrency scaling,
or a generic OLE2/OOXML gain. OLE2/OOXML remains the active priority; ODF is
deferred until that goal is complete, and iWork is excluded.

## Source boundary and behavioral checks

The candidate adds private cold `#[inline(never)]` helpers for
`claim_sector_index_error`, `claim_sector_bounds_error` and
`claim_sector_conflict_error`, and marks only `claim_sector` with ordinary
`#[inline]`. Their messages remain byte-for-byte compatible:

```text
{role} sector {sector} does not fit usize
{role} sector {sector} is outside the file
Sector {sector} is claimed by both {existing} and {role}
```

The method still converts the hostile `u32`, performs checked `get_mut`,
rejects an existing role, and stores the new role in that order. Failed
conversion, missing-slot and conflict paths leave the role map unchanged.
`claim_chain` retains its partial-on-late-error behavior, stream validation
retains collect-then-claim ordering, and
`validate_physical_sector_layout` retains its complete final scan. The
[source review](../results/change-0533/source-review.md),
[test review](../results/change-0533/test-review.md),
[`claim_sector` source](../../../crates/litchi-cfb/src/file.rs), and
[`tests.patch`](../results/change-0533/tests.patch) bind these obligations.

## Native and allocation evidence

The [matched comparison](../results/change-0533/comparison.json) passes the
frozen primary and allocation checks. Primary native p50 improvements are:

| Primary workflow | Repeat 1 | Repeat 2 |
| --- | ---: | ---: |
| `xls_source_backed_open` | 23.3082% | 23.2278% |
| `xls_source_backed_open_one_cell` | 22.6598% | 20.5382% |
| `xls_owned_source_open` | 21.6328% | 22.4034% |
| `xls_owned_source_open_one_cell` | 22.2845% | 24.5798% |

The four primary rows pass the 3% rule in both repeats. CFB `few-large`
improves 23.6156% and 23.6937% in p50; across all 12 scenarios, p50, mean,
p95 and p99 improve in both repeats. The all-row p50 improvement range is
2.6432–24.6726%. No matched peak-RSS row exceeds the 5% review threshold;
the observed matched peak-RSS changes range from −0.7800% to +0.5981%.

Each of the four allocation vectors is exactly identical in all 24 rows
between stages and repeats: allocation calls, reallocation calls, allocated
bytes and incremental region peak are unchanged. Allocator elapsed time is
kept separate from native elapsed time and does not authorize a native
latency claim.

## Profiles and generated code

The [profile comparison](../results/change-0533/profile-comparison.json)
contains eight constructor children, 40 timed dumps and six separate CFB setup
dumps per stage: 16 children, 80 timed dumps and 12 setup dumps overall.
Positive incoming edges classify timed calls; setup is excluded. Every
XLS-owned and CFB parent constructor records fewer instructions in both
repeats. XLS owned-source constructor Ir falls 18.9967%/19.0280%; CFB
few-large falls 21.4860% in both repeats, while many-small and tiny fall
0.3504% and 0.2325%.

The six baseline out-of-line `claim_sector` variants are each 363 bytes with
a 112-byte stack reservation. The candidate has no selected out-of-line
`claim_sector` variant: the caller disassembly shows the sector load, role-map
bounds comparison, out-of-bounds branch, existing-role check, refusal branch
and role store in both stream-validation loops, with calls only on error
paths. The candidate's bounds and conflict helpers are 132 and 171 bytes.
The conversion helper is absent from this x86-64 binary because every `u32`
fits in `usize`; its checked source branch remains portable behavior.

Zero candidate attribution to the old out-of-line claim symbol does not mean
ownership checks became free: the remaining work is now inside callers.
The physical-reconciliation instruction totals remain unchanged in every
paired profile, and its twelve recorded variants remain 275 bytes each.
Parent-constructor totals establish the instruction reduction without adding
overlapping inclusive owner costs.

Inlining grows selected stream-validation code from 25,832 to 28,373 bytes
and selected load-FAT code from 80,175 to 85,913 bytes. The normal executable
grows from 60,132,200 to 60,140,440 bytes (+8,240 bytes, +0.0137%). This
caller text growth is disclosed without an instruction-cache or working-set
claim. The standalone harness is its own workspace and has no explicit
release LTO profile; root-workspace LTO is not an explanation for these
observations.

## Adverse variation and hardware limits

The comparison retains four matched over-5% review flags:

| Row | Metric | Change |
| --- | --- | ---: |
| `xls_eager_open_one_cell`, repeat 1 | elapsed standard deviation | +25.3253% |
| `cfb_open` `tiny`, repeat 2 | elapsed standard deviation | +9.5771% |
| `xls_owned_source_open_list_worksheets`, repeat 2 | elapsed maximum | +340.6789% |
| `xls_owned_source_open_list_worksheets`, repeat 2 | elapsed standard deviation | +239.9376% |

Thirty same-build max, standard-deviation and system-time variation records
remain in the bundle with completed individual review. In the owned-list row,
candidate repeat-2
maximum is `669,572 ns` versus `151,941 ns` for the baseline; the outlier is
retained, and no uniform tail-maximum stability claim follows. The review
also keeps the distinction between matched regressions and same-build host
variation. The three CFB `rss.system_seconds` records repeat the same
whole-child `0.17 -> 0.16 s` value across tiny, many-small and few-large;
their 0.01-second resolution makes them coarse and they are not independent
or operation-local observations.

The flagged rows still improve their paired central and percentile metrics:
the eager one-cell row improves p50/mean/p95/p99 by 6.2282%/6.3651%/5.3455%/
5.5623%, the CFB tiny row by 4.3103%/4.1241%/5.0420%/5.7851%, and the owned
list repeat-2 row by 24.6726%/24.0106%/23.2160%/23.0165%. The flags concern
spread or maximum only; no samples are discarded and no cause or stable-tail
claim is inferred.

Hardware captures cover the whole child, including setup, copies, queries,
oracles, drop and report generation. Cycles fall 3.8477%/3.1821%,
instructions 9.2474%/9.1767%, branches 9.3843%/9.3087%, and branch misses
2.4075%/4.3496%. IPC decreases from 1.6539/1.6674 to 1.5610/1.5641, while
branch-miss ratio increases from 0.2305%/0.2339% to 0.2482%/0.2467%.
These counters cannot support operation-local IPC, branch-prediction or
throughput claims and do not replace the constructor clock or profile gate.

## Quality, custody and disposition

The [quality summary](../results/change-0533/quality-summary.json) records 14
passing checks and 4,378 executed tests, including both CFB feature modes and
the XLS, DOC and PPT suites. The bundle retains source/build manifests,
candidate and baseline binary identities, corpus hashes, host sidecars, raw
vectors, assembly, profiles, hardware output and the [mechanism review](../results/change-0533/mechanism-review.md).
No new fuzz campaign is claimed; the existing test gates are separate
correctness evidence. A changed candidate or rebuilt final binary requires a
fresh full ABBA comparison under the frozen plan.

The [decision](../results/change-0533/decision.json) records `accepted` after
the independent adverse review, retaining all four matched flags and all 30
same-build records with individual explanations. The coordinator's post-cleanup
verification and evidence sealing remain open; those custody steps do not
change the measured disposition.

## Next OLE2/OOXML measurement

The next separate hypothesis is a fresh paired walk for the residual
physical-reconciliation cost. The chain collector remains the
largest current exclusive owner, but the rejected 0524 visited-bit fusion is
not revived; this candidate also does not merge collection, claiming or the
final physical pass. Any follow-up still needs source-bound assembly and
matched native/allocation evidence. ODF remains deferred behind completion of
the OLE2/OOXML optimization goal.
