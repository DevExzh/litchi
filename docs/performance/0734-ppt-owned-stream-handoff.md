# 0734 — PPT owned stream handoff

Retained: moving owned PPT buffers into the CFB writer improves primary
paired p50 by 5.55% and removes five allocations on both tested decks.
Secondary latency is inconclusive; one primary tail regression remains
explicit below. The change removes copying at the PPT-to-CFB writer boundary.
Previously, finish owned every stream payload but borrowed it into
`OleWriter::create_stream`, which allocated and copied it again. The candidate
uses existing `create_stream_owned` for the assembled Document stream, updated
Current User stream, and untouched auxiliary payloads. It consumes only the
private editor at its existing finish boundary.

Source layout is adopted before transfer. Complete paths and stream insertion
order remain unchanged. Duplicate selected complete paths take the previous
borrowed-copy route; the ordinary unique-path route transfers each payload
once. Missing selected paths preserve prior behavior. The unchanged no-op
branch returns exact original bytes. The original source Arc remains immutable.
The removed clone can no longer fail allocation; other allocation, package,
limit, mapping and reopen failures remain typed. No format validation is removed.

Both sector policies retain their existing meaning. Reuse planning still
validates the composed stream view before emission. The bounded output sink,
PPT rewrite validation, public reopen and semantic validation, unrelated-stream
and picture checks, and both artifact digests remain in place. There is no new
cache, public API, execution policy, ambient I/O or unsafe code.

## Workload and evidence

The ordinary public owner opens a PPT snapshot, removes live slide index one,
and commits. Primary `45543.ppt` goes from 11 to 10 slides; independent secondary
`41246-1.ppt` goes from 36 to 35. Source sizes are 385,024 and 275,456 bytes;
output sizes are 390,144 and 285,184. Both outputs have five logical streams.
The preservation oracle checks exact output hashes/inventories, normalized raw
CFB directory metadata, live record bytes, slide order and survivor payloads,
logical text, outline/list data, notes and comments. Each report rejects eight
deliberate corruption controls. The primary retains the sealed 0728 oracle.

Before captures precede the production edit. Identical probe source and build
flags produce separate baseline and candidate binaries. The fixed comparison
uses nine process pairs per fixture, 50 samples and three warmups each; the
allocation lane uses three process pairs per fixture with one sample and no
warmup. CPU 12 is pinned, pair order rotates, and all execution is serial.
Native timings exclude oracle validation outside the owner. Allocation
instrumentation is measured separately and makes no latency claim.

[The evidence packet](results/change-0734/README.md) retains source snapshots,
all commands and logs, before captures, exact fixture/build/probe identity,
qualification, the prospective schedule, raw measurements and independent
replay. Bootstrap intervals use nine paired process-level percentage changes,
10,000 resamples, seed 7334 and percentile 95% bounds. Samples within a process
are not treated as independent repetitions.

## Results and decision

Retain the scoped ownership change. The primary fixture improves consistently;
the secondary has a concrete allocation reduction without a clear latency
change. All 48 processes and 1,812 measured owner outputs pass the exact oracle
and independent audit. No measurements were discarded or selectively repeated.

| Metric | `45543.ppt` | `41246-1.ppt` |
| --- | ---: | ---: |
| Baseline process p50 range | 1,052.895–1,063.405 µs | 1,220.901–1,230.146 µs |
| Candidate process p50 range | 993.644–1,004.590 µs | 1,217.581–1,246.121 µs |
| Median paired p50 change | −5.55% | +0.35% |
| Paired p50 bootstrap 95% interval | [−6.21%, −4.96%] | [−0.41%, +0.87%] |
| Median paired mean change | −7.57% | −0.35% |
| Paired mean bootstrap 95% interval | [−7.74%, −7.18%] | [−0.55%, +0.33%] |
| Allocated bytes, before → after | 11,776,674 → 11,391,768 | 10,724,448 → 10,444,631 |
| Allocation calls, before → after | 5,663 → 5,658 | 18,667 → 18,662 |
| Peak live bytes, before → after | 2,686,521 → 1,992,885 | 1,976,223 → 1,976,223 |
| Retained output bytes | 390,144 unchanged | 285,184 unchanged |

All three allocation repeats agree exactly for each variant and fixture.
Allocated bytes decline by exactly the five output payloads: 384,906 bytes
(3.27%) and 279,817 bytes (2.61%). Primary peak live bytes decline 25.82%; the
secondary peak is unchanged. Earlier buffer release changes lifetimes as well
as removing copies; the measured peak change is not a sum of attributed phase
costs. The secondary timing intervals cross zero, so no secondary speedup is
claimed.

The sole positive >5% latency flag is primary cycle 2, repeat 0: p99/maximum
increases 9.56%, from 1,170.776 to 1,282.706 µs. With 50 samples, nearest-rank
p99 is the maximum, so these two fields describe the same observation. It is
retained without rerunning. Across all primary pairs, p50 improves 4.66–6.56%
and mean improves 6.78–8.12%. Seven p50 improvements and all nine mean
improvements exceed 5%; four p99/maximum improvements also exceed 5%. No p95
flag or secondary timing flag exceeds 5%. The tail observation is accepted as
an explicit limitation alongside the consistent central and allocation gains;
this batch makes no universal tail-latency claim. [Every flagged comparison](results/change-0734/report-stats.txt)
and [the disposition](results/change-0734/disposition.json) remain reviewable.

The 19 evidence-corruption controls pass. Eighteen are rejected by both
validators; the altered probe-command control is rejected by the analyzer
only. The per-validator coverage is retained in the packet. The synthetic
preflight passes on its first retained attempt.

## Correctness and ADR compliance

Twelve final serial gates pass: seven owner gates and five candidate probe
gates. Owner tests pass 1,211 with three existing ignored tests; doctests pass
14 with eight ignored; probe tests pass four. The unchanged baseline probe also
passes its five gates. Clippy and rustdoc deny warnings. The boundary audit
passes 64 packages and 241 internal dependency declarations with 11 explicit
existing debt items. Focused new tests compare the old borrowed writer under
Reuse and Rewrite, changed-record finish/readback, exact no-op and immutable
source behavior, finite limits, and duplicate-path fallback.

Independent source review approves measurement. It records non-blocking gaps
for injected writer-I/O failure and same-leaf/different-complete-path fixtures.
The source census permits exactly one changed file among 7,206 Rust/manifest
entries. Accepted ADR and scenario-taxonomy hashes remain unchanged.

| Constraint | Evidence |
| --- | --- |
| ADR 0001/0002/0024 ownership and topology | Private PPT finish uses the existing CFB-owned ingress; no new dependencies. |
| ADR 0003 snapshots, edits and patches | Transfer occurs inside consumed finish; immutable source and outer atomic publication stay intact. |
| ADR 0005 memory and evidence | Eliminated clone measured separately from latency; finite output limits and caller policy retained. |
| ADR 0006 preservation and security | Both validation layers, semantic reopen, limits, typed failures and exact output oracle remain. |
| ADR 0008 verification | Fresh affected-owner/probe gates, retained source and reproducible evidence. |
| ADR 0010/0011 package ownership | No archive implementation types enter a facade or public CRUD API. |

The scope is this public PPT operation on two fixtures on the recorded host.
Peak live bytes are boundary-relative allocator ownership, not RSS. The packet
does not establish cold I/O, instructions, concurrency, all producers or broad
CRUD speedups. The wider non-iWork performance goal remains active.

Owned build and binary scratch is removed. Post-cleanup analysis, independent
audit and all 19 corruption controls pass using exact cleanup receipts.

The staged whitespace check excludes three raw test stdout logs solely for
their terminal blank line (`build-0/1.log`, `build-1/1.log`,
`quality-0/2.log`). Their exact sealed bytes are preserved.
