# Change 0152: CFB same-target MiniFAT single-flight release ABBA

Date: 2026-08-16

Status: correctness, liveness, and bounded logical source-event evidence
accepted. Timing is retained for audit only; no latency, allocation, RSS, cold,
physical-I/O, device, or generic performance result is accepted.

## Purpose and implementation history

This record compares the final lock-free uncontended-path revision
`f46381c6f3bede3a26ad1e214ef8a80ac7ce2b2b` with its clean control
`e486e4b12f5b6654ceeb6b2d137d107b4c48baad`. The feature under test is the
same-target MiniFAT single-flight introduced by `c270c8f3b5dad6e273272afe0a1765ece21e5f84`.

The initial single-flight implementation took its private
`Mutex<Option<slot>>` on every direct claim, including an uncontended
sequential open. The local regression investigation identified that mutex as
an uncontended hot-path cost; those superseded measurements are not release
evidence. `f46381c6f` keeps the overlap rendezvous but changes the normal path:

- an atomic state word carries the target SID, a bounded flight epoch, waiter
  intent, slot presence, and the direct/cache state;
- an uncontended eligible owner claims the next epoch with a compare-and-swap
  and does not lock the single-flight slot mutex;
- an overlapping caller announces intent and enters the mutex/condition
  variable slow path, where the owner can publish one bounded payload;
- the epoch and slot-presence bit prevent a delayed owner or waiter guard from
  releasing a later same-SID flight after an ABA transition; and
- source I/O remains outside the single-flight state lock. There is no worker
  thread, executor, global pool, or unsafe code.

The production change is confined to `crates/litchi-cfb/src/shared.rs`.
`shared_bulk.rs` is byte-for-byte unchanged. The candidate's performance
harness source differs only inside `cfg(test)`: the existing concurrent matcher
accepts the valid one-event vector `[D]`, and additional candidate formula
tests bind that result. The release execution code is identical, and those
tests add no runtime work to the benchmark operation. Timing is withheld for
the observed ABBA noise and review triggers documented below, not because the
release harness executed different code.

## Compared revisions and clean release matrix

The production source hashes are:

| Source | Control `e486e4b` | Candidate `f46381c` |
|---|---|---|
| `crates/litchi-cfb/src/shared.rs` | `872bc0159d39ccf2ceb4a3a5c1267aadf55d493f8330996e5afe4bbc3d6f0a62` | `531d2ba7f2f65c7e606abe5e8bbb1d59946e3bc8af3eff6866028f6727f29846` |
| `crates/litchi-cfb/src/shared_bulk.rs` | `3ea1cccdbdd4b45f3801983bfa8df67b94defe5476834ebc6175011760bb5b48` | `3ea1cccdbdd4b45f3801983bfa8df67b94defe5476834ebc6175011760bb5b48` |
| `tools/perf-baseline/src/main.rs` | `a6b2c94aa68d3f7b8599d235078329072f9f02d4ec2447e641776a9531b022f4` | `b5d9878b5ec75e0e9855ba30221ea31dfb6fab4e6b2e177d177ad2c8b028a7d0` |

The clean Rust 1.95.0 release binaries used for the reports were executable
mode `0775`:

- control: 40,421,200 bytes, SHA-256
  `11f38e1d4de831eece2de59f63d88fdad1d2cd5869755bf36120520b2962f49a`;
- candidate: 40,422,968 bytes, SHA-256
  `3a4880bc67a46a8cd86b5ec73ce5aaec54d2b17334b6535a55c59870baeaa2a7`.

The release runner is `litchi-perf-baseline` 0.1.0, Linux `x86_64`, with 12
selectors across two corpus shapes (24 records per leg). Each leg used 20
warmups and 500 measured samples. The strict order was `A1 control, B1
candidate, B2 candidate, A2 control`, for
`4 × 24 × 500 = 48,000` retained samples. All four reports record clean
worktrees. The generic report configuration carries fresh-child,
process-isolation, and `warm`/`cold-requested` filesystem fields, but those
fields are not exercised by these in-process production CFB selectors. Each
ABBA leg was one fresh CPU-pinned release process, with one execution worker
on one affinity-visible CPU; `filesystem_root_selected=false` was retained as
report metadata. No filesystem cold-state result is claimed.

The host was Linux 6.8.0-101-generic on an AMD EPYC 9575F, Rust
`1.95.0 (59807616e 2026-04-14)`, system allocator, CPU affinity `2`, 4,096-byte
pages, and `perf_event_paranoid=1`. The reports include range-simulation
parameters (100 us fixed latency, 25 us request overhead, 50 MiB/s bandwidth,
and a 4,096-byte maximum physical range), but these production CFB selectors
do not run through that simulator. No simulator-latency result is claimed.
The generic filesystem fields are retained in the reports but are not a
filesystem cache-state observation for these in-process selectors.

The deterministic corpora are:

| Shape / target | Siblings | Archive bytes / SHA-256 | Target SHA-256 | Root MiniStream `R` |
|---|---:|---|---|---:|
| many-small / 36 | 256 | 314,368 / `988be74585fe447a2695a8a584ce459181ab7a20b5bc4dd79d24ebf9e1d49557` | `bfef92407b49c20492ba3a6b991bbe6ada9678f50c4d7f241865239870e35fd7` | 261,184 |
| many-small / 4,095 | 256 | 318,464 / `e5c90470074e936cff06ad73be993acea1017bce11042208c472ab770fb8af48` | `fda5453a18bb3ea66d1125e98f8e5f80f73b6138e868f84313ef8ea34c9204a8` | 265,216 |
| wide-root / 36 | 2,048 | 2,510,848 / `3bc17d57a6504792a3b04d47db237891ddd6615daaa512c6aa05d833427a56f3` | `576db38a7fe2ab2dea96350bcc49e7f4aa2dfd0fc51648419fc9b6f07d2f337f` | 2,096,192 |
| wide-root / 4,095 | 2,048 | 2,514,944 / `5d02a3b1486d676f2db3b4c43fc8698957a24fcd2b347ec4beba8e64d021efcb` | `eba8bf108edb885a295970d3e4e368773530e49ba9c22e142451aafbbdec9434` | 2,100,224 |

## Correctness and liveness evidence

Every leg produced all 24 records with the expected target output lengths and
SHA-256 values. The overlapping selector returned two isolated copies of the
exact target payload on every sample; the output of one caller cannot mutate
the other caller's payload. Each report checked the public
`SharedOleFile::source_version` before and after the operation and verified the
typed missing-stream refusal. The successful direct handoff leaves the root
MiniStream unmaterialized in the candidate's concurrent records.

The production tests cover the relevant state transitions and failure paths,
including:

- `concurrent_eligible_opens_share_one_direct_range_without_cache` and
  `same_sid_singleflight_covers_36_and_4095_with_aliases_and_isolation` for
  successful 36-byte and 4,095-byte handoff, aliases, and payload isolation;
- `delayed_same_sid_waiter_and_different_sid_takeover_converge`,
  `force_cache_requests_before_a_delayed_direct_read_finishes`, and
  `failed_cache_takeover_during_direct_read_leaves_retryable_cache_state` for
  different-SID and cache-takeover ordering;
- `failed_direct_marker_wakes_same_sid_waiter_for_a_retry`,
  `persistent_io_failure_with_multiple_waiters_designates_retries`,
  `persistent_structural_failure_with_multiple_waiters_designates_retries`,
  `completed_without_handoff_designates_a_retry_owner`,
  `same_sid_source_change_wakes_waiter_without_cache_or_retry_read`, and
  `same_sid_structural_failure_wakes_waiter_without_cache` for failure,
  source-change, and retry liveness;
- `direct_owner_unwind_wakes_same_sid_waiter_via_raii` and
  `delayed_old_epoch_waiter_drop_does_not_decrement_new_handoff` for unwind
  cleanup and the ABA-generation guard; and
- `poisoned_singleflight_state_returns_a_typed_error` and
  `cache_request_published_before_direct_failure_remains_permanent` for typed
  poison handling and permanent cache policy after takeover.

The owner publishes either a bounded payload or a failure marker and wakes all
waiters. RAII owner/waiter cleanup removes terminal slots, and errors are never
stored as if they were payload data. The single-flight payload is bounded by
the CFB direct MiniFAT limit of 4,095 bytes and is discarded after the owner
and waiters leave. This is a bound on the handoff slot only, not a resident
memory or allocation claim.

## Exact logical source-event contract

The raw reports record each positional source event as
`[offset, requested bytes, returned logical bytes]`. Define:

```text
D = [target_start, L, L]   # one eligible direct MiniFAT range
C = [512, R, R]             # complete root MiniStream cache read
```

The exact values are:

| Shape / target | `D` | `C` |
|---|---|---|
| many-small / 36 | `[261632,36,36]` | `[512,261184,261184]` |
| many-small / 4,095 | `[261632,4095,4095]` | `[512,265216,265216]` |
| wide-root / 36 | `[2096640,36,36]` | `[512,2096192,2096192]` |
| wide-root / 4,095 | `[2096640,4095,4095]` | `[512,2100224,2100224]` |

The unchanged selectors preserve the existing policy in both revisions:

```text
one-shot:        [D]
repeat-3:        [D,D,D]
repeat-8:        [D,D,D,D,D,D,D,D]
different-SID:   [D,C,0]
public bulk:     [C]
```

Here `0` means that the third A-B-A invocation performs no additional source
read; it is not a zero-byte physical-I/O observation. The exact same-target
overlap is the new handoff case. Normalizing source completion order, the
500-sample pattern counts per leg were:

| Shape / target | A1 control | B1 candidate | B2 candidate | A2 control |
|---|---|---|---|---|
| many-small / 36 | `DD 323`, `CD/DC 177` | `DD 275`, `D 225` | `DD 235`, `D 265` | `DD 301`, `CD/DC 199` |
| many-small / 4,095 | `DD 259`, `CD/DC 241` | `DD 227`, `D 273` | `DD 213`, `D 287` | `DD 191`, `CD/DC 309` |
| wide-root / 36 | `DD 419`, `CD/DC 81` | `DD 408`, `D 92` | `DD 407`, `D 93` | `DD 440`, `CD/DC 60` |
| wide-root / 4,095 | `DD 326`, `CD/DC 174` | `DD 361`, `D 139` | `DD 347`, `D 153` | `DD 308`, `CD/DC 192` |

`DD` is `[D,D]`; `D` is one recorded direct range for the two returned
payloads; and `CD/DC` is either `[C,D]` or `[D,C]`, with the one `D,C` sample
included in that normalized count. Thus every final candidate overlap sample
is direct-only (`D` or `DD`) and has no root-cache diagnostic, while the
control can take the pre-existing root-cache path when the second caller
misses the direct handoff. These are logical positional-source events and
their scheduling patterns, not physical read, syscall, device, or storage
traffic counts.

Summing the recorded `logical_read_calls` across the four overlap selectors
and both control legs gives 8,000 calls; the two candidate legs give 6,473,
19.09% fewer. Per candidate leg and shape/target cell, the reduction spans
9.2% to 28.7%. This is an accepted source-event accounting result for this
matrix only. It is not a physical-I/O reduction, a syscall count, or a latency
claim.

## Timing disposition: retained, not accepted

`elapsed_ns` measures a fresh validated parser open plus the bounded operation;
corpus construction, refusal checks, source-version snapshots, hashes, and
output validation are excluded. The following table is included to make the
final ABBA result auditable. It reports candidate improvement percentages for
aggregate `elapsed_ns` (`positive = faster`), with `A1 control vs B1
candidate / A2 control vs B2 candidate` in each cell. Percentiles use the
runner's midpoint p50 and nearest-rank p95/p99.

| Concurrent selector | Total p50 | Total p95 | Total p99 | Total mean |
|---|---:|---:|---:|---:|
| many-small / 36 | +5.61% / +5.11% | +5.32% / -0.14% | +33.77% / -7.73% | +6.60% / +2.34% |
| many-small / 4,095 | +5.23% / +4.99% | +5.89% / +0.55% | +3.58% / -22.47% | +4.83% / +2.23% |
| wide-root / 36 | +0.07% / +2.16% | +2.10% / +3.79% | -0.65% / +0.56% | +1.08% / +1.97% |
| wide-root / 4,095 | +5.56% / +6.99% | +7.67% / +6.92% | +7.52% / +6.61% | +3.54% / +4.68% |

These values are not a release latency claim. The many-small tails reverse in
the paired direction, the wide-root/36 p50 and p99 are near the local noise
floor or reverse, and the apparent wide-root/4,095 agreement is one local
selector whose source-pattern mix and parser-open cost remain host- and
scheduler-dependent. Only the candidate's `cfg(test)` source-event acceptance
and tests differ; release harness execution is identical. The configured
range-simulation fields were not used by these production selectors. The final
statistics audit also reports operation-only `Q` p50
changes from +1.64% to +19.07% and aggregate `T` p50 changes from +0.07% to
+6.99% for the concurrent cells; these ranges are retained for audit only.
The sequential cells include operation-only regressions above 5%, so they do
not establish a one-shot or repeat speedup. No one-shot, repeat,
per-invocation cache-hit, bulk, concurrent generic, or local wall-clock
speedup is accepted. The source-event contract is the accepted result;
timing requires a future identical-harness, isolated study.

Using a consistent `100 × (control - candidate) / control` formula, a 5%
paired-regression trigger, and a 5% same-implementation drift trigger, the
independent statistics audit found 161 paired metric regressions and 256
same-implementation drift observations requiring review. The largest paired
regressions were operation p99 -80.67% for the wide-root 36-byte one-shot cell,
operation p95 -54.84% for the many-small 4,095-byte repeat-3 cell, and
operation p95 -52.25% for the many-small 4,095-byte repeat-8 cell, all in the
A2-control versus B2-candidate direction. These review triggers reinforce the
timing withholding; they do not change the accepted source-event result.

## Existing mutex and resource-accounting boundaries

The lock-free revision removes the new single-flight slot-mutex acquisition
from the uncontended direct path. It does not remove or bypass the pre-existing
`ministream: Mutex<Option<Arc<[u8]>>>`, which still serializes root MiniStream
initialization and cache reads after a different target, force-cache request,
bulk request, or cache takeover. Source I/O is not held under the
single-flight state mutex, and the two mutexes are independent; this is a lock
ordering design property, not evidence that root-cache contention disappeared.

The bounded handoff payload is at most 4,095 bytes, but the pre-existing root
cache allocation remains outside bulk `Resource::Memory` accounting. This
record therefore makes no allocation-count, allocated-byte, resident-memory,
RSS, peak-memory, or whole-reader bounded-memory claim.

## Artifacts and claim boundary

The [machine-readable summary](../results/cfb-singleflight-abba-0152-summary.json)
has SHA-256
`83acb616e6de05b119b0e52fc39ed9eb669171519a1a0779d7d0f1b84877cc36`.
The complete final raw reports are retained as:

- [A1 control](../results/cfb-singleflight-a1-control-0152.json.zst),
  compressed size 223,074 bytes, SHA-256
  `61bb22a3c9243ff569144ef7ec7a043b69750b6f49fc4b252c0774750ee29457`,
  decompressed size 261,394,368 bytes, raw SHA-256
  `453bd490a1ae65ef072c6d2dd3010d39b3d4315a421ad615a0a11e03f4fa9507`;
- [B1 candidate](../results/cfb-singleflight-b1-candidate-0152.json.zst),
  compressed size 220,083 bytes, SHA-256
  `bd1038cd06ab317387f6a7d575c39fb55c1ed3366ec171a024048c8610b21b2e`,
  decompressed size 261,199,209 bytes, raw SHA-256
  `35f412e8cde7e5138587ded4802acbb34a1f8c7c39c541f7b79053c2d3b9e5fb`;
- [B2 candidate](../results/cfb-singleflight-b2-candidate-0152.json.zst),
  compressed size 223,702 bytes, SHA-256
  `9709cc092bf4f7383b9a2387faeba5f824f77f7a2ba5d1769f99e465c158072f`,
  decompressed size 261,182,987 bytes, raw SHA-256
  `8358bbcc94c08cfa91bd1db7de3d31418b272678ba8d28cd00e7e01290bd93a2`;
- [A2 control](../results/cfb-singleflight-a2-control-0152.json.zst),
  compressed size 220,132 bytes, SHA-256
  `e77182de24d3c17cb6a8018d0bb3bdb173f67ba34c7d7af7695a79e9644717d9`,
  decompressed size 261,395,675 bytes, raw SHA-256
  `446d7b5764c18ea0f1f7f01863e63970cb35dec197ac62ea953fe9542cf7e47e`.

There is no claim for physical I/O or cold storage, allocation/RSS/peak
memory, device or remote latency, decompression, native DOC/XLS/PPT semantics,
OOXML, ODF, RTF, iWork, or formats outside the named production CFB/OLE2
selectors.
