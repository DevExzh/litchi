# XLS commit editor reuse

Date: 2026-08-11

Production base: `cb0e07e44f5acea96e6a60cf766cca988830bc04`

Scope: native BIFF8 XLS cell-value publication only. iWork/IWA crates were
explicitly excluded.

## Hypothesis

A changed `cell_values::Transaction::commit` parsed the rewritten Workbook
stream before installing it, then rendered the package and opened a new
`Snapshot` from the resulting bytes. Snapshot construction repeated that same
owner parse, reopened and recaptured the CFB package, independently opened the
complete public `Workbook`, and finally performed typed readback. The first
owner parse was discarded, and the second package open recaptured state already
held by the editor after `put_stream_shared` had rendered and reopened the
candidate.

Reusing that already validated editor for snapshot construction should remove
the discarded BIFF parse and redundant CFB open/capture without weakening any
publication gate.

## Change

`Snapshot::from_bytes` now delegates to a private constructor that accepts an
already-open `PackageEditor`. Changed commit passes the editor produced by
`put_stream_shared` directly to that constructor and no longer performs the
discarded preliminary `parse_workbook_stream` call.

The following gates remain in the same changed-publication path:

- `put_stream_shared` checks the package, renders it, reopens the candidate CFB,
  checks protection, and recaptures every stream before returning;
- snapshot construction renders the final exact bytes, runs the offset-bearing
  cell-owner parser once, and independently runs `Workbook::new` over the
  complete final CFB;
- every mutable worksheet must still be published by that complete reader;
- fixed-width, structural, and resource readback remains mandatory;
- exact-source forward patching, stale-source refusal, exact inverse bytes,
  one-stream diagnostics, and the exact no-op fast path are unchanged.

No public API, dependency edge, format capability, resource limit, executor,
or durability contract changed.

## Matched latency measurement

Both binaries use the identical harness at the production base. The before
binary SHA-256 is
`5831033b14ac4f288c15cc4ceb57678ad8dbb4581fc3bc3fa7368fc44b22351a`;
the after binary SHA-256 is
`edb3a8100249e7e17ec761e149b5bd77e8ba6061d6168b4e34ebaa041db45b05`.

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator, CPU 2 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic large XLS input is 163,840 bytes
with 8,192 numeric cells, archive SHA-256
`228c6585a4d26141aebfaf7b08844a2ee445b269d406006a1fdb0484619120fb`,
and Workbook-stream SHA-256
`f806d23f52c978f5215b05fd232b055725a2605d52122ea74ce0cec357ea9386`.

The primary ABBA run used 50 warmups and 500 samples per leg. Pooling the two
legs gives 1,000 raw samples per state; pooled statistics are recomputed from
the samples rather than from leg medians.

| Large XLS one edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 1,776.672 us | 1,639.463 us | **-7.72%** |
| p95 | 1,994.977 us | 1,828.267 us | **-8.36%** |
| mean | 1,800.965 us | 1,658.614 us | **-7.90%** |

The approximate independent-sample 95% interval for the mean delta is
`[-8.45%, -7.36%]`. Matched A and B p50 comparisons improved by 9.34% and
6.15%, respectively. Within-state p50 drift was 0.75% before and 2.75% after.

Raw primary reports and their SHA-256 digests:

- `abba-xls-commit-reuse-one-edit-before-a.json`:
  `a09f24950ab110119722bdf4783f35ffc622564bf6c579779b26249921e6c793`
- `abba-xls-commit-reuse-one-edit-before-b.json`:
  `6e4ddd4559340bae70701f0964c1a44e3ef660bea7392d505d54e5bd06edc60b`
- `abba-xls-commit-reuse-one-edit-after-a.json`:
  `58031e44fb86da68fd3f08b2fcfe36f84331ff4a32c999079b9774ed387234c0`
- `abba-xls-commit-reuse-one-edit-after-b.json`:
  `e0ab09703eea7b05f068bdb68cffc5ddd4fc32c1553c9501b0b41277f316d24a`

## Guardrails

Independent large-input ABBA runs used 30 warmups and 250 samples per leg for
open and full scan. The microsecond no-op path was repeated with 50 warmups and
1,000 samples per leg after the smaller run showed temporal drift.

| Guardrail | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| Open | 1,415.234 us | 1,429.694 us | +1.02% | +2.02% | +4.33% |
| Exact no-op edit/save | 3.225 us | 3.144 us | -2.50% | -4.64% | -2.48% |
| Full cell scan | 90.886 us | 86.068 us | -5.30% | -7.34% | -10.49% |

The changed branch cannot execute in these guardrails. The long no-op legs
still moved materially with time in opposite directions, but their pooled
result is neutral-to-improved and the absolute p50 delta is 0.081 us. The open
guard remains below the 5% review threshold. Raw reports are the
`abba-xls-commit-reuse-{open,noop,full-scan}-*.json` files beside the primary
reports.

Before and after binaries also passed one-sample tiny runs of all four cases.
The harness verified the complete logical grid, exact no-op bytes, changed
bytes, forward patch, inverse restoration, diagnostics, and full snapshot
reopen. Both states reported identical input archive and Workbook-stream
hashes.

## Allocations, RSS, and hardware counters

Matched Heaptrack processes used two warmups and 20 samples of the same large
one-edit case. These are whole-process totals and include the exhaustive
post-timing verifier in both states:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 244,180 | 241,283 | -1.19% |
| Temporary allocations | 23,633 | 22,575 | -4.48% |
| Peak heap | 8.12 MiB | 8.12 MiB | unchanged |
| Heaptrack RSS | 19.89 MiB | 20.28 MiB | +1.96% |

Uninstrumented GNU Time ABBA runs used ten warmups and 200 samples per leg.
Maximum RSS was 30,848/30,976 KiB before and 30,848/30,848 KiB after, so there
is no measured RSS regression.

Matched `perf stat` ABBA runs over the same 210 iterations per leg reported:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 6,689 ms | 6,219 ms | -7.03% |
| cycles | 32,680,532,461 | 30,163,194,039 | -7.70% |
| instructions | 93,656,036,340 | 93,072,436,540 | -0.62% |
| branches | 25,715,008,984 | 25,609,891,276 | -0.41% |
| branch misses | 111,456,214 | 115,814,383 | +3.91% |
| cache references | 9,149,571,883 | 11,125,325,220 | +21.59% |
| cache misses | 124,006,733 | 125,681,850 | +1.35% |
| page faults | 78,472 | 115,451 | +47.12% |
| CPU migrations | 0 | 0 | unchanged |

The two greater-than-5% counter movements were reviewed. Cache references rose
while cache misses stayed within 1.35%; the resulting miss ratio improved from
1.36% to 1.13%, alongside lower cycles and task time. Process-wide minor page
faults rose, but Heaptrack allocation calls and temporary allocations fell,
peak heap was unchanged, and uninstrumented peak RSS did not increase. These
events include process startup and four post-timing snapshot validations per
sample, so they do not contradict the direct latency, cycle, allocation, and
RSS evidence for the changed commit path. They remain recorded as follow-up
guardrails rather than being omitted.

Raw evidence is in `perf-xls-commit-reuse-*.csv`,
`time-xls-commit-reuse-*.txt`, and
`heaptrack-xls-commit-reuse-{before,after}.txt`.

## Correctness verification

- focused `cell_values` library tests: 23 passed;
- real-XLS `xls_cell_values` integration tests: 5 passed, including independent
  `Workbook::new` reopen, exact other-stream preservation, stale patch refusal,
  and exact inverse restoration;
- complete `litchi-xls --all-features` test and doctest suite passed;
- warning-denied production-library XLS clippy and the unchanged benchmark
  harness's 23 tests and warning-denied clippy passed;
- `git diff --check` and formatting checks passed.

The broader warning-denied `litchi-xls --all-targets` clippy command remains
blocked by eight pre-existing warnings in unrelated test code (needless borrows,
two module-inception lints, and two default-then-field assignments). They were
not folded into this performance batch.

The final full reader and typed readback remain the publication boundary. The
next XLS optimization should target a different source of whole-workbook work,
not remove either retained validation layer.
