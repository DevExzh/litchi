# Change 0054: ODS shared durable-patch blobs

Date: 2026-08-11

Production control: `df8398132d252de56c4fe2468f090ae35f69cb7d`

Scope: `litchi-core` blob ownership and private ODS durable-patch construction
only. iWork/IWA crates were explicitly excluded.

## Hypothesis and change

An ODS `Patch` already retains its exact immutable source and target packages
as `Arc<[u8]>`. The semantic patch formerly inserted borrowed views of those
same packages into its forward and reverse `BlobBundle`s. Each insert
allocated and copied the complete archive, then the ODS owner computed the
same two SHA-256 fingerprints again for operation preconditions. For the fixed
media-rich case, Heaptrack attributed 33.58 MB of retained payload copies to
the six inserts made by three commits, while `perf` attributed 7.28% of
complete-process cycles to the four package hashes in patch construction.

`BlobBundle::insert_shared` now accepts an immutable `Arc<[u8]>`, applies the
same per-blob, duplicate, count, total-byte and checked-accounting rules, and
retains that allocation directly. ODS clones its already retained package
Arcs into the forward and reverse bundles and reuses the resulting content
addresses as the source and target SHA-256 strings. This removes two complete
archive copies and two of four package SHA passes from each durable patch.

The target bundle is still processed before the source bundle, preserving
limit-error precedence. A duplicate continues to bypass saturated count and
total-byte limits only after the per-blob bound succeeds. No operation,
precondition, BlobId, deterministic JSON, patch direction, inverse ordering,
source check, package byte, or readback result changes. The new low-level API
adds no dependency, runtime, executor, lock, cache, unsafe code, parser
leniency, or format-specific type to `litchi-core`.

## Matched latency evidence

The frozen release binaries have SHA-256:

- control: `1943fff9eff6d34b895f8b66fbd9e3d0978e573090e5d5f3a1ca76c50f7badcf`;
- candidate: `685edf901c08364fbf4045f564c048dbc9a6c5e984c8ff361d263e7e45059929`.

Both use the unchanged standalone harness, release profile, Rust 1.95.0,
Linux 6.8.0-101-generic, the Rust system allocator, and CPU 11 pinned with
`taskset`. The deterministic ODS contains 2,048 cells and eight exact 2 MiB
incompressible media resources. Its 11-member archive is 16,790,689 bytes,
with SHA-256
`46b7f61cb74639115f6d120dc6498b97d6b310d51c78c4fb85ac60d6fc758b14`.
Every iteration changes `Sheet 1!R16C16`, commits, checks changed state,
verifies the complete package/media content, hashes the deterministic result,
and reopens it. Only commit latency is timed.

The primary `ods_media_one_edit_save` measurement used 10 warmups and 50
samples in each of two control/candidate and two candidate/control pairs.
Pooling raw samples gives 200 observations per state while balancing binary
order.

| Media-rich ODS one-cell edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 326.694 ms | 297.958 ms | **-8.80%** |
| mean | 333.797 ms | 303.533 ms | **-9.07%** |
| p95 | 388.832 ms | 334.979 ms | **-13.85%** |
| p99 | 462.518 ms | 465.889 ms | +0.73% |

The approximate independent-sample 95% interval for the mean delta is
`[-10.68%, -7.45%]`. All four matched pairs improve: p50 deltas are -10.28%,
-6.38%, -9.06%, and -11.23%; mean deltas are -8.92%, -6.57%, -8.95%, and
-11.52%. The p99 movement is within the 3% tail gate and is disclosed rather
than hidden in the pooled center.

Medium and large ODS guards used 20 warmups and 100 samples in each A/B/B/A
leg, or 200 observations per state and cell. Complete primary-duration guards
all improve:

| Guard | Before p50 | After p50 | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|---:|---:|
| open, medium | 3.236 ms | 3.113 ms | -3.81% | -4.08% | -6.34% |
| open, large | 50.179 ms | 48.457 ms | -3.43% | -7.87% | -22.26% |
| exact no-op edit/save, medium | 4.352 ms | 4.216 ms | -3.13% | -7.96% | -37.58% |
| exact no-op edit/save, large | 68.692 ms | 65.338 ms | -4.88% | -6.66% | -10.69% |
| one-cell edit/save, medium | 22.642 ms | 21.572 ms | -4.73% | -10.76% | -39.60% |
| one-cell edit/save, large | 349.713 ms | 336.204 ms | -3.86% | -6.73% | -21.98% |

Read-only p50 guards are also neutral or better: sheet listing is 40 -> 20 ns
medium and 60 -> 60 ns large; one-cell lookup is 240 -> 160 ns and 2.263 ->
2.043 us; full cell sweep is 16.873 -> 14.050 us and 470.964 -> 451.513 us;
full cell text is 52.932 -> 46.803 us and 1.621 -> 1.556 ms. These tiny paths
do not execute durable patch construction, so their movement is treated as
environmental guard evidence rather than an optimization claim.

## Attribution and resources

Matched 20-sample `perf record` processes captured 9,622 control and 8,940
candidate samples with no lost-sample warning. Kernel symbols are restricted
on this host, but userspace ODS and SHA frames resolve:

| Cycle attribution | Before | After |
|---|---:|---:|
| all `sha2::sha256::x86_sha::compress` self cycles | 9.19% | 6.06% |
| patch-construction package SHA stacks | 7.28% | 3.99% |
| redundant `DiagnosticFingerprint::of` stack | 3.66% | absent |
| content-address insertion SHA stack | 3.62% | 3.99% |

Matched whole-process `perf stat` A/B/B/A used five warmups and ten measured
commits per leg. Pooling the two legs per state gives:

| Counter | Delta |
|---|---:|
| task clock | -6.51% |
| cycles | -6.21% |
| instructions | -3.25% |
| branches | -1.47% |
| branch misses | -0.16% |
| cache references | -4.78% |
| cache misses | -12.24% |
| page faults | -23.00% |
| CPU migrations | 1 -> 1 per process |

Heaptrack used zero warmups and three measured commits per state:

| Whole-process metric | Before | After |
|---|---:|---:|
| allocation calls | 890,189 | 890,177 |
| temporary allocations | 191,981 | 191,981 |
| peak heap | 142.79 MB | 140.05 MB (-1.92%) |
| Heaptrack RSS | 165.83 MB | 164.38 MB (-0.87%) |
| leaked bytes | 1.78 KB | 1.78 KB |

The control's `BlobBundle::insert -> ODS Patch::build` site accounts for 33.58
MB over six retained payload allocations and is absent in the candidate. The
whole-process peak falls less because the two eliminated packages did not both
overlap the process's other maximum-live allocations.

Uninstrumented GNU Time A/B/B/A reports 158,680-158,696 KiB maximum RSS before
and 158,580-158,692 KiB after: effectively flat and slightly better. Major
faults remain zero; mean minor faults fall from 551,674.5 to 426,368. Raw ABBA
reports, profiles, counter reports, RSS reports, and binary provenance are
indexed by
[`ods-shared-patch-blobs-sha256.txt`](../results/ods-shared-patch-blobs-sha256.txt).

## Correctness and quality gates

New `litchi-core` tests prove exact `Arc` allocation identity, identical
content addresses, byte-identical deterministic reversible-patch JSON against
borrowed insertion, duplicate behavior at saturated bounds, count-before-total
error precedence, per-blob precedence, and failure atomicity. A focused ODS
test proves that forward/inverse semantic bundles retain the exact target and
source package allocations and that their BlobIds exactly match the existing
operation-precondition SHA strings.

Verification completed:

- `litchi-core --all-targets --all-features`: 141 unit tests and every example
  target passed;
- `litchi-ods --all-targets --all-features`: 133 unit tests and every
  integration target passed;
- warning-denied Core all-target/all-feature Clippy and ODS production-library
  Clippy passed, including the deprecation cleanup already committed in
  `1194fbc7f`;
- warning-denied Core/ODS rustdoc passed after one inherited stale ODS
  public-to-removed-model link was converted to plain field documentation;
- all 32 standalone harness tests and warning-denied all-target Clippy passed;
- the Core detection libFuzzer target compiles;
- formatting and `git diff --check` pass.

The broader ODS all-target Clippy gate still reports the unrelated pre-existing
test-only layout/style findings recorded by earlier performance work; this
batch does not modify those tests. No dedicated ODS fuzz manifest exists in
the current tree.

## Remaining work

This ownership handoff only removes duplicate durable-patch package storage and
hashing. It does not alter ODS ZIP publication, content comparison, worksheet
serialization, compactness/audit, security policy, final package reopen, media
verification, operation vocabulary, or durable apply/inverse semantics. The
next bounded RTF parser-block reservation has separate positive p50/p95 and
heap evidence but needs a quiet p99 acceptance run. Native DOC table-state
lookup and wider OOXML multi-Part source publication remain separate tranches.
iWork/IWA stays excluded while its crates are modified independently.
