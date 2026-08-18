# Change 0190: reusable CFB stream-chain validation scratch

Date: 2026-08-18

## Decision

Retain a private, fallible scratch buffer for each of the MiniFAT and FAT
stream-validation loops in `litchi-cfb::OleFile::open`. The change removes two
short-lived allocations per nonempty stream while preserving the complete CFB
chain, ownership, overlap, cycle, marker, truncation, and physical-layout
validation boundary.

This is generic CFB/OLE2 substrate work. It changes no public API, input or
source abstraction, cache, lock, unsafe-code boundary, writer, publication
path, or DOC/XLS/PPT semantic behavior.

## Mechanism and invariants

Before this change, every nonempty stream called
`collect_sector_chain_exact`, which allocated an exact-capacity `Vec<u32>` and
a checked bitset sized to the complete allocation table. The chain was used
once to claim physical sectors or logical mini-sectors and then discarded.

`validate_stream_allocations` now owns one lazily allocated
`SectorChainScratch` for MiniFAT streams and one for regular FAT streams. Each
scratch retains only a chain vector and a checked visited map. Collection still:

1. handles the zero-length/`ENDOFCHAIN` case first;
2. rejects invalid starts and declared lengths before reserving;
3. reserves the chain before the visited map with the existing
   `sector-chain entries` and `sector-chain map` resource labels;
4. checks conversion, bounds, cycles, intermediate markers, and the exact
   terminal marker in the original order; and
5. completes the same mini-sector or physical-sector ownership claim before
   the next stream is considered.

Scratch state is cleared before every collection and after collection errors.
The retained root mini-stream chain, general FAT/MiniFAT/directory collectors,
global mini-sector ownership map, physical sector roles, and final FAT/layout
reconciliation are unchanged. No input-derived allocation became infallible.

Focused tests cover buffer reuse across different chain lengths, reservation
before a growing walk, reuse after a cycle error, zero-length chains, and exact
early/late termination diagnostics.
The existing allocation-validation suite still opens the complete legacy
Office corpus and covers FAT/DIFAT/MiniFAT overlap and malformed declarations.

## Allocation evidence

A matched whole-process Heaptrack capture ran both deterministic CFB shapes in
one process with three warmups and 100 retained opens per shape:

| Metric | Control | Candidate | Change |
|---|---:|---:|---:|
| Allocation calls | 988,558 | 509,749 | **-48.44%** |
| Heaptrack temporary allocations | 242,178 | 2,567 | **-98.94%** |
| Peak heap | 2.72 MiB | 2.72 MiB | flat at displayed precision |
| Heaptrack RSS | 14.59 MiB | 14.88 MiB | +1.99%, descriptive only |
| Leaked bytes | 544 | 544 | flat |

The control profile attributes exactly 237,312 calls to each of the two
per-stream allocation sites. That is `(3 + 100) * (256 + 2,048)` calls per
site. The candidate retains one bounded scratch buffer per role, growing it
fallibly to the largest encountered chain instead of allocating per stream.

These are process-total profiles: corpus generation, warmups, verification,
and JSON reporting are included. The accepted result is therefore the exact
whole-process allocation-call reduction for this command, supported by the
named stack attribution. It is not an operation-local allocated-byte, RSS, or
native semantic claim.

## Matched release timing

Frozen release binaries ran on CPU 2 in A1 control, B1 candidate, B2 candidate,
A2 control order. Every leg used 200 warmups and 5,000 retained samples per
shape with the Rust system allocator.

| Shape / statistic | A1 | B1 | B2 | A2 | paired reductions |
|---|---:|---:|---:|---:|---:|
| many-small p50 | 144,729 ns | 132,863 ns | 134,306 ns | 135,862 ns | 8.20% / 1.15% |
| many-small mean | 149,391 ns | 135,087 ns | 137,367 ns | 140,397 ns | 9.57% / 2.16% |
| many-small p95 | 178,375 ns | 145,621 ns | 148,934 ns | 160,680 ns | **18.36% / 7.31%** |
| many-small p99 | 195,939 ns | 156,866 ns | 157,336 ns | 189,120 ns | **19.94% / 16.81%** |
| wide-root p50 | 981,284 ns | 959,929 ns | 986,476 ns | 1,001,351 ns | **2.18% / 1.49%** |
| wide-root mean | 1,005,439 ns | 966,252 ns | 1,000,535 ns | 1,023,066 ns | **3.90% / 2.20%** |
| wide-root p95 | 1,160,185 ns | 1,015,095 ns | 1,088,637 ns | 1,164,732 ns | **12.51% / 6.53%** |
| wide-root p99 | 1,273,270 ns | 1,071,241 ns | 1,305,967 ns | 1,357,817 ns | 15.87% / 3.82% |

The accepted latency result is deliberately per statistic. Many-small p95 is
7.31%-18.36% lower and p99 16.81%-19.94% lower; its p50 and mean are withheld
because the control drifts 6.13%/6.02%, above their 5% gates. Wide-root p50 is
1.49%-2.18% lower, mean 2.20%-3.90% lower, and p95 6.53%-12.51% lower; p99 is
withheld because candidate drift is 21.91%, above its 15% gate. Every accepted
statistic agrees in both pairs and both implementations pass that statistic's
predeclared drift gate.

Cold or physical I/O, concurrent allocator contention, peak RSS, native
DOC/XLS/PPT semantics, save, and edit performance remain unmeasured.

## Verification

```text
cargo test --locked -p litchi-cfb reusable_chain_scratch --lib
cargo test --locked -p litchi-cfb allocation_validation --lib
RUSTFLAGS='-D warnings -D deprecated' cargo clippy --locked \
  -p litchi-cfb --all-targets -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-cfb --no-deps
cargo fmt --all -- --check
python3 tools/check_crate_boundaries.py
```

The focused scratch tests pass 5/5 and the allocation-validation suite passes
12/12. Strict Clippy and rustdoc pass. The complete serial CFB suite passes
233/234; the one failure is the existing untouched
`detected_temp_substitution_is_not_deleted_by_cleanup` portable temporary-file
identity regression, which also fails in exact isolation with `ENOENT` and is
outside this read-only-open change.

Artifacts:

- [machine-readable summary](../results/cfb-chain-scratch-0190-summary.json)
- [artifact manifest](../results/cfb-chain-scratch-0190-manifest.json)
- compressed A1/B1/B2/A2 raw reports and before/after Heaptrack captures listed
  in the manifest
