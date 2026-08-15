# Change 0142: CFB atomic-save phase attribution

Date: 2026-08-15

Status: accepted as current-revision timing and exact logical-source-work
attribution. This record does not claim an optimization speedup.

## Scope

The isolated `cfb_file_same_length_overlay_atomic_save` selector now records
three non-overlapping intervals inside its existing operation timer:

1. filesystem source plus validated `SharedOleFile` open;
2. same-length overlay planning, composed-CFB reopen and selected-stream
   validation;
3. synced sibling staging, atomic publication and final pre-rename validation.

Each interval records elapsed nanoseconds and logical `ReadAt` calls, requested
bytes and returned bytes. The parent fails closed unless the three counter
deltas sum exactly to the existing whole-operation counters and their elapsed
sum does not exceed the outer timer. The harness-only implementation is commit
`3c23ab40dd9a2fd01875ef501f9723771cf06a7e`; no `litchi-cfb` production code or
fingerprint invariant changed.

## Reproducible release capture

The release binary was built from a clean detached checkout of that commit
with Rust 1.95.0. Its SHA-256 is
`c56d190f6a70c2bfcce4969f50268c89ccdbb6cd1954b0f8b0a35c0006c286e0`
and its size is 39,382,224 bytes. The run was pinned to CPU 2 on the existing
AMD EPYC 9575F / Linux 6.8.0-101-generic / ext2/ext3 host:

```sh
taskset -c 2 litchi-perf-baseline \
  --warmup 20 --samples 200 \
  --filesystem-cache warm,cold-requested \
  --case cfb_file_same_length_overlay_atomic_save \
  --json cfb-save-phase-current-0142.json
```

Every measured sample ran in a fresh child. The deterministic five-entry,
few-large CFB is 16,913,408 bytes (source SHA-256
`7ffbd37c3e472a21b382bcbb02e430a62164e58d2270bbee0deaa584ff47a94d`).
Every sample changed one 36-byte MiniFAT stream, reported one physical span,
published 16,913,408 bytes, and produced SHA-256
`7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc`.

| State / interval | p50 | p95 | p99 | Mean |
|---|---:|---:|---:|---:|
| warm total | 138,153,550 ns | 149,454,932 ns | 154,633,791 ns | 138,053,676 ns |
| warm open | 311,740 ns | 387,815 ns | 434,820 ns | 319,233 ns |
| warm plan | 33,442,779 ns | 34,591,265 ns | 35,140,524 ns | 33,521,696 ns |
| warm atomic publication | 103,842,832 ns | 115,819,221 ns | 120,665,794 ns | 104,157,172 ns |
| cold-requested total | 135,319,622 ns | 164,646,346 ns | 169,500,975 ns | 142,085,434 ns |
| cold-requested open | 1,418,851 ns | 3,173,056 ns | 3,759,775 ns | 1,795,068 ns |
| cold-requested plan | 46,936,548 ns | 50,052,043 ns | 51,816,022 ns | 47,389,805 ns |
| cold-requested atomic publication | 86,794,070 ns | 114,218,409 ns | 119,152,697 ns | 92,845,925 ns |

Phase percentiles are computed independently and are not additive.
`cold-requested` means accepted `posix_fadvise(DONTNEED)` advice, not proven
physical cold storage.

## Exact source-work attribution

The logical counters were identical in all 400 samples and both cache states:

| Interval | Calls | Requested/returned bytes | Share of operation bytes |
|---|---:|---:|---:|
| open | 264 | 135,680 | 0.1599% |
| plan and candidate validation | 784 | 33,962,596 | 40.0321% |
| atomic publication | 777 | 50,740,224 | 59.8080% |
| **operation** | **1,825** | **84,838,500** | **100%** |

This locates the remaining read amplification after Change 0103: complete
fingerprint scans dominate planning and publication, while CFB open contributes
little logical byte volume. The scans remain required for stable-token source
mutation defense. A larger fingerprint-only request window is therefore a
bounded candidate for matched A/B measurement; removing another fingerprint
stage is not.

The [compact machine-readable record](../results/cfb-save-phase-current-0142-summary.json)
contains all 400 total and phase latency observations, exact counters, process
vectors, environment, corpus and output identities. Its SHA-256 is
`23667a7d518b98959d80a987e62bfcb2e46d67f2c24d2cac900f4dd77cc3de5d`.
The [zstd-compressed full capture](../results/cfb-save-phase-current-0142.json.zst)
has SHA-256
`3a032f4c0aee67ac735b01dcb4d147ba56abbf8332f5be09ff6cea999c84b991`;
after decompression the 14,651,444-byte JSON has SHA-256
`3281a398f9b430b9e9b7a01a9b0db65349dfc32f38c0b0cb901e0963dd1e7a5e`.

## Claim boundary

This is one current revision, not matched before/after evidence. It establishes
phase attribution and exact logical source work only. Logical `ReadAt` bytes
are not physical device I/O, decompression, copying or memory-bandwidth
measurements. The procfs RSS high-water field is process-lifetime VmHWM, not an
operation peak. No allocation, peak-memory, physical-cold, remote-source,
DOC/XLS/PPT semantic CRUD or optimization speedup claim follows.
