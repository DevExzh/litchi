# Verified compressed source-part transfer

This batch implements ZIP-reader-issued compressed payloads for source-backed
OPC part addition and PPTX cross-slide publication. Change 0430 attributed
83.22% (bytes) and 83.38% (warm file) of iteration-ancestry CPU samples to
publication's Deflate intersection. Those are sampled ancestry proportions,
not API wall-time fractions or predicted speedups.

The hypothesis is that retaining the source's verified compressed media avoids
recompression during publication. Planning and publication still perform their
existing decoded payload, XML, dependency, and candidate checks. New members
receive canonical wrappers; untouched destination members retain their existing
physical preservation contract. Source authority, cancellation, budgets, CRC,
ZIP layout, and partial-output handling remain correctness requirements.

## Matched measurement protocol

`protocol.json` was frozen before source edits. Each role runs 16 serial fresh
processes: plain/media-rich corpora, each over bytes, warm files, a caller range
source capped at 4096 bytes without delay, and a caller range source capped at
65536 bytes with a fixed 200 microsecond delay. R1 uses forward lane order; R2
reverses it. Each process performs three warmups and retains 30 samples on CPU 2.

The API duration is the sum of source open, destination open, planning, and
publication clocks. Corpus generation, source copies, sink reservation, gates,
resource observations, and drops are outside these clocks. See the retained
report clock policy for exact exclusions. File inputs were recently staged and
hash-read: these are warm-file measurements. Range delay is caller simulation,
not an actual network measurement. Process peak RSS includes setup and all
untimed work; lifecycle RSS is observed at endpoints rather than continuously.

The before executable is the unchanged 0429/0430 executable, copied and hashed
before capture. `build-before.json` binds it to the retained source manifest.
Both final roles passed all 480 samples and the frozen report verifier.
The final candidate is `0556401e21f1ba740ff033f67c4d4d2151741b8a`;
`build-after.json` binds its release binary and unchanged harness source.

## Evidence limits

[source-oracle-audit.md](source-oracle-audit.md) describes the independent
producer semantic and preservation gates and the limits of portable replay.
Output hashes may change for added wrappers; input identities and semantic
acceptance must remain matched. The benchmark harness source must remain
identical across role manifests.

Historical build metadata is retained in `input-build-0429.json`; prior machine
metadata is in `machine-prior.json`. The Rust/TOML/lock manifests do not cover
every compile-time fixture or template. Recapture therefore requires the
external repository and its assets; this bundle is not a hermetic build kit.
No claim is made about allocator counts, physical memory-copy counts, cold
caches, native PPTX cross-slide copying, concurrent scaling, or actual remote
networks.

## Matched result

API median milliseconds, R1 / R2:

| Provider | Plain before | Plain after | Media before | Media after |
|---|---:|---:|---:|---:|
| Bytes | 2.242 / 2.252 | 2.244 / 2.251 | 252.841 / 257.666 | 28.294 / 28.263 |
| Warm file | 2.561 / 2.555 | 2.534 / 2.540 | 255.177 / 255.196 | 34.008 / 34.393 |
| Short range | 2.248 / 2.252 | 2.237 / 2.246 | 263.248 / 254.163 | 29.694 / 29.791 |
| Delayed range | 122.324 / 122.327 | 122.323 / 122.321 | 654.019 / 660.341 | 527.453 / 527.441 |

The synthetic media-rich workload records 86.5–89.0% lower API medians for
bytes, warm files and short ranges, and 19.4–20.1% lower for the simulated
delayed range. Plain changes range from -1.045% to +0.091%. No API median
regression or repeat difference exceeds the 5% review threshold. Full p95/p99,
means, descriptive bootstrap intervals and throughput remain in `comparison.json`.

The [first attempt](../change-0431-first-attempt/README.md) retains an
8.8–10% delayed-range regression. It captured compressed bytes in 16 KiB
requests. The refined 64 KiB requests reduce publication source calls from
1,193 to 425 for bytes/file/delayed providers; the 4 KiB capped provider stays
at 4,265. `refinement-check.json` confirms identical input, output identity,
and charged input/work across both candidate captures. Both candidates add
16,786,581 source input bytes and 33,559,592 charged source work units relative
to baseline. Verification still decodes and compares the logical payload.

Media process peak RSS remains 783.2–785.3 MiB across baseline and refined
captures, with no >5% peak regression. Lifecycle endpoint RSS varies with
already differing baseline residency. Nine repeat flags are baseline-role
short-range RSS observations. The 625 general regression triggers include
expected input/work accounting and endpoint residency; they are not 625 API
latency regressions. No causal memory reduction or full-heap bound is claimed.

## Validation and replay

The refined affected suite passed 1,755 tests with five ignored. Strict Clippy,
affected formatting and another 1,000-run AddressSanitizer fuzz smoke passed.
Rustdoc, workspace/default-feature checks and ownership-boundary checks are
retained from the implementation validation. Development failures and the
unchanged unrelated formatting debt are retained in `checks/`.

`verify.py --portable-check` replays the source-bound report checks and
comparison in an isolated copy and rejects numeric, output-identity and bound
verifier mutations. `seal.py` records deterministic lossless log compression
and a complete hash inventory. Replay validates retained reports and their
producer gates; it does not independently reopen exported output archives.
The broader non-iWork goal remains active; see [goal-scope.md](goal-scope.md).

Both `precleanup-portable` and `aftercleanup-portable` passed. The latter runs
after removal of the captured binaries and task drafts; no executable is
needed for portable replay. `cleanup.json` retains exact removed-file hashes.
