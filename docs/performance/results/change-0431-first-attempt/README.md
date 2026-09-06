# First attempt: verified compressed source-part transfer

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
The 480 baseline samples passed the frozen report verifier. Preliminary
R1/R2 API median differences are below 3.5% in every lane. No after-result or
performance improvement is claimed by this baseline statement.

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

## First-attempt result and disposition

This retained attempt compares the original baseline with implementation
`e68c1decb472e4823c7bea9322f12419ddc51bee`. Both roles passed all 480 samples.
API median milliseconds, R1 / R2:

| Media provider | Before | First after |
|---|---:|---:|
| Bytes | 252.841 / 257.666 | 32.783 / 28.373 |
| Warm file | 255.177 / 255.196 | 34.814 / 35.020 |
| Short range | 263.248 / 254.163 | 29.373 / 29.448 |
| Delayed range | 654.019 / 660.341 | 719.306 / 718.339 |

The delayed-range result regressed 8.8–10%. Compressed capture added
16,786,581 source bytes and 1,193 publication source calls in the bytes/file/
delayed-range reports (including layout/verification reads); before publication
reported zero source calls at that phase. Capture used 16 KiB chunks. The
[main batch](../change-0431/README.md) refines that capture granularity
to bounded 64 KiB reads and retains the completed matched after protocol. This first
attempt is retained as evidence, not accepted as the final cross-provider result.

Plain API medians changed by approximately -1.8% to -0.005%. The after bytes
lane has a 13.45% repeat-median review trigger. Media process peak RSS stayed
within 783–785 MiB across both roles. Endpoint RSS varies substantially,
including in the baseline before measured API calls; endpoint differences do
not establish a causal publication memory change. The comparison retains all
resource changes, including the intended additional input and verification work.

`initial-portable` passed sealed portable replay and mutation checks. Preserve
these raw observations when assessing the refined result. No allocator-count,
physical-copy-count, cold-cache, network, or scaling claim follows from them.

The refreshed `precleanup-portable` and `aftercleanup-portable` replays also
passed, including isolated-copy and mutation checks. The main bundle retains
task scratch cleanup inventory; this historical replay needs no binary.
