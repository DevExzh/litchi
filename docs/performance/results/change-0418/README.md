# Change 0418: owned PPTX cross-copy candidate reuse

The matched media-rich owned lifecycle has 38.27–38.69% lower p50 elapsed
time, with an approximately 8% increase in whole-process peak RSS. The
[change record](../../changes/0418-pptx-cross-copy-candidate-reuse.md),
[result table](result-table.md) and [memory review](memory-review.md) describe
the scope and retention decision. The full non-iWork performance goal remains
open.

## Inputs and boundaries

The measured control is `79dfee5025276676d433a80b0b5475ae03f09db8`, and the
candidate is `f8f9e6667284ae28fdbf9b9313b2ff9583ec74fc`. Both are retained in
branch history. They share the lifecycle harness and independent reader/test
corrections; the production difference is candidate reuse. `build-identity.json`
binds source files, compiler flags, clean worktrees and all four binaries.

The host is an AMD EPYC 9R45 KVM guest, Linux 7.0.0-1011-aws, with 32 logical
CPUs. Rust 1.98.1 release builds retain debug information, frame pointers and
unwind tables. Captures use CPU 2 and one worker. Build, test, measurement and
profile workloads were serialized; unrelated activity on the shared host is
uncontrolled. Full host and command metadata remain in the journals.

`protocol-initial.json` preserves the initial declaration. `protocol.json`
clarifies flags and diagnostic lanes before building or capturing. Normal
ABBA runs retain 6,400 observations; the separate allocator lane retains 480,
and 32 one-sample preflights exercise all selector/lane/leg combinations.
Each selector/leg runs in a fresh process, with warmups and samples sharing
that process. This is warm generated in-memory evidence.

The primary and plain lifecycle timers include owned ingress, snapshots,
planning, atomic application and sequential publication. Input Vec clones,
sink reservation, reopen/verification and teardown are outside. The old phase
timers retain their original sum-of-phases boundary. The old media phase run
has 100 samples per leg and is a diagnostic below the strict registry minimum.
Allocator elapsed is excluded from latency comparison. Operation allocation
is available only on the two lifecycle selectors; live and high-water values
are global snapshots, and GNU time RSS covers the whole process.

## Replay

From the repository root, with Python 3 and `zstd` available:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0418/verify.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0418/render-summary.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0418/profile/summarize.py --replay
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0418/negative-probes.py
```

`verify.py` recomputes all raw statistics and allocation/RSS projections,
checks source/binary/catalog/output identities and validates refusal and
preservation gates. Its default mode compares retained projections exactly;
`--write` explicitly regenerates them. Compressed sidecars represent unchanged
raw bytes and preserve the journals' raw hashes. Evidence decoding is bounded.
`SHA256SUMS` inventories the stored files and can be checked with `sha256sum -c`
from this directory.

The negative probe script copies only its required evidence and tools to
temporary exports. It first replays an unmodified export, then an explicitly
path-rewritten export with regenerated projections. Twelve independent
corruptions repair outer hashes where appropriate and must fail semantic
checks even with projection regeneration enabled. Temporary exports are
removed by the script.

The exact capture and profiling commands are retained in `capture.json` and
`profile.json`; the build/capture/profile drivers are under `scripts/`.
Reproduction requires clean worktrees at the two measured revisions and fresh
output paths. Historical temporary worktrees and copied binaries are removed
after evidence validation. Shared Cargo target caches are retained.

## Profiles and limits

The separate CPU and PMU runs each use 20 samples after three warmups per role.
CPU recording uses `cycles:u`, 999 Hz and frame-pointer call chains. The
no-inline header/self/children/stack exports and ELF `.comment` records were
captured while the exact binaries were available. `profile/analyze.py --write`
performs that export; re-symbolization requires the matching binaries.
`profile/summarize.py --replay` uses retained text and metadata and remains
usable after binary/worktree cleanup.

Profiles cover the whole command, including setup, warmups, verification and
reporting. Event periods are weights, not elapsed phase shares. Hardware
counters retain runtime and multiplex percentages. Raw zero or unsupported
aliases do not prove zero cache activity. No timer-local IPC, physical-I/O,
operation-peak, full retained-output memory or native Office acceptance claim
is made.

`latency/` contains the primary selector's canonical strict ABBA package.
The canonical registry policy uses 5%/5%/10%/15% drift ceilings for
p50/mean/p95/p99; the batch's `summary.json` additionally enforces the declared
5% ceiling on every statistic. All primary statistics pass both. The registry
claim is latency-only and excludes the diagnostic phase selector and memory
metrics.

The full candidate PPTX/OPC tests pass (1,270 tests, three ignored); the matched
PPTX control passes 828 tests with two ignored. `checks/` retains initial
failures and successful reruns. Unqualified Clippy has three unchanged lint
findings; the candidate passes with exactly the documented command-scoped
exemptions. The default 36-case/198-row matrix and coverage-index statuses are
unchanged; two opt-in lifecycle additions bring the selector registry to 427.
