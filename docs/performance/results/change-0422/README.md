# Operation-region allocator baseline

This batch adds an independently initialized operation-region maximum of
absolute process live allocation requests. The prior corrected lifetime peak
can be dominated by setup. [Design and boundaries](design.md) explains the
serialized observer, concurrent callback ordering and remaining limits.

`protocol.json` was frozen before build/capture. It selects the same owned
media-rich and plain PPTX lifecycles as 0421: CPU 2, one worker, two fresh
processes per selector, 30 samples after three warmups, Rust 1.98.1, warm
generated inputs, and the established release/debug/frame-pointer flags.
The shared host's background load is uncontrolled. This is a current baseline,
with no cross-version allocator timing or memory-reduction claim.

The final source revision, clean before/after source manifests and both binary
identities are in `build-candidate.json`. Each run retains source/corpus/output
bindings, semantic/preservation/refusal gates, all allocation vectors, stdout,
stderr and GNU time. `summary.json` and `result-table.md` retain all vector
statistics, including region peaks. New reports identify
`serialized_region_peak_v3`; historical v2 reports retain their separate
process-counter meaning and cannot be paired under one policy.

## Replay

From the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0422/summarize.py --replay
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0422/check-report-guards.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0422/check-portable-tool-binding.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0422/portable-replay.py
```

The bundle includes four shared validator modules in `replay-tools/`, pinned
by their manifest hashes before export. Portable replay needs no original
checkout or binary. It replays four baseline reports, probes actual-report
invariants, and validates a separate normal identity report. The normal report
has one sample and no warmup; it is a functional check excluded from baseline
measurements. A standalone bundle test also rejects a changed validator.

`SHA256SUMS` inventories every retained artifact. `compression.json` preserves
raw logs byte-for-byte with original and stored hashes. `cleanup.json` records
batch worktree/binary removal while preserving shared build caches.

## Fresh capture

Create a clean detached worktree at the candidate revision. Copy the frozen
protocol into a fresh output directory. Invoke the retained scripts with Python
and `PYTHONDONTWRITEBYTECODE=1`:

```text
build.py candidate WORKTREE --root OUTPUT --binary-prefix /tmp/UNIQUE_PREFIX
capture.py --root OUTPUT
summarize.py --root OUTPUT
```

Fresh builds use the existing 0418 build/source helpers and shared target
cache. Keep builds, tests and workloads serialized. Drivers reject existing
outputs. Remove only reproduction-owned worktrees and binaries afterward.

## Scope

The region peak includes live bytes present at entry and every callback that
linearizes inside the interval, including other process threads. Operation
workers must complete before finish to include all their work. This is a
logical callback-order observation, excluding allocator-internal realloc
copy overlap and physical RSS. The mutex can perturb scheduling; instrumented
elapsed time does not support a performance claim. Normal binary execution
retains its system allocator without this wrapper.

Explicit retained-object/drop boundaries, near-limit cases, matched
source-backed media lifecycles, and the full non-iWork performance program
remain open. This observer is measurement infrastructure, not an optimization
or a bounded aggregate-memory guarantee.
