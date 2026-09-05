# Owned OPC payload storage reuse

This bundle compares clean control `68e4b0dfa` with the candidate revision
recorded in `build-candidate.json`. The harness is identical. The candidate
adds an OPC constructor that chooses an existing payload allocation only after
fully validating and decoding the new archive. Exact URI, content type,
visible donor bytes, decoded target bytes, and no-greater vector capacity gate
reuse. Target metadata, relationships and source authorization come from the
new archive. PPTX cross-copy uses this constructor for clean owned destinations.

The [result table](result-table.md) and [resource review](resource-review.md)
record the measured decision. Normal ABBA processes use 100 samples and ten
warmups; separate allocator processes use 30 samples and three warmups. These
are resource diagnostics, below the 500-sample release latency threshold.
Allocator elapsed times are excluded. No new latency claim is registered.

Both builds use Rust 1.98.1 and identical release/debug/frame-pointer flags.
Each measurement uses CPU 2 and one worker on the AMD EPYC 9R45 KVM host.
Inputs are generated in memory and warm; host background activity is
uncontrolled. Build identities and run journals retain exact commands, source
manifests, binary hashes, corpus/output hashes and environment information.

The [source review](source-review.md) accounts for payload holders. The prior
0419 candidate Heaptrack capture supplies the profile basis because its
production and harness sources match this control; `host-and-profile-basis.json`
records that binding. No new heap trace is captured. Full archive decompression
still occurs before storage reuse, so this experiment does not eliminate that
work or establish reduced physical copying. Live and high-water allocator
fields are process snapshots; RSS includes setup, gates and teardown. Neither
is an operation-local peak or an aggregate memory budget.

## Replay

From the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0420/summarize.py --replay
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0420/verify-probes.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0420/capture-selftest.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0420/portable-replay.py
```

The portable replay exports this bundle and four hash-bound Python validators
into a temporary directory. It verifies the retained reports without the
original worktrees or binaries. Both post-cleanup and isolated-export replay
passed; `cleanup.json` records removal of this batch's worktrees and binaries
while preserving shared caches. `SHA256SUMS` inventories retained artifacts.
Compressed validation logs preserve their raw bytes and hashes in
`compression.json`, including the corrected test module's first compile failure.
The first negative-probe fixture incorrectly treated a valid sample-order
permutation as corrupt. Its failed run is retained; the corrected probe keeps
100 entries but duplicates an index, and is rejected along with a wrong output
digest and a false correctness gate. No verifier policy was weakened.

## Fresh measurements

Create detached clean worktrees at the revisions in the two build identities.
Copy `protocol.json` and `measurement-protocol.json` to a fresh output directory.
Run the following scripts from this checkout, with Python and
`PYTHONDONTWRITEBYTECODE=1`:

```text
build.py control CONTROL_WORKTREE --root OUTPUT --binary-prefix /tmp/UNIQUE_PREFIX
build.py candidate CANDIDATE_WORKTREE --root OUTPUT --binary-prefix /tmp/UNIQUE_PREFIX
capture.py --root OUTPUT
summarize.py --root OUTPUT
```

Builds share the existing `tools/perf-baseline/target` cache and copy binaries
before the next build. Serialize all task builds, tests and measured CPU work.
The drivers refuse existing build/capture outputs. Fresh hardware or software
requires fresh identities; results need not reproduce exactly on another host.
Remove only reproduction-owned worktrees and binaries, retaining shared caches.

## ADR compliance

| Constraint | Implementation and evidence |
|---|---|
| 0001: priorities and API layers | Correctness precedes reuse. The additive constructor is in advanced OPC; ordinary PPTX CRUD signatures do not change. |
| 0002, 0010, 0011, 0024: ownership | Storage selection belongs to OPC; PPTX imports no physical ZIP type or new dependency. |
| 0003: immutable snapshots and atomic edits | Shared immutable `Arc<Vec<u8>>` payloads preserve copy-on-write isolation; detached candidate validation and application remain unchanged. |
| 0005: bounded resources and evidence | Existing read limits and archive output limit remain. Larger donor capacities are rejected. Matched resource diagnostics keep live/peak/RSS scope explicit. |
| 0006: preservation and security | New bytes undergo existing complete validation and decompression. No donor metadata or authority is imported. Exact output, typed refusals and malicious donor storage are tested. |
| 0008: verification | OPC/PPTX unit, integration and doc tests, focused ownership tests, Clippy, rustdoc and repository gates are retained. |

The [goal audit](goal-audit.md) keeps the broader non-iWork program open.
Near-limit memory, source-backed lifecycles, native-producer breadth, physical
cold/range I/O and scaling remain separate work.
