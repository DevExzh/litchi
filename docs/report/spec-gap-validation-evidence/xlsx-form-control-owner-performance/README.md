# XLSX form-control owner performance evidence

This directory contains candidate-only evidence for the owner-review gate. The
original top-level receipts remain pinned to the frozen pre-correction source
gate and are preserved as historical evidence. New replays write a fresh
directory under [`runs/`](runs/) by default, so a replay never overwrites those
receipts.

The top-level `raw-receipts.jsonl`, receipt index, sanity file, and paired
provenance/hash files are the earlier source-bound candidate run; their hashes
and contents are intentionally unchanged. Use the latest run directory when
reading the live-byte allocator fields or the corrected harness provenance.

The frozen source snapshot predates the later canonical-SML entity-escaped-URI
correction reported by owner review. These artifacts therefore remain a
pre-correction smoke until the corrected source gate is restored and replayed.
Owner review currently holds source-v2 on duplicate VML attributes, diagnostic
projection-byte undercharge, and name-limit `Resource` error labeling; this
historical run does not certify those corrections.
The measurement contract is in
[`performance-plan.md`](performance-plan.md). The retained source delta,
source bundle, source lock, harness lock, gate, and replay instructions are
kept in this directory.

[`restore-source.sh`](restore-source.sh) reconstructs a bounded source tree
from the base commit, applies [`source-delta.patch`](source-delta.patch),
installs [`source-Cargo.lock`](source-Cargo.lock), and verifies the resulting
files against [`source-root-gate.json`](source-root-gate.json),
[`source-manifest.sha256`](source-manifest.sha256), and
[`source-bundle.tar`](source-bundle.tar). It selects workspace manifests,
crate sources/build inputs, pinned Cargo/toolchain/lint configuration, and the
retained form-control fixtures so replay does not copy the repository's
unrelated large test corpus.

The harness verifies the full restored-source manifest, build configuration,
harness manifest, source/harness locks, compiler identity, and relevant build
environment before and after build and collection, and records those checks in
each run directory. The release binary hash is also recorded before and after
collection. Temporary build
targets and generated packages are removed by the harness trap.

Each new raw receipt records cumulative `requested_event_bytes`, a
`live_bytes_delta` from the allocator's pre-phase live-byte baseline, and
`peak_live_bytes`, the maximum live-byte increase above that baseline during
the phase. These live-byte fields account for successful alloc, dealloc, and
realloc deltas across phase boundaries. They are allocator-observation
metrics, not physical heap or RSS proofs; RSS remains a separate process
snapshot field. The historical top-level receipts retain their original
`peak_requested_event_bytes` field and are not rewritten or re-described as
live-heap measurements.

The smoke records eager and source-backed package-open and owner-projection
lanes separately, single and many-control queries, iteration, cheap clones,
typed caps, source read counters, exact selected-property bytes, and a
deterministically generated package with a 1 MiB unrelated opaque member. It
does not establish a runtime baseline or a speedup claim.

## Independent root replay

`runs/root-replay` independently restored and verified all 5,840 pinned source files, build configurations and locks, then completed eight correctness checks and 560 measurement receipts. Source and binary hashes were unchanged across the run; fixture and member hashes matched the profiler replay. The owned restore, harness, build target and synthetic fixture directories were removed after success. `root-verification.json` records this check. This remains historical pre-correction evidence and does not approve the current owner fixes.

The retained source patch keeps required unified-diff context spaces, and the historical build log retains its terminal blank line. These raw replay artifacts are intentionally unchanged; whitespace checking excludes those two files.
