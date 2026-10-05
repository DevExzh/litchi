# Change 0918 DOC ownership comparison

This packet defines a frozen comparison between the retained 0917 production
build and an isolated candidate that changes the changed-stream ownership in
`DOCRevisionEditor::commit` from `Arc::from(Vec<u8>)` to
`Arc::from(&[u8])`. The candidate has no public API or retention-policy
change. The input document identity is shared by both versions; their
production source lineages are recorded separately in the freeze manifest:
the retained baseline is built from `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`
and the candidate is built from the current `f5485b3ca8582051247d31946b1c5f45c8a561ad`
line with its isolated patch receipt.

The prospective protocol is frozen in
[`comparison-protocol.json`](comparison-protocol.json). It requires 192
native children: twelve paired rounds over both real DOC fixtures and edit
counts 1, 2, 8, and 32, with baseline/candidate order alternating and fixture
and count order reversed on alternating rounds. Native children use default
retention, five warmups, and fifty measured samples. The allocator lane has 96
children: both versions, both repetitions, all three retention routes, both
fixtures, and all four edit counts, with zero warmups and one sample.

Every child runs the unchanged 0917 probe and the unchanged 0917 validator.
The validator requires all twenty correctness gates, including exact output
and source controls, public semantic reread, patch inverse/refusal controls,
untouched stream and directory preservation, retention transitions, and the
allocator retained-output rule. A paired reduction additionally requires the
input source hash, output identities, capacities, route controls, and stage
observations to match across versions.

The candidate is eligible for adoption only if all of these conditions hold:

* at least one 32-edit default-retention fixture has at least five percent
  lower median `allocated_bytes`;
* no allocator cell has an `allocated_bytes` increase above one percent;
* no allocator cell has a `peak_live_bytes` increase above five percent; and
* every native default-retention fixture/count pair has a paired bootstrap
  lower bound of at least `1/1.05` for baseline p50 divided by candidate p50.

The reduction preserves p50, p95, p99, RSS, allocation, and peak-live adverse
flags. A failed gate rejects adoption even if the candidate remains a small
source diff.

## Commands

Run a pre-freeze qualification with the same CLI for either version. This is a
0-warmup/1-sample correctness check and is not part of the performance
schedule:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B comparison.py qualify \
  --version baseline --lane native \
  --binary /external/baseline/doc_retention_probe \
  --case docnohf --edits 1 --retention default --warmup 0 --samples 1 \
  --qualification-root /external/0918-qualification
```

Qualification writes an exclusive child directory containing raw stdout,
stderr, GNU `time` output, the exact argv, timestamps, exit status, binary
descriptor, fixture identity, and the validated report summary. Reusing a
qualification label is refused, so the coordinator can retain every failed
attempt while preparing the frozen capture manifest.

Freeze only after source/build receipts, both binaries, and the unchanged
0917 probe have passed review:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B comparison.py freeze \
  --baseline-native /external/baseline/doc_retention_probe \
  --baseline-allocator /external/baseline/doc_retention_probe_alloc \
  --candidate-native /external/candidate/doc_retention_probe \
  --candidate-allocator /external/candidate/doc_retention_probe_alloc \
  --baseline-source-receipt ../change-0917/receipts/source-receipt.json \
  --baseline-build-receipt ../change-0917/receipts/build-receipt.json \
  --candidate-source-receipt /external/candidate/source-receipt.json \
  --candidate-build-receipt /external/candidate/build-receipt.json \
  --candidate-patch candidate.patch
```

The candidate source receipt must report revision
`f5485b3ca8582051247d31946b1c5f45c8a561ad` and carry a
`candidate_patch`/`patch` descriptor, or a `source_hashes.candidate_patch`
value. The supplied patch file must have the same SHA-256; the freeze copies
that small descriptor into the packet. This binds the three-line change to
the candidate build without treating the input document hash as a source
code identity.

The coordinator starts the two lanes only after all profiling, build, quality,
and review work is idle:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B comparison.py run --lane native
PYTHONDONTWRITEBYTECODE=1 python3 -B comparison.py run --lane allocator
PYTHONDONTWRITEBYTECODE=1 python3 -B comparison.py reduce
PYTHONDONTWRITEBYTECODE=1 python3 -B comparison-replay.py
```

Replay of a copied comparison packet uses the same reconstructed repository
layout as earlier packets. The packet must be at
`<root>/docs/performance/results/change-0918`, with copied
`docs/performance/results/change-0917/capture.py`, the complete unchanged 0917
`probe/` tree, and both fixture files below `<root>`. The replay command checks
these paths before loading the runner and never falls back to the original
checkout. Qualification replay has a separate scope: its receipts retain
absolute external raw, fixture, and binary paths, so `qualification-replay.py`
requires those original custody paths.

The runner refuses to overwrite a child or receipt, enforces the schedule
prefix, retains failed attempts, and refuses reduction until every expected
child passes. Uncompressed raw stdout, stderr, and GNU `time` logs stay in the
external capture root. Each packet artifact is a deterministic gzip with both
compressed and decompressed descriptors; replay verifies the copied packet
gzip bytes and does not trust an original absolute raw path. Executables and
raw artifacts at or above one MiB remain outside Git.

No performance capture or adoption result belongs to this source-only setup
until the coordinator freezes the identity manifest and starts the schedule.
