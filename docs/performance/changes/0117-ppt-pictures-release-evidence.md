# Change 0117: native PPT lazy `Pictures` release evidence

Date: 2026-08-15

Status: accepted as current-revision correctness and positional-read evidence;
latency comparison rejected by the predeclared stability gates.

## Scope and identity

This tranche measures the eight opt-in native PPT `Pictures` selectors from
change 0116. It covers four distinct phases and never adds or geometrically
combines their statistics:

- presentation open without an image query;
- a cold all-`images()` query after untimed open;
- a cached all-`images()` query after untimed open and first query; and
- fresh package open, presentation open, and all-images query in one interval.

The locked release binary was built from revision
`56ba4a0962c398b8e66f6a466074a2504657aeb4`. Its SHA-256 is
`d9dcb380aae9d8222269a7d1407f8612601aac825521d8e5892c179f0f454209`
and its size is 36,735,408 bytes. The measured harness source matched HEAD.
The worktree also contained three unrelated pre-existing user paths plus the
measurement-owned draft documentation/output paths listed in the compact
result. This is qualified dirty-worktree provenance, not a cryptographic
clean-source binding.

The generated in-memory corpus has eight slides and 32 distinct 256 KiB PNG
records. The 8,469,504-byte archive SHA-256 is
`4aeab2d71b21ed721a8638e1f483d87f548cdc690bc1f0998b400eb9df52edbf`;
the 8,389,408-byte `Pictures` stream SHA-256 is
`fe505e1229365db20385c55df519fcc4f5e8d8d628e82fd6dbdb1c135ac9705a`;
the ordered semantic SHA-256 is
`259a62ebab9639393fe73710c1c838c02416d221ed49c4b88399544f80b06516`.

## Measurement protocol

Every selector ran in a fresh child pinned with `taskset`. Each phase used both
`eager-source-source-eager` and `source-eager-eager-source` orderings. Timed
source-backed samples used uninstrumented immutable
`litchi_core::OwnedSource`. Separate untimed `InstrumentedSource` replays
produced the source-read vectors after timing had finished.

The acceptance policy was fixed before interpreting the results:

- both orderings must agree;
- the p50 effect must exceed 5% to be material;
- same-implementation leg drift must be at most 5% for p50 and 10% for p95;
- p99 is not claimed; and
- different phase medians must not be added or pooled.

The first attempt used CPU 2, ten warmups, 100 samples per child, and no
cooldown. It was rejected for 10–13% eager p50 drift in every phase. The raw
files are in [`ppt-pictures-0117/`](../results/ppt-pictures-0117/).

The extended attempt used CPU 8, twenty warmups, 200 samples per child, eight
fresh children per phase, and a one-second cooldown before each child. Its raw
files are in
[`ppt-pictures-0117-stable/`](../results/ppt-pictures-0117-stable/). The compact
machine-readable analysis is
[`ppt-pictures-release-0117.json`](../results/ppt-pictures-release-0117.json).
Each raw JSON retains its sorted elapsed vector, environment, corpus manifest,
and independent counter evidence; each adjacent `.time.txt` retains the
whole-process GNU `time -v` observation.

The compact result binds each raw directory with SHA-256 over every file in
lexical filename order, updating the digest with the UTF-8 filename, one NUL
byte, and the exact file bytes. This is a deterministic content manifest, not a
Merkle tree or filesystem-metadata hash.

Representative child command:

```sh
/usr/bin/time -v -o <leg>.time.txt \
  taskset -c 8 tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 20 --samples 200 --case <one-selector> --json <leg>.json
```

The summary's deterministic bootstrap uses Python `random.Random(0x117)` and
20,000 independent with-replacement resamples over each ordering's two pooled
200-sample legs. Each replicate stores
`median(source) / median(eager) - 1`; the interval uses the sorted 500th and
19,500th values. Those intervals describe the stored distributions; they do
not override a failed drift gate.

## Exact read evidence

Corpus preflight passed the exact and one-byte-under package/stream gates.
Every timed and untimed replay then used the exact finite limits and passed the
semantic digest checks. The source-backed replay vectors were constant across
every sample:

| Phase | Total reads / bytes | Generated `Pictures` window reads / bytes |
|---|---:|---:|
| Open, no query | 136 / 79,265 | 0 / 0 |
| Cold all-images query after open | 1 / 8,389,408 | 1 / 8,389,408 |
| Cached all-images query | 0 / 0 | 0 / 0 |
| Fresh open plus all images | 137 / 8,468,673 | 1 / 8,389,408 |

The overlap counter is valid only for the generated fixture's contiguous
`Pictures` payload window. It is not a general CFB sector-chain map or private
materialization counter.

## Latency observations and rejection

The table reports pooled p50 deltas for source-backed relative to eager in the
extended attempt. Negative values favor source-backed. They are retained as
observations, not accepted performance results.

| Phase | Forward p50 delta | Reverse p50 delta | Failed stability evidence |
|---|---:|---:|---|
| Open, no query | -76.5510% | -76.6779% | eager p50 9.91%; eager/source p95 13.57%/14.78% |
| Cold all-images query | +3,875.3310% | +3,797.5518% | eager p50/p95 8.42%/15.59% |
| Cached all-images query | +17.8950% | +22.2918% | source p50/p95 10.86%/12.89% |
| Fresh open plus all images | +43.7331% | +54.1203% | eager/source p50 8.36%/13.53%; source p95 27.12% |

Both directions show that laziness moves work from open into the first image
query, while cached queries cause no further source reads. The timing magnitude
is not accepted because every phase failed at least one same-implementation
stability gate. The record does not call the open observation a general PPT
speedup, nor call the first-query observation an end-to-end regression.

GNU `time -v` maximum RSS ranged from 91,136 to 91,584 KiB in the extended
children. This is whole-process RSS including startup, corpus generation, and
untimed evidence replays; it is not per-operation allocation, peak-live-byte,
or memory-regression evidence.

## Limitations

No accepted latency, allocation, peak-memory, cold-filesystem, remote-range,
decompression, recompression, memory-copy, hardware-counter, or save-path
claim is made. Producer breadth, arbitrary CFB layouts, signed/encrypted inputs,
and end-to-end edit/publication remain outside this generated-corpus record.
iWork remains excluded while its crates are changed independently.
