# 0917 final capture review

Final read-only audit, 2026-10-05.  The packet is bound to source revision
`bb19ca74bd52044fbd2c1c7cc9875e975bb334c0` and the frozen capture script
SHA-256 `fde5219539cf538d51b8bf92ec5527fd9959ac0f39eb2ea5412b8ddf8cd7a894`.
I did not build, recapture, or modify the frozen scripts, protocol, manifest,
raw evidence, receipts, or reduced result.

## Raw matrix and reduction

The packet contains 144 native child receipts with 30 samples each (4,320
samples total) and 48 allocator child receipts with one sample each.  I
independently reconstructed the exact six-round route pairing from
`native-protocol.json`, checked every receipt schema/status/spec binding, and
verified every retained stdout descriptor and report hash.

I recomputed nearest-rank p50/p95/p99 and means for every child, pooled route
statistics, RSS comparisons, the 16 native route flag records, all 24
allocator cells and 32 allocator route comparisons.  I regenerated every
native bootstrap with 10,000 draws, seed 917, matched-round ratios, and
nearest-rank 2.5th/97.5th bounds.  The independently constructed native and
allocator reductions compare equal to every corresponding field in
`reduced.json`.

The descriptive review flags are limited to two native cells: `docfloat`
edit count 2 has default p95/p99 increases of 5.86%/6.48% and release
p95/p99 increases of 8.81%/13.23%.  The only allocator flag is RSS for
`docfloat` edit count 2, repetition 2, default retention (+7.49%).  There are
no native RSS flags, no allocator peak-live/retained/allocated flags, and all
24 allocator repeat cells have equal observations.  These are protocol review
triggers, not adoption gates.

The required post-capture command also passed:

```text
LITCHI_PERF_0917_ROOT=/home/zhuhe/litchi-performance-0917/capture \
python3 -B docs/performance/results/change-0917/analysis.py replay
```

It returned `verified` with 144 native and 48 allocator children.  A separate
temporary output from `analysis.py reduce` matched the frozen reduced result
in schema, protocol, manifest, native reduction, allocator reduction, and
adoption fields.

## Adapter and custody

The original frozen reducer has the recorded tuple-unpacking failure in its
native reduction loop.  `analysis.py` asserts the frozen capture SHA and makes
exactly one in-memory substitution:

```text
for base_item, (_, item) in zip(...) ->
for (_, base_item), (_, item) in zip(...)
```

The old and replacement strings each occur exactly once.  No frozen input or
raw result changed.  The manifest, replay script, protocol, source receipt,
and build receipt hashes all match their retained descriptors.

The source archive is 164,352,000 bytes with SHA-256
`61df35b45519c95b43fb4194b0f679c69ce17aa8b39e5c112c0aea9b02f95863`.
The before-build, after-build, after-quality, and after-capture source
manifests all have SHA-256
`6be84856c37c4729c79968c3616c7a07aa040a136971dba26301d6589e1b3c49`;
the post-capture record reports 9,268 regular files.  All three probe
manifests are identical.  The build receipt passed, all four quality checks
passed, and the functional qualification receipt reports 48/48 children
passed.  The packet has 192 logs and 576 raw files, with no packet file at or
above 1 MiB; binaries, the source archive, and other large evidence remain
external.

## Copied-packet controls

In a temporary copied checkout hierarchy, adapter replay first verified the
relocated packet.  The following mutations were made only in the copy and
were rejected:

* appended data to a retained raw stdout: descriptor content changed;
* changed a child receipt status: failed child retained;
* changed `reduced.json` schema: reduced schema mismatch;
* added `logs/extra.json`: child receipt set differs.

Copied source and build receipts resolved through relative packet descriptors;
mutating the copied source receipt was rejected by its hash check.  The
original packet and external capture tree were unchanged by these controls.

## Disposition

**Approved for the recorded 0917 same-build DOC retention comparison.**  The
result supports only the frozen two-fixture, 1/2/8/32 repeated paragraph-0
workload and its zero/default/release routes.  Host activity outside the
coordinated capture CPUs remains uncontrolled, and review flags remain
descriptive.  No production optimization or adoption claim follows from this
packet.  I am idle.
