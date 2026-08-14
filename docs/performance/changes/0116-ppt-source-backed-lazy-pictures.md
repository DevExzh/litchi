# Change 0116: native PPT lazy `Pictures` harness evidence

Date: 2026-08-15

Status: Accepted as opt-in harness and source-read correctness evidence. No
latency, allocation, memory, cold-filesystem, or release-ABBA result is
claimed.

## Scope and corpus identity

The standalone `tools/perf-baseline` harness now exposes eight selectors:

* `ppt_pictures_eager_open` and `ppt_pictures_source_backed_open`
* `ppt_pictures_eager_first_image` and `ppt_pictures_source_backed_first_image`
* `ppt_pictures_eager_cached_repeat` and
  `ppt_pictures_source_backed_cached_repeat`
* `ppt_pictures_eager_open_all_images` and
  `ppt_pictures_source_backed_open_all_images`

The default 36-case/198-record matrix is unchanged. iWork remains outside this
tranche while the `iwa-*` crates are being changed separately.

Each selector uses the fixed corpus generator
`litchi-ppt-pictures-lazy-v1`: eight generated slides with 32 distinct,
deterministic PNG records, each exactly 256 KiB. The manifest records the
archive SHA-256, the `Pictures` stream byte length and SHA-256, the generator
identity, and entry count/size. Per-result source evidence records the ordered
canonical semantic SHA-256. The semantic digest includes each image's index,
checked PNG kind, payload length, and payload bytes; it is emitted for both
eager and source-backed results.

The one-sample debug smoke produced a `Pictures` stream of 8,389,408 bytes
with SHA-256
`fe505e1229365db20385c55df519fcc4f5e8d8d628e82fd6dbdb1c135ac9705a` and the
canonical semantic digest
`259a62ebab9639393fe73710c1c838c02416d221ed49c4b88399544f80b06516`. These
identifiers are report evidence for this generated corpus, not a claim about
other PPT producers. The same smoke's 8,469,504-byte archive has SHA-256
`4aeab2d71b21ed721a8638e1f483d87f548cdc690bc1f0998b400eb9df52edbf`.
Its source-backed open observed 136 `ReadAt` calls/79,265 bytes and zero
`Pictures` overlap; first query observed one call/8,389,408 bytes in both
total and overlap counters; the repeated query observed zero calls/bytes in
both counters. These are one-sample debug counters only.

Both implementations use identical finite `RecordLimits`. Corpus construction
executes untimed exact-limit and one-byte-under package/`Pictures` gates. The
source-backed one-under `Pictures` refusal also proves zero reads overlapping
the actual `Pictures` payload range.

## Timing and source evidence

Open selectors time package construction plus presentation open, without an
image query. Cold-query selectors construct/open outside the interval and time
one `images()` call, which returns all 32 images. Repeated-query selectors
perform one untimed verified query, then time the second query. The fresh
open-all-images selectors directly time package open, presentation open, and
one all-images query in the same interval, so no phase medians need to be
added. Warm-up samples are excluded from result vectors.

Source-backed latency samples use immutable, uninstrumented
`litchi_core::OwnedSource`; construction of that prepared source is outside
the intervals. Eager paths likewise construct their borrowed `Cursor` before
the timer. Harness counter atomics therefore do not bias only one side of a
latency comparison. The JSON identifies the timed and evidence adapters.

After all elapsed samples finish, the source summary runs independent untimed
`InstrumentedSource` replays and reports their total `ReadAt` calls/bytes plus
calls/bytes overlapping the contiguous `Pictures` payload window in this
generated fixture. The gates require zero overlap during open, the full stream
byte count during a cold query, zero additional reads during a cached query,
and one full stream overlap during fresh first use. These are not counters from
the timed instances, a general CFB sector-map, or an internal materialization
count. Eager results intentionally leave counter vectors empty.

The harness verifies all queried images against the canonical semantic digest,
but it does not infer a speedup or a private cache/materialization event from
these counters. A later clean, balanced release ABBA run with retained raw
samples and resource evidence is required for any performance claim.

## Verification

The focused harness test builds the corpus twice, checks deterministic archive
and manifest identity, runs all eight selectors, and asserts the phase-specific
source-range evidence. The locked harness check/test and formatting/diff checks
are the tranche gates; no files are staged or committed by this change.
