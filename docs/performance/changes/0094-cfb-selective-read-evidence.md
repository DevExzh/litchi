# Change 0094: CFB selective-read evidence controls

Date: 2026-08-14

## Scope

This change promotes `SharedOleFile::read_stream_range` as a public bounded
caller-buffer API and adds four opt-in `tools/perf-baseline` selectors for the
CFB substrate. Canonical open retains the already validated root-mini-stream
sector index so MiniFAT range chunks use O(1) ordinal lookup without caching
payload bytes or repeating FAT-chain walks. It does not change a semantic DOC,
XLS, or PPT API, and it is not included in the default matrix.

The selectors are paired legacy-cursor full-stream and positional-source
exact-range controls for two deterministic targets:

- `cfb_selective_mini_legacy_read` and
  `cfb_selective_mini_shared_read`: a 36-byte MiniFAT stream;
- `cfb_selective_fat_legacy_read` and `cfb_selective_fat_shared_read`: a 4 MiB
  FAT stream.

Each target is the final named stream among either 256 (`many-small`) or 2,048
(`wide-root`) siblings. All non-target streams are deterministic 1 KiB
incompressible payloads. The generated archive is shared by the matched legacy
and positional controls for each target/shape cell. The manifest records the
exact archive SHA-256, target SHA-256, target length, entry count, logical
payload byte count, generator, shape, and target path.

## Timed and recorded evidence

Each measured sample records separate `open_ns`, `read_ns`, and `total_ns`
segments. Open and selected-stream read counters are reset between stages. The
legacy case materializes the full stream. The positional case allocates an
exact-length caller buffer inside the read stage and fills it through the
bounded range API without populating the root-mini-stream cache. Both
implementations use instrumented source adapters, so the report records
stage-local read calls, returned bytes, and sorted selected-read range sizes.
It also records the selected payload's SHA-256 and the exact returned payload
byte count. This is not presented as a cache-occupancy metric. `sink` is
explicitly `none`; no output or save path is exercised.

Corpus construction, deterministic archive validation, and payload hash checks
are outside the timed stages. The selected bytes must equal the manifest target
payload on every sample, and the hash/returned-byte values must remain
stable across samples. The selectors emit no performance conclusion: they are
evidence for later release ABBA, allocation, and peak-memory attribution.

## Gates and exclusions

The focused tests verify deterministic generation for both shapes and target
kinds, exact target lengths and logical byte totals, archive/target hashes, and
the presence of nonzero open/read I/O plus stage timing arrays for both reader
implementations. The default 36 cases / 198 records remain unchanged; the four
new names increase the selectable case-name count only. iWork remains deferred.

## Accepted release evidence

The compact result is
`docs/performance/results/cfb-selective-range-abba-0106-summary.json`. Two
separate release binaries were run in pinned order before-A, after-A, after-B,
before-B on CPU 0, with 30 warmups and 500 samples per cell. The summary retains
the two binary hashes, corpus hashes, p50/p95/p99 observations, and source
counters; raw local outputs were not committed.

For the 36-byte MiniFAT target, the explicit positional range read reduced one
physical request from 261,184 to 36 bytes on 256 siblings and from 2,096,192 to
36 bytes on 2,048 siblings. Read-stage p50 moved from 9.24-9.82 us to 0.48 us
and from 82.6-84.3 us to 0.65-0.67 us, respectively. Both ABBA directions agree.
End-to-end p50, which includes complete mandatory CFB open/validation, moved by
about 8.4-14.2% and 6.6-11.9%; this is deliberately not described as an
order-of-magnitude end-to-end result.

The 4 MiB FAT control retained exactly one 4 MiB physical request. Its paired
read and total p50 changes stayed within 5% control drift. FAT p95/p99,
MiniFAT p99, cold filesystem, simulated high-latency range, allocation, and
peak RSS conclusions remain open.
