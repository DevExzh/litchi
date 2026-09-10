# 0492 pilot1 raw review

This is a read-only review of the eight `captures/pilot1` reports. There are
four arms (baseline/candidate at zero delay and delayed transport), each run
with normal and allocator binaries, and each report has three measured rows:
24 observed rows in total. The rows are raw observations; pilot1 is not an
accepted analysis or a formal performance claim because of the custody issue
below.

All 24 rows report `docx_provider_lifecycle_v2`, requested source revision
`e44a23396146d504ffc738e0989de896635f02a3`, unchanged source versions, and a
verified 10,000-byte text result with SHA-256
`ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af`. The
corpus is 16,793,036 bytes with SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.
Each terminal receipt says pass with exit code 0 and cleanup passed.

## The 24 observed rows

`logical` and `physical` are shown as `calls / requested bytes`; every call
returned its full request. The candidate read-ahead counters are identical in
each of its rows: 19 logical requests, 16 hits, 3 misses/fills, and 5,445
fill-requested and fill-returned bytes in three physical calls.

| role | transport arm | rows | latency samples (ns, 1..3) | logical | physical | physical media overlap (requested/returned bytes) |
| --- | --- | ---: | --- | --- | --- | --- |
| normal | baseline, 0 us | 3 | 308531, 290362, 284691 | 19 / 3966 | 19 / 3966 | 0 / 0 |
| normal | candidate, 0 us, 4096-byte window | 3 | 298061, 297371, 284412 | 19 / 3966 | 3 / 5445 | 69 / 69 |
| normal | baseline, 1000 us + 104857600 B/s minimum service | 3 | 20836250, 20381789, 20345019 | 19 / 3966 | 19 / 3966 | 0 / 0 |
| normal | candidate, 1000 us + 104857600 B/s minimum service, 4096-byte window | 3 | 3481365, 3465875, 3464525 | 19 / 3966 | 3 / 5445 | 69 / 69 |
| allocator | baseline, 0 us | 3 | 319351, 313512, 302641 | 19 / 3966 | 19 / 3966 | 0 / 0 |
| allocator | candidate, 0 us, 4096-byte window | 3 | 308142, 309132, 296672 | 19 / 3966 | 3 / 5445 | 69 / 69 |
| allocator | baseline, 1000 us + 104857600 B/s minimum service | 3 | 20361108, 20335078, 20353338 | 19 / 3966 | 19 / 3966 | 0 / 0 |
| allocator | candidate, 1000 us + 104857600 B/s minimum service, 4096-byte window | 3 | 3509944, 3480614, 3487103 | 19 / 3966 | 3 / 5445 | 69 / 69 |

The nonempty logical trace is the same in all 24 rows:

```text
(16793014,22), (16791709,46), (16791709,1305), (0,30), (49,363),
(412,16), (412,16), (428,30), (469,234), (703,16), (703,16),
(2907,30), (2965,324), (3289,16), (3289,16), (1420,30),
(1467,1424), (2891,16), (2891,16)
```

The baseline physical trace is that same 19-range trace. The candidate
physical trace is:

```text
(16793014,22), (16791709,1327), (0,4096)
```

The 69 bytes of candidate physical media overlap are compressed ZIP-member
overlap caused by the larger fills. The logical wrapper reports zero media
overlap. This does not show that any media member was decompressed.

With only three rows per arm, the zero-delay normal median moves from 290,362
ns to 297,371 ns (+2.4%), while the delayed normal median moves from
20,381,789 ns to 3,465,875 ns (−83.0%). The corresponding allocator medians
are 313,512 ns to 308,142 ns (−1.7%) and 20,353,338 ns to 3,487,103 ns
(−82.9%). These medians describe this synthetic, warm/recent-file pilot and
are not evidence for a production optimization.

## Custody limitation

Direct inspection of the raw reports and terminal receipts passes the text,
trace, source-version, and counter checks. Strict canonical collection still
rejects pilot1: its captured `started.json` and `terminal.json` source receipt
spells the manifest path as

```text
/home/zhuhe/code/litchi/docs/performance/results/change-0492/validation-sources/18fbdb0df7adc7a2482fbd9e458bcabd7d79581f8842c8e01038e6aeabb07a2b.json
```

where the canonical protocol and retained build receipts spell the same
manifest as

```text
validation-sources/18fbdb0df7adc7a2482fbd9e458bcabd7d79581f8842c8e01038e6aeabb07a2b.json
```

The SHA-256 and 7,159-file count agree, but path spelling is part of the
receipt identity. Therefore pilot1 must remain raw evidence and must not be
used as an accepted `analysis/pilot1` or verification result. The repaired
driver should recapture pilot2 and formal1 with one canonical representation;
rewriting pilot1 after capture would not repair its custody record.

## Methods review and production next step

`methods.md` accurately limits this work to an unmanaged benchmark adapter,
the v2 lifecycle timer, synthetic transport counters, compressed-byte media
overlap, and descriptive pilot/formal comparisons. The report should retain
three implementation qualifications: the shared `MEDIA_SCOPE` wording was
written around the logical wrapper even where the physical proof records
69-byte overlap; baseline rows have no read-ahead allocation despite the
candidate-oriented setup wording; and logical counters/traces contain
nonempty `ReadAt` requests because empty buffers are skipped. A read-ahead
hit means the requested start is in the cached window, so a crossing request
also needs the adapter's returned-prefix semantics when interpreting hits.

After pilot2/formal1 has canonical custody, the concrete production step is a
private, opt-in `ArchiveReadAhead` at the managed `litchi-opc`
`SourceReader`/`SourceSnapshot` boundary for DOCX semantic reads. Keep
`litchi_core::ReadAt`, generic `soapberry-zip`, exact-read paths, and
publication/splice paths unchanged. The managed implementation must reserve
the finite window as `Resource::Memory`, charge each physical fill as
`Resource::InputBytes`, perform cancellation/work checks, preserve source
identity and revision checks, and map source changes to the existing typed
error. Add budget, cancellation, short/zero-fill, mutation, concurrency,
differential, and exact-read negative tests before a new custody-complete
before/after evidence bundle is considered for deployment.
