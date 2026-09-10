# DOCX positional-read locality review

This is a bounded simulation for the 0492 read-ahead pilot. It is not a
timed result and it is not evidence that the formal 0491 provider process used
the exact offsets below.

## Evidence and provenance

The sealed 0491 formal provider evidence was built at source revision
`59b509a7f9508778b20b428d733fc107ec22f142` and was subsequently recorded in
commit `e44a23396146d504ffc738e0989de896635f02a3`. It uses the pinned 0188 corpus
(`16,793,036` bytes, SHA-256
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`).
The normal `file` and delayed-range rows both report, for each of their 30
samples in each repeat, 19 logical calls and 3,966 requested and returned
bytes. The delayed range arm is
`range-65536-1000us-104857600bps-minimum-service`; its fixed delay is paid
once per delegated range call. The formal summary is in
`analysis/provider-formal1.json` (SHA-256
`9c0c8933a664c986330788ca2de70910ce16958bd577155ffb7f6e827e0575fa`).

The offsets come from the separately retained source replay diagnostic, whose
recipe is `source-replay-diagnostic.rs.txt` (SHA-256
`ec5a909e8970a32dc027c167d2fe5102cadc5ec96e105df336046b7c54affd2b`). Its
successful run is receipt `validation/source-replay-diagnostic-run2.json`;
the exact JSON trace is `validation/source-replay-diagnostic-run2.stdout`
(1,088 bytes, SHA-256
`fc3d71a9e11ef1d2af7fda437828f25492ee831ddd4633c90b1c53b87a6f0b0e`). The
replay source reports 18 calls during package open and 5 during document
preparation, including four zero-length calls. Its 19 nonempty calls match the
formal provider's aggregate 19-call/3,966-byte shape:

| phase | all calls | zero-length calls | nonempty calls | nonempty bytes |
| --- | ---: | ---: | ---: | ---: |
| open | 18 | 3 | 15 | 2,480 |
| prepare | 5 | 1 | 4 | 1,486 |
| query | 0 | 0 | 0 | 0 |
| total | 23 | 4 | 19 | 3,966 |

The replay is useful offset evidence, but it is not a formal-provider trace:
the small standalone program uses `Package::from_read_at` with its own
`ReadAt` recorder and source version `(491, 0)`. It does not put the
`PptxRangeSource` transport adapter or the explicit limits/cache construction
used by `docx_provider_lifecycle` around the source. The formal report stores
aggregate range counters, while the sealed replay stores offsets. Therefore
the matching counts and byte totals establish a strong locality lead, not an
identity proof for every formal sample. A 0492 capture with serialized range
vectors is required before treating this as the production provider trace.

The media catalog is the sealed
`source-replay-diagnostic-overlap.json` (SHA-256
`b90274360f6859b866d6187adbf307166067c77185220270585ddeebb4c8493e`), which
lists compressed ZIP data ranges. Media overlap below is overlap with those
compressed ranges; it says nothing about decompression or semantic use of a
media part.

## Exact nonempty replay sequence

Intervals use half-open `[offset, offset + length)` notation. The four
zero-length calls are `[412,0]`, `[703,0]`, `[3289,0]`, and `[2891,0]`; an
empty caller buffer returns zero and does not fill or invalidate a cache.

| # | phase | offset | length | interval |
| ---: | --- | ---: | ---: | --- |
| 1 | open | 16,793,014 | 22 | `[16,793,014,16,793,036)` |
| 2 | open | 16,791,709 | 46 | `[16,791,709,16,791,755)` |
| 3 | open | 16,791,709 | 1,305 | `[16,791,709,16,793,014)` |
| 4 | open | 0 | 30 | `[0,30)` |
| 5 | open | 49 | 363 | `[49,412)` |
| 6 | open | 412 | 16 | `[412,428)` |
| 7 | open | 412 | 16 | `[412,428)` |
| 8 | open | 428 | 30 | `[428,458)` |
| 9 | open | 469 | 234 | `[469,703)` |
| 10 | open | 703 | 16 | `[703,719)` |
| 11 | open | 703 | 16 | `[703,719)` |
| 12 | open | 2,907 | 30 | `[2,907,2,937)` |
| 13 | open | 2,965 | 324 | `[2,965,3,289)` |
| 14 | open | 3,289 | 16 | `[3,289,3,305)` |
| 15 | open | 3,289 | 16 | `[3,289,3,305)` |
| 16 | prepare | 1,420 | 30 | `[1,420,1,450)` |
| 17 | prepare | 1,467 | 1,424 | `[1,467,2,891)` |
| 18 | prepare | 2,891 | 16 | `[2,891,2,907)` |
| 19 | prepare | 2,891 | 16 | `[2,891,2,907)` |

There is a compact metadata cluster at offsets 0 through 3,305, with the
largest requested payload 1,424 bytes. The two tail regions around 16.79 MiB
are separate from that cluster. This is enough locality for a small bounded
window, but not a justification for a whole-file read or an unbounded cache.

## One-window simulation

The simulation uses one retained window and a fresh cache for the sequence.
Only the requested bytes are copied to the caller. Every request in this
trace fits wholly within its hit window, so the result is the same whether a
hit requires full containment or permits a prefix when the request crosses a
window end; the crossing behavior still needs a separate correctness test. On
a miss, one fill is issued and replaces the previous window. Fill ends are
clamped to the captured source length of 16,793,036 bytes. The reported fill
bytes are the bytes returned and retained by that bounded fill. The
simulation assumes the inner range adapter can honor a 4 KiB fill, as the
planned 0492 candidate does by placing the read-ahead wrapper under the
existing 64 KiB range cap.

### 4 KiB window beginning at the requested offset

This variant is the literal “start at requested offset” experiment:

| fill | triggering request | physical fill interval | returned bytes | subsequent result |
| ---: | --- | --- | ---: | --- |
| 1 | #1 `[16,793,014,22)` | `[16,793,014,16,793,036)` | 22 | — |
| 2 | #2 `[16,791,709,46)` | `[16,791,709,16,793,036)` | 1,327 | #3 hits |
| 3 | #4 `[0,30)` | `[0,4,096)` | 4,096 | #5–#19 hit |

It produces 3 fills and 16 hits. The physical fill total is 5,445 bytes,
versus 3,966 logical bytes: 1.3729198185x amplification and 1,479 bytes of
overfetch. The range transport would receive three delegated fills rather than
19, subject to the adapter's normal delay and pacing rules. The metadata fill
`[0,4,096)` crosses the first media range
`[4,027,2,101,824)`, so its compressed-media overlap is 69 bytes. The two
tail fills do not overlap a media range.

### 4 KiB window aligned down

The candidate design also considers aligning a miss down to the 4 KiB window
boundary. The first tail request then covers the other two tail requests in
one fill:

| fill | triggering request | physical fill interval | returned bytes | subsequent result |
| ---: | --- | --- | ---: | --- |
| 1 | #1 at 16,793,014 | `[16,789,504,16,793,036)` | 3,532 | #2–#3 hit |
| 2 | #4 at 0 | `[0,4,096)` | 4,096 | #5–#19 hit |

It produces 2 fills and 17 hits. The physical fill total is 7,628 bytes,
1.9233484619x the logical total, with 3,662 bytes of overfetch. It still has
69 bytes of compressed-media overlap from `[0,4,096)`. This is the better
call-count result, but the requested-start policy has less read amplification.
The forward-start policy is the safer first pilot; keep aligned-down behavior
as a comparison until its short-fill and cross-window rules are covered by
tests.

For scale, a 16 KiB aligned window would fill
`[16,777,216,16,793,036)` and `[0,16,384)`: two fills, 32,204 returned
bytes, and 22,013 bytes overlapping compressed media. That larger window is
not justified by this trace, particularly because the aligned tail window
reaches 9,656 bytes into `word/media/image8.png` and the metadata window
reaches 12,357 bytes into `image1.png`.

## Interpretation and required instrumentation

The trace supports a bounded 4 KiB pilot. It does not support carrying over
the baseline zero-media-overlap statement: a read-ahead fill deliberately
fetches bytes the package did not request, and the first metadata fill fetches
69 compressed media bytes. Logical ranges and physical fill ranges must remain
separate in the report.

Before a performance claim, the provider benchmark should emit, per sample
and with a fresh adapter, the following independently scoped evidence:

1. The outer logical range vector, proving the package still received exactly
   the 19 requested intervals and 3,966 returned bytes.
2. The inner physical fill vector below `ReadAheadReadAt`, proving fill count,
   offsets, requested/returned bytes, short fills, and media overlap.
3. Read-ahead counters for hits, misses, fills, fill requested/returned bytes,
   and maximum retained fill size, plus the existing range adapter counters
   for delegated delay/pacing calls.
4. The source version before and after the timed lifecycle and the adapter's
   version check on every hit/fill. The candidate must fail closed on a
   revision change and must not share a cache across samples.

The current `CountingSnapshot.ranges` is retained in memory and used for
media proof, but the sealed v1 formal JSON serializes only aggregate counters.
Adding an opt-in trace schema and running a fresh 0492 before/candidate pilot
will close that evidence gap. The standalone replay should remain labeled as
diagnostic provenance even after that capture exists.
