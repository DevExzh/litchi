# Ready-to-paste log sections for change 0752

## HOTSPOTS.md

## 0752 — per-tiny-write CRC-32 and budget handles in the streaming writers

[0752](0752-streaming-writer-small-write-batching.md) removes the streaming
writers' per-write costs that exact semantics allow. Owned ZIP entries stage
short writes for a one-pass CRC-32: the harness lock's crc32fast 1.5.0 took a
table lookup per byte for every DOCX write, since all are shorter than its
128-byte SIMD threshold. Budget charging walks the parent chain by reference,
a `Reservation` holds one handle, and the new borrowed `ScopedReservation`
holds none. DOCX escaping copies plain runs whole with the old 64-byte chunk
boundaries. Large DOCX creation falls 46.7% (91.5 → 48.6 ms). XLSX and PPTX
creation improve 1–3%: Deflate dominates them. Disproven: coalescing
compressor input is not byte-transparent for zlib-rs 0.6.7 (seven re-splits of
the DOCX payload give seven different valid streams), so it stays out. It
would save about 10 ms of the DOCX iteration at the cost of new output bytes.
The remaining DOCX hotspot is its seven one-atomic `consume` charges per
paragraph (26% of samples), which only a semantic decision can reduce.

## REPORT.md

## 0752 — streaming writer small-write batching

[0752](0752-streaming-writer-small-write-batching.md) is retained with
`performance_claim: none`. Base `6d989cad63`; commits `9ce59dc565`,
`709efbafab`, `1167398a3e`. With both legs built by the same command,
`docx_streaming_create` moves 91.474 → 48.596 ms (−46.66%, CI [−47.25%,
−45.24%]) large and 5.754 → 3.094 ms (−46.30%) medium, on 27.5% fewer
instructions. `xlsx_streaming_create` moves −1.94% / −3.34% and
`pptx_streaming_create` −1.45% (CI spans zero) / −1.86%. Each streaming
archive makes one more 4 KiB allocation. Controls: semantic DOCX edit +0.10%,
XLSX ordinary save −1.15%, CFB shared bulk read +2.52% / +3.06% with
identical instruction counts. An A/A pair of the unchanged base shows +1.47%
[+0.10%, +3.16%] on the same corpus, and its cost moves inside unchanged CFB
functions, so the shift is attributed to code placement. Every published
digest is identical in the 112 processes that report one. A base-versus-candidate
probe gives identical transcripts over 28,266 limit, cancellation and sink
scenarios.

## GOAL_AUDIT.md

## 0752 — exact semantics for streaming-writer write batching

[0752](0752-streaming-writer-small-write-batching.md) keeps every budget
limit, refusal value, cancellation point, reported progress and output byte of
the DOCX, XLSX and PPTX streaming writers. Compressor input is not coalesced,
because zlib-rs's stream depends on its write boundaries, as measured.
DOCX escaping preserves the 64-byte chunk boundaries for the same reason. The
CRC stage is invisible outside the data descriptor. The borrowed reservation
charges and releases through the same functions as the owned one. Evidence:
unit tests at N−1/N/N+1 on a three-level hierarchy, the old escaping and
scanning loops kept as test oracles, the owned-versus-borrowed complete-ZIP
differential, and a probe built against both trees whose 28,266-scenario
transcripts are byte-identical. Decision 8's error-timing allowance was not
used. The additions (`reserve_scoped`, `ScopedReservation`) are
non-breaking.
