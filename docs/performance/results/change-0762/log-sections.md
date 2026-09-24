# Ready-to-paste log sections for change 0762

## HOTSPOTS.md

## 0762 — per-write Deflate calls, the pre-finish sync flush and level-6 match search in the streaming writers

[0762](0762-streaming-batch-compression.md) applies owner decision 4 of 0758
to the owned ZIP entries every streaming DOCX, XLSX and PPTX writer uses.
Input is staged and handed to zlib-rs in 16 KiB chunks cut at absolute member
offsets, which also makes the stream a pure function of the member bytes. The
pre-finish sync flush is gone (6.8 bytes and a block per member). Owned
members compress at level 5, whose shorter match search halves the XLSX
sheet's compression at +0.21–0.23% size on real fixture members (≤ 0.01% on
the harness corpora). Large creation falls 27.8% (DOCX), 44.8% (XLSX) and
9.9% (PPTX). Remaining: the DOCX and XLSX per-charge budget atomics (record
0763), and PPTX's per-member dynamic Huffman construction and 128 KiB
hash-table reset, which no level but level 1 (+25% size) avoids. Disproven:
larger chunks buy nothing once calls carry 4 KiB, and levels 3 and 4 grow real
content by more than 1%.

## REPORT.md

## 0762 — streaming batch compression

[0762](0762-streaming-batch-compression.md) is retained with
`performance_claim: none`. Base `1d1044e3ac`; commit `bc18e8abdd`. With both
legs built by the same command, `docx_streaming_create` moves 49.353 →
35.661 ms (−27.77%, CI [−28.47%, −27.22%]) large and 3.156 → 2.278 ms
(−27.84%) medium; `xlsx_streaming_create` 163.942 → 90.550 ms (−44.81%) and
10.631 → 5.639 ms (−47.01%); `pptx_streaming_create` 190.417 → 171.484 ms
(−9.89%) and 6.378 → 5.835 ms (−8.57%). Instructions fall 23%, 18–19% and
5–7%. The streaming packages change bytes; all 17,050 members decompress to
the same bytes; eight of nine corpora shrink (to −2.17%) and one grows 4 bytes
(+0.12%). Each archive makes one more 16 KiB allocation. Controls (preservation
paths, identical output): semantic DOCX edit +1.19% (CI [−0.08%, +1.53%])
with −0.02% instructions, XLSX ordinary save −3.50% (host-load noise) with
+0.01% instructions.

## GOAL_AUDIT.md

## 0762 — deterministic batch compression for creation writers only

[0762](0762-streaming-batch-compression.md) changes output bytes only where
owner decision 4 of 0758 allows: owned ZIP Deflate entries, reached in
production only by the three streaming writers. The preservation writer and
every other borrowed Deflate path keep level 6, one codec call per write and
their bytes (62 DOCX fixture publications and the XLSX ordinary-save digest
compared across the legs). Owned output is a pure function of the member bytes
and explicit flush offsets: the codec is called only with the rest of a
complete 16 KiB chunk and an empty output buffer. Evidence: a staged model
checked against flate2's own encoder, 24 random write splittings, short and
interrupting sinks, five DOCX text splittings in five processes, and
mutation checks. No limit is ever exceeded on output; compressed-size, output
and sink refusals can surface up to one 16 KiB chunk later than before (on
top of the codec's own buffering), and the uncompressed limits stay exact per
write. After review, a poisoned owned entry refuses `write`, `flush` and
`finish` before its sink sees another byte, and `finish` retries an
interrupted sink call instead of losing the archive; both were pre-existing.
No durable format binds these bytes.
