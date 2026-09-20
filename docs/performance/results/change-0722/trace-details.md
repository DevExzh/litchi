# Untimed structural reader matrix

Both lanes preserve the public report, ordered MCE vectors, structured chunks, and ranges.
Counts include successful EOF reads. A parser error returned by the read itself is not a successful read.
Baseline range reads are counted before admission, with an unknown end offset (`null`) because the event borrows the reader.
Baseline range observers run after admission and classification; candidate observers run immediately before range classification, only while that state remains admitted.
Refusal observer counts therefore have different boundaries and are not interchangeable reader counts.
These are parser events, not physical I/O, timing, instruction, or allocation measurements.

| Case | XML bytes | Baseline alt / range reads | Candidate fused reads | Baseline / candidate observers | Ordered MCE input lengths |
| --- | ---: | ---: | ---: | ---: | --- |
| active-offset-count-overflow | 205 | 7 / 7 | 7 | 7 / 7 | 0, 2 |
| bom-plain | 197 | 6 / 6 | 6 | 6 / 6 | 0, 1 |
| empty-at-depth-256 | 5536 | 257 / 0 | 257 | 0 / 129 | none |
| foreign-lookalikes | 265 | 9 / 9 | 9 | 9 / 9 | 0, 1 |
| generated-medium | 21517 | 1410 / 1410 | 1410 | 1410 / 1410 | 0, 200 |
| malformed-alternate-content-no-choice-no-alt | 351 | 10 / 10 | 10 | 10 / 10 | 0, 3 |
| malformed-tail | 196 | 3 / 0 | 3 | 0 / 3 | none |
| marker-text | 262 | 12 / 12 | 12 | 12 / 12 | 0, 1 |
| nested-alt-in-paragraph | 269 | 9 / 9 | 9 | 9 / 9 | 2, 2 |
| nested-alt-in-table | 301 | 14 / 14 | 14 | 14 / 14 | 2, 2 |
| numbered-list | 4563 | 178 / 178 | 178 | 178 / 178 | 0, 5 |
| range-depth-before-bad-alt | 2871 | 257 / 0 | 257 | 0 / 129 | none |
| range-depth-only | 2864 | 260 / 129 | 260 | 128 / 129 | 0 |
| strict-no-alt-plain-blocks | 187 | 8 / 8 | 8 | 8 / 8 | 0, 3 |
| strict-valid-fallback-with-inactive-anchor | 454 | 15 / 15 | 15 | 15 / 15 | 2, 4 |
| transitional-no-alt-plain-blocks | 211 | 8 / 8 | 8 | 8 / 8 | 0, 3 |
| unbound-fragment | 49 | 7 / 7 | 7 | 7 / 7 | 0, 2 |
| unknown-must-understand-no-alt | 324 | 7 / 7 | 7 | 7 / 7 | 0, 2 |
| valid-choice-with-inactive-fallback | 486 | 16 / 16 | 16 | 16 / 16 | 2, 5 |
| valid-fallback-no-alt-blocks | 428 | 16 / 16 | 16 | 16 / 16 | 0, 5 |
| valid-fallback-with-inactive-anchor | 486 | 16 / 16 | 16 | 16 / 16 | 2, 5 |

Input hashes:

- `trace/baseline/stderr`: `3b47fdb6718d11769227bbabba5213fdb382e1b6b30802173a7176e5744d7614`
- `trace/candidate/stderr`: `37d5b18d0e53029dc325f9a4f1dcd578a40c9509c55383de11486de020dd48f1`
- `trace-analyze.py`: `63d74f745a2e8902347f26627a5cde877101d6253d2ca6fc38698b3e82250c74`
- `trace-details.py`: `ed11586e8e77e0d38da1715742e7c53f94b9ec765cbefb8dc4048bd1776816ce`
