# ODS formula string literal performance

This report compares the baseline parser at commit `cbc60f1123d105f67f4abdd0d53034fb3325c25f` with the frozen string-literal candidate. The final candidate includes the exact NUL precheck and decoded-length reservation in `formula.rs` (SHA-256 `0b59a9f20b2c9c62fadce3172c646178f8c3cc596c594a48787a0250a651e748`); the regression fixture is `d58abe8eaa7bd9e7ba96139407d9ffcfe5f57350302a5ee795df5297cf7597ac`. The candidate rejects U+0000 before literal storage, reserves decoded content rather than the remaining formula suffix, and preserves doubled-quote and UTF-8 values.

## Method and scope

The standalone [harness](harness/) calls the public `FormulaParser` API. Each lane runs in a fresh release process pinned to CPU 2 with three warmup batches and fifteen measured batches. Repeat counts are recorded in each [group.json](baseline/group.json); ordinary lanes use 1,000 calls per batch, while long inputs use lower bounded repeats. The instrumented allocator records allocation requests and live bytes; `/usr/bin/time -v` records process maximum RSS.

p50 time and allocator calls/bytes are divided by calls per batch. Peak live delta is the raw maximum for a measured batch and is intentionally not divided. RSS is a process-level maximum. The final paired order was baseline comparable, baseline coverage, baseline literals, candidate comparable, candidate coverage, candidate literals. The six raw CSVs, per-lane stdout/status/time files, and digests are retained in [paired/sequence.json](paired/sequence.json) and [paired/raw-sha256.txt](paired/raw-sha256.txt).

The harness checksum contains token count plus original formula text length. Equal checksums therefore establish those fields only; decoded-value preservation is covered by the regression tests and source review. The profile excludes package I/O, XML parsing, evaluation, workbook resolution, URI resolution, publication, and networking.

## Comparable lanes (12)

All 12 baseline and 12 candidate lanes exited status zero. Expected outcomes, successful-call counts, and checksums match; allocator fields remain unchanged on these non-string lanes.

| Case | Input B | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested B/call | Peak live Δ B (batch max) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| parse-bracket-current-cell | 9 | 1000 | success | 172.5 → 175.6 | +1.8% | 3 → 3 | 458 → 458 | 458 → 458 | 2272 → 2212 |
| parse-bracket-sheet-cell | 18 | 1000 | success | 256.9 → 254.3 | -1.0% | 5 → 5 | 697 → 697 | 697 → 697 | 2272 → 2212 |
| parse-bracket-range | 29 | 1000 | success | 407.2 → 398.6 | -2.1% | 7 → 7 | 712 → 712 | 712 → 712 | 2268 → 2276 |
| parse-quoted-doubled-sheet | 19 | 1000 | success | 264.3 → 261.3 | -1.2% | 5 → 5 | 697 → 697 | 697 → 697 | 2272 → 2212 |
| parse-local-refs-256 | 1175 | 128 | success | 21515.7 → 21383.7 | -0.6% | 265 → 265 | 115671 → 115671 | 58775 → 58775 | 2324 → 2276 |
| parse-sum-unbracketed | 15 | 1000 | success | 279.9 → 278.5 | -0.5% | 5 → 5 | 468 → 468 | 468 → 468 | 2272 → 2276 |
| parse-vlookup-unbracketed | 26 | 1000 | success | 446.9 → 449.5 | +0.6% | 8 → 8 | 3172 → 3172 | 1828 → 1828 | 2268 → 2276 |
| parse-malformed-zero-row | 9 | 1000 | refusal | 174.0 → 172.3 | -0.9% | 4 → 4 | 101 → 101 | 76 → 76 | 2208 → 2208 |
| parse-malformed-missing-separator | 8 | 1000 | refusal | 260.6 → 260.3 | -0.1% | 6 → 6 | 138 → 138 | 94 → 94 | 2272 → 2212 |
| parse-malformed-unclosed-bracket | 8 | 1000 | refusal | 158.7 → 160.3 | +1.0% | 4 → 4 | 112 → 112 | 88 → 88 | 2272 → 2212 |
| parse-malformed-missing-row | 8 | 1000 | refusal | 207.4 → 208.4 | +0.5% | 5 → 5 | 139 → 139 | 114 → 114 | 2208 → 2200 |
| parse-malformed-bad-quote | 18 | 1000 | refusal | 161.2 → 161.8 | +0.3% | 4 → 4 | 122 → 122 | 88 → 88 | 2272 → 2276 |

The ordinary parser lanes remain within ±2.1% p50 in this capture. RSS varies with process startup and host state; the displayed values are retained for every lane rather than treated as parser-retained memory.

## Extended reference coverage (16)

All 16 baseline and 16 candidate lanes exited status zero. Valid lanes succeeded and the over-limit IRI lane refused on both sides, with matching successful-call and checksum fields where applicable.

| Case | Input B | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested B/call | Peak live Δ B (batch max) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| coverage-source-cell | 28 | 1000 | success | 307.8 → 302.1 | -1.9% | 5 → 5 | 717 → 717 | 717 → 717 | 2356 → 2212 |
| coverage-source-range | 32 | 1000 | success | 365.5 → 365.9 | +0.1% | 6 → 6 | 722 → 722 | 722 → 722 | 2312 → 2272 |
| coverage-empty-source | 12 | 1000 | success | 217.8 → 210.0 | -3.6% | 4 → 4 | 685 → 685 | 685 → 685 | 2256 → 2276 |
| coverage-unicode-escaped-source | 49 | 1000 | success | 448.8 → 425.5 | -5.2% | 5 → 5 | 758 → 758 | 758 → 758 | 2272 → 2276 |
| coverage-whole-columns | 11 | 1000 | success | 246.7 → 245.4 | -0.5% | 5 → 5 | 685 → 685 | 685 → 685 | 2256 → 2276 |
| coverage-whole-rows | 11 | 1000 | success | 176.8 → 179.4 | +1.5% | 3 → 3 | 683 → 683 | 683 → 683 | 2252 → 2204 |
| coverage-cross-sheet-range | 25 | 1000 | success | 352.2 → 357.6 | +1.5% | 6 → 6 | 487 → 487 | 487 → 487 | 2272 → 2276 |
| coverage-nested-inherited | 21 | 1000 | success | 393.5 → 400.9 | +1.9% | 8 → 8 | 741 → 741 | 741 → 741 | 2272 → 2276 |
| coverage-ref-error | 11 | 1000 | success | 129.6 → 125.9 | -2.8% | 3 → 3 | 683 → 683 | 683 → 683 | 2272 → 2276 |
| coverage-colon-sheet-1k | 1033 | 1000 | success | 1240.8 → 1240.0 | -0.1% | 4 → 4 | 2506 → 2506 | 2506 → 2506 | 2272 → 2276 |
| coverage-colon-sheet-4k | 4105 | 1000 | success | 4247.3 → 4238.5 | -0.2% | 4 → 4 | 8650 → 8650 | 8650 → 8650 | 2272 → 2192 |
| coverage-colon-sheet-16k | 16393 | 128 | success | 16034.4 → 16047.5 | +0.1% | 4 → 4 | 33226 → 33226 | 33226 → 33226 | 2272 → 2464 |
| coverage-source-iri-1k | 1036 | 1000 | success | 4899.7 → 4767.6 | -2.7% | 5 → 5 | 2733 → 2733 | 2733 → 2733 | 2208 → 2276 |
| coverage-source-iri-4k | 4108 | 1000 | success | 18778.9 → 17856.1 | -4.9% | 5 → 5 | 8877 → 8877 | 8877 → 8877 | 2228 → 2276 |
| coverage-source-iri-16k | 16396 | 128 | success | 74018.0 → 69925.9 | -5.5% | 5 → 5 | 33453 → 33453 | 33453 → 33453 | 2272 → 2208 |
| coverage-source-iri-over-16k | 16397 | 128 | refusal | 31953.0 → 35769.1 | +11.9% | 6 → 6 | 16657 → 16657 | 16437 → 16437 | 2272 → 2272 |

The long colon-sheet lanes are within 0.3% in this final capture. The over-limit IRI refusal is +11.9%; this is an observed rich-reference result and the harness does not isolate whether cost comes from reference parsing, annotation handling, or code generation. The largest positive RSS change in the reference group is 2,272 → 2,464 KiB (+8.5%) for the 16 KiB colon-sheet lane, despite identical allocator counters. This crosses the 5% review threshold; process startup and host state can affect RSS, and this capture does not establish its cause.

## Literal lanes (9)

All 9 baseline and 9 candidate literal lanes exited status zero. Successful lanes retain matching harness checksums; the unterminated lane refuses on both sides.

| Case | Input B | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested B/call | Peak live Δ B (batch max) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| literal-many-short-64 | 259 | 128 | success | 4308.2 → 4615.1 | +7.1% | 71 → 71 | 36675 → 28547 | 22787 → 14659 | 2272 → 2276 |
| literal-many-short-256 | 1027 | 64 | success | 16832.6 → 18398.7 | +9.3% | 265 → 265 | 246339 → 115523 | 189443 → 58627 | 2524 → 2212 |
| literal-many-short-1024 | 4099 | 16 | success | 77562.2 → 100589.9 | +29.7% | 1035 → 1035 | 2559555 → 463427 | 2330627 → 234499 | 4512 → 2532 |
| literal-many-short-4096 | 16387 | 1 | success | 340741 → 310061 | -9.0% | 4109 → 4109 | 35405379 → 1855043 | 34488323 → 937987 | 28768 → 3656 |
| literal-many-empty-1024 | 3075 | 8 | success | 78364.1 → 48145.2 | -38.6% | 1035 → 11 | 2033731 → 461379 | 1804803 → 232451 | 4064 → 2492 |
| literal-mixed-utf8-doubled-quotes | 7171 | 4 | success | 25372.8 → 24030.2 | -5.3% | 265 → 265 | 1041987 → 127299 | 985091 → 70403 | 3296 → 2212 |
| literal-single-plain-64k | 65542 | 8 | success | 17930.1 → 20333.8 | +13.4% | 3 → 3 | 131527 → 131526 | 131527 → 131526 | 2528 → 2508 |
| literal-single-doubled-quote-64k | 65542 | 8 | success | 24492.6 → 34255.2 | +39.9% | 3 → 3 | 131527 → 98758 | 131527 → 98758 | 2592 → 2468 |
| literal-unterminated-64k | 65541 | 8 | refusal | 16410 → 16278.8 | -0.8% | 5 → 4 | 131163 → 65627 | 131104 → 65568 | 2484 → 2512 |

The memory result is the intended tradeoff. At 64/256/1,024/4,096 short literals, requested bytes per call fall by approximately 22%/53%/82%/95%, and peak live deltas fall by approximately 36%/69%/90%/97%. Empty literals reduce allocation calls from 1,035 to 11 per parser call and peak live bytes by about 87%. Mixed UTF-8/doubled-quote input reduces requested bytes by about 88% and peak live bytes by about 93%. A plain 64 KiB string keeps essentially the same requested and peak bytes; doubled quotes reduce both by about 25%. The 4,096-short lane is 9.0% faster and empty literals are 38.6% faster in the paired p50 capture.

The remaining measured latency costs are visible: 64/256/1,024 short-literal lanes are +7.1%/+9.3%/+29.7%, plain 64 KiB is +13.4%, and doubled-quote 64 KiB is +39.9%. Unterminated 64 KiB is within 1% after the bulk NUL-check optimization. These costs are retained as scope tradeoffs; there is no broad speedup claim.

## Targeted A/B/A/B follow-up

The five material positive literal deltas were repeated in fresh processes in baseline/candidate/baseline/candidate order with the same 3/15 settings. The table reports the median of each side's two p50-per-call rows.

| Case | Baseline p50 ns/call (ABAB median) | Candidate p50 ns/call (ABAB median) | Δ |
| --- | ---: | ---: | ---: |
| literal-many-short-64 | 4303.8 | 4646.1 | +8.0% |
| literal-many-short-256 | 16835 | 18417.0 | +9.4% |
| literal-many-short-1024 | 77381.9 | 92141.3 | +19.1% |
| literal-single-plain-64k | 17925.7 | 20301.3 | +13.3% |
| literal-single-doubled-quote-64k | 24282 | 36700.8 | +51.1% |

The follow-up confirms the +19% 1,024-short, +13% plain 64 KiB, and +51% doubled-quote 64 KiB costs at this smaller sample; the 64/256 short-literal lanes remain roughly +8%/+9%. The main pair and follow-up differ in magnitude, but both identify the same direction. Raw rounds and binary digests are in [targeted-abab/sequence.json](targeted-abab/sequence.json).

## Hardware counters

`perf stat` captured cycles, instructions, branches, branch misses, and cache misses for the 4,096-short and doubled-quote 64 KiB lanes on CPU 2. Values below are totals divided by `(warmups + iterations) × repeat`, and include process setup, warmups, and output, so they are diagnostic.

| Lane | Side | Cycles/call | Instructions/call | Branches/call | Branch misses/call | Cache misses/call | IPC |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| literal-many-short-4096 | baseline | 2877797.4 | 5561231.7 | 1086038.7 | 2478.7 | 19690 | 1.932 |
| literal-many-short-4096 | candidate | 1699880.4 | 3756988.1 | 717736.2 | 1311.3 | 5786.3 | 2.210 |
| literal-single-doubled-quote-64k | baseline | 134669.2 | 910478.8 | 210116.5 | 131.3 | 349.3 | 6.761 |
| literal-single-doubled-quote-64k | candidate | 186559.1 | 1176750.1 | 316641.1 | 129.5 | 323.4 | 6.308 |

The 4,096-short candidate lowers instruction and cycle totals substantially while the doubled-quote 64 KiB candidate increases them, matching the latency direction. Raw perf stdout/stderr and command receipts are in [hardware/perf-stat/summary.json](hardware/perf-stat/summary.json).

## Correctness, resource, and provenance gates

The frozen candidate passed 828 tests across 50 targets, Clippy with warnings denied, warning-denied rustdoc, doctests, and formatting; the root receipt is [gates/results.json](../gates/results.json). The candidate worktree checks passed 49 formula unit tests, 6 string regression tests, 6 scan regressions, 4 tokenizer regressions, 7 function-catalog tests, and 10 reference integration tests (82 total), recorded in [candidate/isolated-checks.json](candidate/isolated-checks.json). The baseline string fixture run is intentionally status 101: its source-bound receipt records the baseline formula hash and final fixture hash, with exactly the expected 4 passes and 2 failures, in [baseline/isolated-string-regression.json](baseline/isolated-string-regression.json).

Final source and harness provenance are [candidate/source-sha256.json](candidate/source-sha256.json), [baseline/source-sha256.json](baseline/source-sha256.json), and [candidate/harness-sha256.txt](candidate/harness-sha256.txt). The saved binaries were independently verified by root: baseline SHA-256 `52fc5e50b1a8ddebc0dc04fcdde8ef9062728b0bf33e45d39b3a5e2182b550bb` (1,173,352 bytes) and final candidate SHA-256 `d9dcf284a1b311191bbd1b959ec9c0c6915dec5a585f4c5e588ef29e570839d5` (1,140,352 bytes). The first candidate already rejected NUL, but checked each byte in the quote loop. Moving that check to one slice operation reduced the measured 64 KiB plain-string regression from 78% to 13% and restored the unterminated case from an 85% slowdown to within 1% of baseline. First-candidate measurements before the bulk NUL-check optimization are retained under [pre-nul-sequence.json](pre-nul-sequence.json) and the `*-pre-nul/` lane directories; they are superseded by the final pair.

## Limitations

This is one pinned virtual CPU, one host, fresh processes, and fifteen measured batches per lane. The A/B/A/B follow-up narrows process variance for selected literal lanes but does not establish a statistical tail bound. RSS reflects launcher and host state. The harness covers bounded literal corpus sizes and selected parser controls; it does not establish general parser throughput, whole-formula complexity, or package-level performance.
