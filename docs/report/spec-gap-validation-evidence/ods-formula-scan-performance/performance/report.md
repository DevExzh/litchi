# ODS formula scan performance

This report compares the committed parser at `5fae21a34e163a3300fc0ac4f9c302f5731e3cbb` with the frozen tokenizer scan candidate. The candidate source passed the root gate receipt before measurement. Its implementation patch contains only `crates/litchi-ods/src/codec/formula.rs` and the new scan regression test; the reference parser is unchanged from the base commit.

The candidate caches the end of the last legacy sheet-name run and reuses it for compact classification and legacy cell parsing. A bounded compact scan handles complete function and cell candidates, while the established compatibility path remains available for space-bearing sheet names, absolute references, malformed suffixes, and other partial-token behavior. The new regression tests cover cache reuse, token boundaries, malformed inputs, and absolute forms; the existing typed allocation-failure tests remain part of the suite.

## Method and scope

The standalone [harness](harness/) calls `litchi_ods::codec::formula::FormulaParser` through its public parser API. Every lane uses a fresh release process pinned with `taskset -c 2`, three warmup batches, and fifteen measured batches. Ordinary lanes execute 1,000 parser calls per batch; 256-reference and 16 KiB lanes use the bounded repeat counts recorded in each table. The instrumented `System` allocator records requests and live-byte peaks; `/usr/bin/time -v` records process maximum RSS. p50 time, allocator calls, and requested bytes are normalized by calls per batch. Peak live delta is the raw maximum for one measured batch and is intentionally not normalized.

The paired order is baseline comparable, candidate comparable, baseline coverage, candidate coverage, baseline scaling, candidate scaling. Both sides use the same final harness and saved release binaries. Outcomes, checksums, statuses, and allocator fields are retained in the raw CSVs. The run measures token construction, reference parsing, compatibility handoff, error handling, and allocation/destruction. It excludes package I/O, XML parsing, evaluation, workbook resolution, publication, URI resolution, and networking.

## Comparable lanes (12)

All 12 baseline and 12 candidate lanes exited with status zero. Every expected outcome and checksum matches. The table reports p50 nanoseconds per parser call; arrows are baseline → candidate.

| Case | Input bytes | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested bytes/call | Peak live delta (batch max) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| parse-bracket-current-cell | 9 | 1000 | success | 173.02 → 180.79 | +4.5% | 3.00 → 3.00 | 458.00 → 458.00 | 458 → 458 | 2272 → 2272 |
| parse-bracket-sheet-cell | 18 | 1000 | success | 254.41 → 261.01 | +2.6% | 5.00 → 5.00 | 697.00 → 697.00 | 697 → 697 | 2272 → 2272 |
| parse-bracket-range | 29 | 1000 | success | 404.04 → 405.68 | +0.4% | 7.00 → 7.00 | 712.00 → 712.00 | 712 → 712 | 2272 → 2248 |
| parse-quoted-doubled-sheet | 19 | 1000 | success | 271.20 → 262.06 | -3.4% | 5.00 → 5.00 | 697.00 → 697.00 | 697 → 697 | 2228 → 2272 |
| parse-local-refs-256 | 1175 | 128 | success | 21239.24 → 22518.94 | +6.0% | 265.00 → 265.00 | 115671.00 → 115671.00 | 58775 → 58775 | 2208 → 2304 |
| parse-sum-unbracketed | 15 | 1000 | success | 278.96 → 276.65 | -0.8% | 5.00 → 5.00 | 468.00 → 468.00 | 468 → 468 | 2272 → 2272 |
| parse-vlookup-unbracketed | 26 | 1000 | success | 447.98 → 451.03 | +0.7% | 8.00 → 8.00 | 3172.00 → 3172.00 | 1828 → 1828 | 2272 → 2272 |
| parse-malformed-zero-row | 9 | 1000 | refusal | 175.28 → 174.44 | -0.5% | 4.00 → 4.00 | 101.00 → 101.00 | 76 → 76 | 2272 → 2272 |
| parse-malformed-missing-separator | 8 | 1000 | refusal | 262.18 → 259.44 | -1.0% | 6.00 → 6.00 | 138.00 → 138.00 | 94 → 94 | 2272 → 2232 |
| parse-malformed-unclosed-bracket | 8 | 1000 | refusal | 160.18 → 159.80 | -0.2% | 4.00 → 4.00 | 112.00 → 112.00 | 88 → 88 | 2220 → 2268 |
| parse-malformed-missing-row | 8 | 1000 | refusal | 211.52 → 210.61 | -0.4% | 5.00 → 5.00 | 139.00 → 139.00 | 114 → 114 | 2272 → 2272 |
| parse-malformed-bad-quote | 18 | 1000 | refusal | 164.11 → 161.78 | -1.4% | 4.00 → 4.00 | 122.00 → 122.00 | 88 → 88 | 2208 → 2272 |

The ordinary function lanes are close: SUM is -0.8% and VLOOKUP is +0.7%. The paired 256-reference lane is +6.0% in the main capture. Its targeted A/B/A/B follow-up estimates +1.9% p50 (mean +2.0%) and therefore does not reproduce the full +6.0% result. The follow-up raw rounds are retained in [local-refs-abab](local-refs-abab/). Allocation calls, requested bytes, released bytes, checksums, successes, and peak live deltas match in all comparable pairs. RSS varies with process startup and host state; its largest comparable change is 2,208 → 2,304 KiB (+4.3%).

## Extended reference coverage (16)

All 16 baseline and 16 candidate lanes exited with status zero. Valid lanes succeeded and the over-limit IRI lane refused on both sides; expected outcomes and checksums match pairwise.

| Case | Input bytes | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested bytes/call | Peak live delta (batch max) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| coverage-source-cell | 28 | 1000 | success | 320.16 → 312.80 | -2.3% | 5.00 → 5.00 | 717.00 → 717.00 | 717 → 717 | 2324 → 2252 |
| coverage-source-range | 32 | 1000 | success | 372.36 → 360.10 | -3.3% | 6.00 → 6.00 | 722.00 → 722.00 | 722 → 722 | 2272 → 2204 |
| coverage-empty-source | 12 | 1000 | success | 210.27 → 208.47 | -0.9% | 4.00 → 4.00 | 685.00 → 685.00 | 685 → 685 | 2208 → 2272 |
| coverage-unicode-escaped-source | 49 | 1000 | success | 442.31 → 444.74 | +0.5% | 5.00 → 5.00 | 758.00 → 758.00 | 758 → 758 | 2324 → 2208 |
| coverage-whole-columns | 11 | 1000 | success | 247.47 → 248.75 | +0.5% | 5.00 → 5.00 | 685.00 → 685.00 | 685 → 685 | 2272 → 2336 |
| coverage-whole-rows | 11 | 1000 | success | 175.63 → 175.94 | +0.2% | 3.00 → 3.00 | 683.00 → 683.00 | 683 → 683 | 2272 → 2208 |
| coverage-cross-sheet-range | 25 | 1000 | success | 351.44 → 360.49 | +2.6% | 6.00 → 6.00 | 487.00 → 487.00 | 487 → 487 | 2268 → 2208 |
| coverage-nested-inherited | 21 | 1000 | success | 396.99 → 390.55 | -1.6% | 8.00 → 8.00 | 741.00 → 741.00 | 741 → 741 | 2272 → 2268 |
| coverage-ref-error | 11 | 1000 | success | 127.06 → 130.60 | +2.8% | 3.00 → 3.00 | 683.00 → 683.00 | 683 → 683 | 2208 → 2272 |
| coverage-colon-sheet-1k | 1033 | 1000 | success | 1430.38 → 1430.18 | -0.0% | 4.00 → 4.00 | 2506.00 → 2506.00 | 2506 → 2506 | 2192 → 2272 |
| coverage-colon-sheet-4k | 4105 | 1000 | success | 5039.44 → 5023.84 | -0.3% | 4.00 → 4.00 | 8650.00 → 8650.00 | 8650 → 8650 | 2268 → 2272 |
| coverage-colon-sheet-16k | 16393 | 128 | success | 19235.72 → 19230.24 | -0.0% | 4.00 → 4.00 | 33226.00 → 33226.00 | 33226 → 33226 | 2336 → 2272 |
| coverage-source-iri-1k | 1036 | 1000 | success | 4862.98 → 4725.80 | -2.8% | 5.00 → 5.00 | 2733.00 → 2733.00 | 2733 → 2733 | 2272 → 2272 |
| coverage-source-iri-4k | 4108 | 1000 | success | 18499.76 → 17955.81 | -2.9% | 5.00 → 5.00 | 8877.00 → 8877.00 | 8877 → 8877 | 2272 → 2228 |
| coverage-source-iri-16k | 16396 | 128 | success | 72866.28 → 70337.84 | -3.5% | 5.00 → 5.00 | 33453.00 → 33453.00 | 33453 → 33453 | 2272 → 2336 |
| coverage-source-iri-over-16k | 16397 | 128 | refusal | 37514.32 → 33575.32 | -10.5% | 6.00 → 6.00 | 16657.00 → 16657.00 | 16437 → 16437 | 2208 → 2272 |

The ordinary source-qualified and axis/reference forms range from -3.5% to +2.8%; the long colon-bearing sheet-name lanes are within 0.3% in this scan. The over-limit IRI refusal is -10.5%. These are observed results; this harness does not isolate cost among reference parsing, code generation, and refusal handling. Allocation calls, requested bytes, released bytes, checksums, successes, and peak live deltas are identical across all coverage pairs. The largest RSS difference is 2,324 → 2,208 KiB (-5.0%), within process-level variation.

## Tokenizer scaling (8)

The added cases use space-separated and contiguous `A1` token families at 256, 1,024, 4,096, and 16,384 tokens. All 8 baseline and 8 candidate lanes succeeded with matching checksums and allocator fields.

| Case | Input bytes | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested bytes/call | Peak live delta (batch max) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| scaling-space-a1-256 | 771 | 128 | success | 48513.12 → 17267.97 | -64.4% | 264.00 → 264.00 | 57923.00 → 57923.00 | 29699 → 29699 | 2208 → 2272 |
| scaling-space-a1-1024 | 3075 | 32 | success | 559237.94 → 76347.56 | -86.3% | 1034.00 → 1034.00 | 233027.00 → 233027.00 | 118787 → 118787 | 2528 → 2464 |
| scaling-space-a1-4096 | 12291 | 8 | success | 7916235.00 → 275847.50 | -96.5% | 4108.00 → 4108.00 | 933443.00 → 933443.00 | 475139 → 475139 | 2740 → 2784 |
| scaling-space-a1-16384 | 49155 | 2 | success | 150946370.00 → 1152390.50 | -99.2% | 16398.00 → 16398.00 | 3735107.00 → 3735107.00 | 1900547 → 1900547 | 5320 → 5300 |
| scaling-contiguous-a1-256 | 516 | 16 | success | 46540.81 → 20172.56 | -56.7% | 264.00 → 264.00 | 57668.00 → 57668.00 | 29444 → 29444 | 2252 → 2272 |
| scaling-contiguous-a1-1024 | 2052 | 8 | success | 500491.12 → 89284.12 | -82.2% | 1034.00 → 1034.00 | 232004.00 → 232004.00 | 117764 → 117764 | 2580 → 2528 |
| scaling-contiguous-a1-4096 | 8196 | 2 | success | 6823512.00 → 327106.50 | -95.2% | 4108.00 → 4108.00 | 929348.00 → 929348.00 | 471044 → 471044 | 2784 → 2784 |
| scaling-contiguous-a1-16384 | 32772 | 1 | success | 105574770.00 → 1377997.00 | -98.7% | 16398.00 → 16398.00 | 3718724.00 → 3718724.00 | 1884164 → 1884164 | 5348 → 5272 |

The candidate removes the measured repeated lookahead cost: space-separated p50 improves by 64.4%, 86.3%, 96.5%, and 99.2% at 256, 1,024, 4,096, and 16,384 tokens; contiguous p50 improves by 56.7%, 82.2%, 95.2%, and 98.7%. The final scaling batch elapsed time is 9.044 seconds for baseline versus 0.266 seconds for candidate. These cases exercise the intentionally bounded scanner and do not establish complexity or throughput for other formula grammar families.

## Targeted 256-reference follow-up

The main paired lane had repeat 128, three warmups, and fifteen measured batches. To assess its +6.0% result, the saved binaries were run in fresh processes in baseline/candidate/baseline/candidate order with the same lane and harness settings.

| Metric | Baseline median of two rounds | Candidate median of two rounds | Δ |
| --- | ---: | ---: | ---: |
| p50 ns/call | 20997.32 | 21406.66 | +1.95% |
| mean ns/call | 21005.33 | 21424.96 | +2.00% |
| p95 ns/call | 21100.72 | 21556.98 | +2.16% |
| Max RSS KiB | 2262 | 2298 | +1.59% |

The ABAB estimate is the appropriate bounded follow-up for this lane: it still shows a small roughly 2% candidate cost, while the original +6.0% single paired result is not stable at that magnitude. All allocation and semantic fields match in every round. Raw CSVs, commands, binary digests, and sequence metadata are in [local-refs-abab](local-refs-abab/).

## Annotation-removal check

A prior isolated experiment removed exactly three reference-parser code-generation annotations (`inline` on `Parser::bump_component`, `cold` on `limit_error`, and `inline` on `copy_component`) and ran the four flagged long lanes in A/B/A/B order. Colon-sheet medians changed by +0.37%, +0.30%, and -0.03%; the over-limit IRI refusal changed by -3.02% within the observed round variance. Removing those annotations did not reverse the long-input behavior, so the reference annotations remain unchanged. The complete raw rounds and receipt are in [annotation-removals/sequence.json](annotation-removals/sequence.json).

## Hardware counters

`perf stat` captured cycles, instructions, branches, branch misses, and cache misses on CPU 2 for SUM, 256 local references, and both 16,384-token scaling forms. Counters include process setup, warmups, and output, so they are diagnostic; totals are divided by `(warmups + iterations) × repeat`.

| Lane | Side | Cycles/call | Instructions/call | Branches/call | Branch misses/call | Cache misses/call | IPC |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| sum | baseline | 1445.7 | 3137.8 | 605.8 | 1.2 | 3.2 | 2.171 |
| local-refs-256 | baseline | 96870.5 | 279852.7 | 55040.7 | 55.3 | 23.0 | 2.889 |
| scaling-space-16384 | baseline | 553349203.8 | 4445486608.9 | 1211238879.6 | 28415.3 | 20310.7 | 8.034 |
| scaling-contiguous-16384 | baseline | 480483911.4 | 3913776495.4 | 1212314594.3 | 25501.4 | 20655.7 | 8.145 |
| sum | candidate | 1439.9 | 3135.4 | 602.5 | 1.1 | 2.8 | 2.178 |
| local-refs-256 | candidate | 97690.9 | 281861.4 | 55284.4 | 42.6 | 22.4 | 2.885 |
| scaling-space-16384 | candidate | 5346920.9 | 16333240.5 | 3387062.5 | 2843.8 | 11637.7 | 3.055 |
| scaling-contiguous-16384 | candidate | 6421881.2 | 20932277.3 | 4266537.1 | 5573.0 | 13658.6 | 3.260 |

The counter follow-up agrees with the scaling result. SUM counters are nearly unchanged; the 256-reference lane has a small candidate increase in cycles and instructions, consistent with the ABAB latency result. Both 16,384-token scaling lanes show orders-of-magnitude lower candidate instruction and cycle totals. Raw perf output and commands are retained in [paired/perf-stat](paired/perf-stat/), with the sequence in [paired/perf-stat/sequence.json](paired/perf-stat/sequence.json).

## Correctness, resource, and provenance gates

The root receipt records 822 tests across 49 targets, strict Clippy, warning-denied rustdoc, doctests, formatting, and boundary checks. Candidate isolated checks passed 49 formula unit tests, six scan regression tests, four tokenizer regression tests, seven function-catalog tests, and ten reference integration tests. The baseline compatibility check passed six scan regression tests before candidate source was applied. These logs and statuses are linked from [candidate/isolated-checks.json](candidate/isolated-checks.json) and [baseline/isolated-scan-regression.log](baseline/isolated-scan-regression.log).

Across all 36 paired lanes, expected outcomes, checksums, allocation calls, requested/released bytes, successful-call counts, and peak live deltas are identical. The candidate therefore preserves the measured allocation and semantic fields while changing the scan path. RSS remains a process-level maximum and varies with launcher and host state.

The final candidate source manifest covers all nine files in the root receipt at [candidate/source-sha256.json](candidate/source-sha256.json); the baseline manifest is [baseline/source-sha256.json](baseline/source-sha256.json). The baseline and candidate binaries are identified by [baseline/binary-sha256.txt](baseline/binary-sha256.txt) and [candidate/binary-sha256.txt](candidate/binary-sha256.txt). Final paired raw checksums and order are in [paired/sequence.json](paired/sequence.json) and [paired/raw-sha256.txt](paired/raw-sha256.txt). The replayable zero-context patch and its digest are [candidate.patch](../candidate.patch) and [candidate-patch-sha256.txt](../candidate-patch-sha256.txt); root independently checked it against a clean base archive. The isolated worktree, Cargo target, baseline snapshot, executable copies, and profiling helper files were removed after hash verification; sizes and post-delete checks are in [cleanup.json](cleanup.json).

## Limitations

The comparison uses one pinned virtual CPU, fresh processes, and fifteen measured batches per lane. A/B/A/B follow-up narrows process noise for the 256-reference lane but does not establish a statistical tail bound. The scaling inputs intentionally target the measured repeated-lookahead family. The evidence supports this bounded optimization and preservation checks; it does not establish a general parser throughput, whole-formula complexity, or package-level performance guarantee.
