# ODS formula tokenizer performance

This report compares the committed reference parser at `b88341bf4` with the
frozen tokenizer candidate. The candidate source passed the root gate receipt
before measurement. Its scoped implementation patch contains only
`formula.rs`, `reference.rs`, and the four-test tokenizer regression file.

The candidate reuses one compact identifier/cell scan, defers fallible cell
component copies until a complete coordinate is known, propagates non-syntax
errors from speculative cell parsing, and trims reference component bookkeeping.
The measurements below show the resulting cost across the established parser
family and the extended reference coverage.

## Method and scope

The unchanged standalone [harness](harness/) calls
`litchi_ods::codec::formula::FormulaParser` through its public parser API. Every
lane uses a fresh release process pinned with `taskset -c 2`, three warmup
batches, and fifteen measured batches. Ordinary lanes execute 1,000 parser calls
per batch; the 256-reference and 16 KiB lanes execute 128 calls per batch. The
instrumented `System` allocator records requests and bytes, while
`/usr/bin/time -v` records process maximum RSS.

The paired order was baseline comparable, candidate comparable, baseline
coverage, candidate coverage. The saved baseline executable was used after the
candidate build overwrote the shared Cargo release target. Raw p50 time and
allocator counts are divided by calls per batch. `Peak live delta` is the raw
maximum for a measured batch and is intentionally not divided. RSS includes
process startup and static data. Parser checksums, expected outcomes, process
statuses, commands, and timestamps are retained beside each raw CSV.

This measures token construction, reference parsing, error handling, and
allocation/destruction. It excludes package I/O, XML parsing, evaluation,
workbook resolution, publication, URI resolution, and networking. The fifteen
batch result is an engineering comparison rather than a statistical tail bound.

## Comparable lanes

All 12 baseline and 12 candidate lanes exited with status zero. The expected
success/refusal value and checksum fields match for every pair. Times are p50
nanoseconds per parser call; arrows are baseline → candidate.

| Case | Input bytes | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested bytes/call | Peak live delta (batch max bytes) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |

| parse-bracket-current-cell | 9 | 1000 | success | 173.82 → 173.15 | -0.4% | 3 → 3 | 458 → 458 | 458 → 458 | 2224 → 2268 |
| parse-bracket-sheet-cell | 18 | 1000 | success | 261.67 → 252.43 | -3.5% | 5 → 5 | 697 → 697 | 697 → 697 | 2268 → 2204 |
| parse-bracket-range | 29 | 1000 | success | 417.37 → 411.51 | -1.4% | 7 → 7 | 712 → 712 | 712 → 712 | 2268 → 2268 |
| parse-quoted-doubled-sheet | 19 | 1000 | success | 269.86 → 275.09 | +1.9% | 5 → 5 | 697 → 697 | 697 → 697 | 2268 → 2324 |
| parse-local-refs-256 | 1175 | 128 | success | 62512.79 → 60251.30 | -3.6% | 265 → 265 | 115671 → 115671 | 58775 → 58775 | 2204 → 2268 |
| parse-sum-unbracketed | 15 | 1000 | success | 287.47 → 274.28 | -4.6% | 5 → 5 | 468 → 468 | 468 → 468 | 2332 → 2256 |
| parse-vlookup-unbracketed | 26 | 1000 | success | 474.95 → 452.36 | -4.8% | 8 → 8 | 3172 → 3172 | 1828 → 1828 | 2220 → 2268 |
| parse-malformed-zero-row | 9 | 1000 | refusal | 174 → 175.05 | +0.6% | 4 → 4 | 101 → 101 | 76 → 76 | 2332 → 2204 |
| parse-malformed-missing-separator | 8 | 1000 | refusal | 264.90 → 260.90 | -1.5% | 6 → 6 | 138 → 138 | 94 → 94 | 2268 → 2204 |
| parse-malformed-unclosed-bracket | 8 | 1000 | refusal | 159.79 → 160.81 | +0.6% | 4 → 4 | 112 → 112 | 88 → 88 | 2324 → 2268 |
| parse-malformed-missing-row | 8 | 1000 | refusal | 216.88 → 213.07 | -1.8% | 5 → 5 | 139 → 139 | 114 → 114 | 2268 → 2268 |
| parse-malformed-bad-quote | 18 | 1000 | refusal | 162.99 → 163.28 | +0.2% | 4 → 4 | 122 → 122 | 88 → 88 | 2264 → 2320 |

The ordinary unbracketed SUM and VLOOKUP lanes improve by 4.6% and 4.8%;
the 256-reference lane improves by 3.6%. Bracketed and malformed lanes vary
from a 3.5% improvement to a 1.9% regression. Allocation calls, requested
bytes, and peak live deltas are identical across all 12 pairs. RSS varies with process conditions and remains
within 5.5% in the individual malformed lane (2,332 → 2,204 KiB); this is
startup RSS rather than retained parser state.

## Extended reference coverage

All 16 baseline and 16 candidate lanes exited with status zero. Valid lanes
succeeded and the over-limit IRI lane refused as expected on both sides; the
outcome and checksum fields match pairwise.

| Case | Input bytes | Calls/batch | Outcome | p50 ns/call | Δ | Alloc calls/call | Requested bytes/call | Peak live delta (batch max bytes) | Max RSS KiB |
| --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |

| coverage-source-cell | 28 | 1000 | success | 326.55 → 307.99 | -5.7% | 5 → 5 | 717 → 717 | 717 → 717 | 2332 → 2268 |
| coverage-source-range | 32 | 1000 | success | 383.64 → 378.57 | -1.3% | 6 → 6 | 722 → 722 | 722 → 722 | 2268 → 2268 |
| coverage-empty-source | 12 | 1000 | success | 215.38 → 211.11 | -2.0% | 4 → 4 | 685 → 685 | 685 → 685 | 2264 → 2324 |
| coverage-unicode-escaped-source | 49 | 1000 | success | 430.74 → 435.76 | +1.2% | 5 → 5 | 758 → 758 | 758 → 758 | 2224 → 2204 |
| coverage-whole-columns | 11 | 1000 | success | 263.10 → 245.10 | -6.8% | 5 → 5 | 685 → 685 | 685 → 685 | 2268 → 2328 |
| coverage-whole-rows | 11 | 1000 | success | 184.04 → 175.91 | -4.4% | 3 → 3 | 683 → 683 | 683 → 683 | 2268 → 2256 |
| coverage-cross-sheet-range | 25 | 1000 | success | 364.27 → 351.08 | -3.6% | 6 → 6 | 487 → 487 | 487 → 487 | 2268 → 2268 |
| coverage-nested-inherited | 21 | 1000 | success | 402.73 → 402.61 | -0.0% | 8 → 8 | 741 → 741 | 741 → 741 | 2196 → 2224 |
| coverage-ref-error | 11 | 1000 | success | 133.97 → 130.73 | -2.4% | 3 → 3 | 683 → 683 | 683 → 683 | 2268 → 2268 |
| coverage-colon-sheet-1k | 1033 | 1000 | success | 1241.38 → 1423.35 | +14.7% | 4 → 4 | 2506 → 2506 | 2506 → 2506 | 2268 → 2264 |
| coverage-colon-sheet-4k | 4105 | 1000 | success | 4241.79 → 5039.45 | +18.8% | 4 → 4 | 8650 → 8650 | 8650 → 8650 | 2332 → 2268 |
| coverage-colon-sheet-16k | 16393 | 128 | success | 16060.70 → 19276.27 | +20.0% | 4 → 4 | 33226 → 33226 | 33226 → 33226 | 2268 → 2348 |
| coverage-source-iri-1k | 1036 | 1000 | success | 4814.99 → 4863.45 | +1.0% | 5 → 5 | 2733 → 2733 | 2733 → 2733 | 2224 → 2240 |
| coverage-source-iri-4k | 4108 | 1000 | success | 18538.50 → 18618.19 | +0.4% | 5 → 5 | 8877 → 8877 | 8877 → 8877 | 2160 → 2224 |
| coverage-source-iri-16k | 16396 | 128 | success | 150391.41 → 151626.17 | +0.8% | 5 → 5 | 33453 → 33453 | 33453 → 33453 | 2204 → 2264 |
| coverage-source-iri-over-16k | 16397 | 128 | refusal | 81992.88 → 108853.86 | +32.8% | 6 → 6 | 16657 → 16657 | 16437 → 16437 | 2160 → 2268 |

Short source-qualified and axis/reference forms improve by up to 6.8%. The
large colon-rich sheet-name lanes regress by 14.7%, 18.8%, and 20.0% at 1, 4,
and 16 KiB. The over-limit source-IRI refusal lane is 32.8% slower while still
refusing before accepting the oversized component. Allocation calls, requested
bytes, and peak live deltas remain identical for every coverage pair. The
largest observed RSS change is the refusal lane, 2,160 → 2,268 KiB (+5.0%).

These long-input regressions are observed candidate results. This harness
does not isolate whether their cost arises in reference parsing, annotation
handling, or code generation, so no causal attribution or speedup claim is
made for the rich lanes.

## Hardware counters

`perf stat` used the same CPU pin, warmups, iterations, and binaries as the
paired parser runs. Counters include process setup, warmups, and output, so they
are diagnostic. Values are totals divided by parser calls (18,000 for SUM and
2,304 for 256 references).

| Lane | Side | Cycles/call | Instructions/call | Branches/call | Branch misses/call | Cache misses/call | IPC |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| SUM unbracketed | baseline | 1,585.3 | 3,242.6 | 636.5 | 1.3 | 3.6 | 2.045 |
| SUM unbracketed | candidate | 1,514.3 | 3,133.8 | 605.0 | 1.2 | 3.3 | 2.069 |
| 256 local references | baseline | 107,782.4 | 318,218.2 | 64,879.3 | 104.1 | 39.0 | 2.952 |
| 256 local references | candidate | 96,824.3 | 279,896.6 | 55,047.9 | 59.2 | 33.3 | 2.891 |

The candidate reduces the paired counter totals by approximately 4.5% cycles
and 3.4% instructions for SUM, and 10.2% cycles and 12.0% instructions for
256 references. The counter receipts are [paired/perf-stat-sequence.json](paired/perf-stat-sequence.json), with raw stderr and stdout beside them.

## Symbol inspection

The release symbol receipts are [baseline-formula-symbols.txt](candidate/baseline-formula-symbols.txt) and [candidate-formula-symbols.txt](candidate/candidate-formula-symbols.txt). The candidate `parse_with_limits` symbol is 8,203 bytes versus 6,955 bytes in the baseline; `try_parse_cell_ref` is 2,512 versus 2,600 bytes, and reference `parse_body` is 3,411 versus 3,571 bytes. `parse_number` remains 4,221 bytes.

The candidate still contains private `copy_formula_component`,
`reference::copy_component`, and `Parser::bump_component` symbols. Disassembly
shows helper calls remain on legacy/error and reference paths, so the candidate
should not be described as completely helper-free. `bump_component` shrinks
from 262 to 99 bytes, and the compact scan reduces the measured common-path
work; the long component lanes show the tradeoff directly. The scoped formula/reference disassembly
receipts are [candidate-formula-disassembly.txt](candidate/candidate-formula-disassembly.txt) and [baseline-formula-disassembly.txt](candidate/baseline-formula-disassembly.txt).

## Correctness and resource gates

The isolated candidate checks passed 48 formula unit tests, four tokenizer
regressions, seven function-catalog tests, and ten reference integration tests.
Their commands and logs are in [isolated-checks.json](candidate/isolated-checks.json). The source manifest covers all eight files in the root gate receipt at [source-sha256.json](candidate/source-sha256.json), and the replayable zero-context patch is [candidate.patch](../candidate.patch) with its digest in [candidate-patch-sha256.txt](../candidate-patch-sha256.txt). The patch was checked and applied to a clean b883 archive, and the resulting three files matched the frozen checkout byte-for-byte.

The release binaries are identified by [baseline/binary-sha256.txt](baseline/binary-sha256.txt) and [candidate/binary-sha256.txt](candidate/binary-sha256.txt). Final paired run metadata and raw checksums are retained in [paired/sequence.json](paired/sequence.json), [paired/raw-sha256.txt](paired/raw-sha256.txt), [baseline](baseline/), [candidate](candidate/), [baseline-coverage](baseline-coverage/), and [candidate-coverage](candidate-coverage/). Temporary worktree, Cargo target, baseline snapshot, and executable copies were removed after hash verification; sizes and post-delete checks are in [cleanup.json](cleanup.json).

## Limitations

The run uses one pinned virtual CPU and a small fixed batch count. RSS is a
process-level maximum, and `perf stat` includes launcher/setup work. The rich
long-component regressions are measured and disclosed above. The evidence
supports semantic and allocation-preservation checks plus bounded comparative
observations; it does not establish a general production throughput guarantee.
