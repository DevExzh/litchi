# Bitwise scalar evaluator performance

This batch measures the five OpenFormula §6.6 bitwise functions: `BITAND`,
`BITLSHIFT`, `BITOR`, `BITRSHIFT`, and `BITXOR`. The comparable group runs
against the baseline and candidate binaries and contains existing scalar
controls and failures plus representative already-supported logical, lazy, and
text workloads. The bitwise group is candidate-only capability evidence. A
baseline run of that group would be a refusal control because the baseline
catalog recognizes these names while its scalar evaluator reports
`Unsupported(Function)`.

The [harness runner](bitwise-harness/run.py) and
[Rust workload](bitwise-harness/src/main.rs) define the corpus, phase
boundary, exact expected values, and allocator observations. The captured raw
data is [baseline comparable](baseline/comparable/raw.csv), [candidate
comparable](candidate/comparable/raw.csv), and [candidate bitwise](candidate/bitwise/raw.csv).

## Measurement boundary

Each selected case runs in three phases:

* `parse` parses a fresh expression on every operation.
* `evaluate` parses once before timing and evaluates the retained expression on
  every operation.
* `parse-evaluate` parses and evaluates on every operation.

The execution context and limits are created before the timed region and kept
alive through `live_after` and allocator counter reads. The runner uses three
warmups and 15 measured iterations, pins each process to CPU 2, and records
the exact command, timestamps, stdout, stderr, `/usr/bin/time -v` RSS output,
and status. Scale lanes batch 64, 32, 8, or 2 operations for 64, 256, 1024,
or 4096 terms. Fixed controls and edge cases batch 128 operations.

Elapsed values are for a complete measured batch; normalized per-operation
values divide by the recorded repeat count. `evaluate` and `parse-evaluate`
consume every returned scalar with a checksum in the timed region, so those
values include result consumption and do not represent pure evaluator
instruction latency. Expected scalar values and typed formula errors are
checked during preflight before timing. A matching aggregate checksum is only
a secondary guard and does not establish semantics by itself.

## Corpus and semantic checks

The comparable corpus has 39 cases in each phase (117 rows per revision): 16
scalar scale lanes, two scalar boolean controls, seven typed refusal controls,
and 14 logical/lazy/text controls. The latter cover representative `AND`,
`OR`, `XOR`, `NOT`, `IF`, `IFERROR`, and `IFNA` workloads, an unused heavy
`IF` branch, an unselected reference, and escaped/concatenated text.

The candidate bitwise corpus has 31 cases in each phase (93 rows). Twenty
scaled cases use a flat sum of shallow calls at 64, 256, 1024, and 4096 calls:

| Family | Call | Value per call |
| --- | --- | ---: |
| AND | `BITAND(6;3)` | 2 |
| OR | `BITOR(6;3)` | 7 |
| XOR | `BITXOR(6;3)` | 5 |
| left shift | `BITLSHIFT(5;2)` | 20 |
| right shift | `BITRSHIFT(20;2)` | 5 |

The 11 fixed cases cover text and fractional coercion, negative operands, the
48-bit result bound, an overflowing left shift, negative shift behavior, invalid
text, wrong arity, and lazy selection. The expected candidate outcomes are:

| Case | Expected candidate outcome |
| --- | --- |
| `bitwise-coerce-text` | Number 2 |
| `bitwise-coerce-fraction` | Number 2, truncating 6.5 toward zero |
| `bitwise-error-negative` | `ScalarError::Number` |
| `bitwise-error-overflow` | `ScalarError::Number` |
| `bitwise-error-shift` | `ScalarError::Number` |
| `bitwise-shift-negative` | Number 2 |
| `bitwise-error-numeric` | `ScalarError::Value` |
| `bitwise-error-arity` | `ScalarError::Value` |
| `bitwise-lazy` | Number 2; the reference branch is not resolved |
| `bitwise-lazy-error` | Number 7 through `IFERROR` |
| `bitwise-lazy-selected` | Number 7 from the selected `BITOR` branch |

All 117 baseline rows, 117 final-candidate comparable rows, and 93 final-
candidate bitwise rows completed with status 0. Preflight exact value checks
passed for every evaluated row. The seven comparable failure controls retained
their expected unsupported reference/array/name, work, memory, object, and
cancellation outcomes. Every candidate bitwise row is an expected successful
evaluation at the harness level: invalid operands return typed formula values
(`ScalarValue::Error`) and therefore are checked as successful scalar results
rather than process-level refusals.

## Provenance

The baseline is commit `c32455df83ea67f840a6d4f8da8953f113b590f4`, with 211
source entries and evaluator source hash
`cb47a3c6a6a0169f225d0e4b979fe5ec9ebf629bc1a5fdef86dca18f9c3c8efa`. Its
profile binary was verified before cleanup with SHA-256
`cc03701b74689d33f1f90cbcdca7f8f5ad6d3aa86ea6affdfc0a212a49a22de3` and
1,249,608 bytes. The [baseline source/build/binary receipts](baseline/)
record the exact inputs and commands; the saved ELF was removed after its
digest was independently verified.

The final outlined candidate has 212 source entries, evaluator source hash
`a150855aa43b64753ba90ba98322b213d1dfa0bc33b0e317da0f3e7e11a959db`, and
bitwise regression-test source hash
`a64fcb6976541001c3ce33a51e9bfa44283717db14ea61cf853dd5325a65d8b8`. Its
profile binary was verified before cleanup with SHA-256
`63ba1ca4f10b4ec8c62d0ef506c6ee4f866676ddb6a7a74173bae26239aa1599` and
1,259,168 bytes. The [candidate source/build/binary receipts](candidate/)
bind the captures to that snapshot; the saved ELF was removed after its
digest was independently verified. The four harness hashes are retained in
both [baseline](baseline/harness-sha256.txt) and
[candidate](candidate/harness-sha256.txt).

The fresh comparable sequences ran at 13:15:07.883–13:15:09.692 UTC on
2026-09-13. The bitwise sequence ran at 13:13:33.386–13:13:35.280 UTC from
the same outlined candidate binary. The [environment receipt](environment.json)
records Rust/Cargo versions, CPU, kernel, and the shared-host limitation.

The candidate before the focused `apply_bitwise` outline is retained under
[before-outline](before-outline/). Its comparable ABAB screening had persistent
`evaluate/control-utf8-left-1024` and `evaluate/control-utf8-left-4096` p50
regressions of +8.60% and +11.09%, with corresponding `parse-evaluate` flags
of +7.04% and +9.00%. The outline's first capture removed those evaluate
flags, and the final repeat confirmed them at -0.91% and +0.00% for
`evaluate`, and +0.84% and +0.04% for `parse-evaluate`. These captures were
made in different host intervals, so this is a disposition of the tested
change rather than a causal speedup claim.

## Comparable results

The fresh outlined candidate capture's median p50 delta across the 39 cases
was +0.60% for `parse`, -0.14% for `evaluate`, and +0.45% for
`parse-evaluate`. The final repeat selected 46 case/phase pairs: all fresh p50/p95/p99/RSS
flags plus the four previously regressed text lanes, retained even when they
no longer crossed the screening limit. Its p95/p99 flags were mostly
isolated samples from fixed failures, escaped text, and large logical inputs;
the final [four-round interleaved repeat](abab/summary.json) adjudicated those
screening observations.

The final repeat left four p50 latency flags, all in `parse`:

| Phase/case | Baseline p50 ns/op | Candidate p50 ns/op | Delta |
| --- | ---: | ---: | ---: |
| `parse/control-utf8-left-64` | 157.3 | 165.5 | +5.22% |
| `parse/control-utf8-left-256` | 270.5 | 289.4 | +7.00% |
| `parse/logical-iferror-4096` | 561,379.8 | 607,800.0 | +8.27% |
| `parse/logical-ifna-4096` | 526,512.2 | 559,337.5 | +6.23% |

No `evaluate` or `parse-evaluate` p50 flag remained in the repeat. The repeat
still contained p95/p99 observations in both directions, including large
swings for `parse/control-utf8-left-1024`, fixed failure controls, and escaped
text. These tails came from a 15-sample per-lane distribution and are
screening observations rather than stable tail regressions. No RSS delta
exceeded 5% in the interleaved repeat.

Every comparable row had identical p50 allocation calls, requested bytes,
released bytes, live-after values, and raw peak-live deltas between baseline
and final outlined candidate. This establishes no tracked heap-allocation
change in the comparable corpus. Process RSS remains a separate whole-process
observation.

## Candidate-only bitwise results

The table reports final outlined-candidate `evaluate` p50 time per operation
after one parse, while consuming the scalar result in the timed region. It is
capability evidence and has no baseline timing counterpart.

| Family | 64 calls | 256 calls | 1024 calls | 4096 calls |
| --- | ---: | ---: | ---: | ---: |
| `BITAND` | 15,303 ns | 56,968 ns | 222,030 ns | 884,860 ns |
| `BITOR` | 15,301 ns | 57,956 ns | 223,050 ns | 888,750 ns |
| `BITXOR` | 15,627 ns | 58,919 ns | 227,734 ns | 904,305 ns |
| `BITLSHIFT` | 15,436 ns | 60,953 ns | 237,595 ns | 930,724 ns |
| `BITRSHIFT` | 15,671 ns | 59,383 ns | 231,315 ns | 922,874 ns |

For these scale lanes, `evaluate` allocation calls are 12, 14, 16, and 18
per operation at 64, 256, 1024, and 4096 calls. Requested bytes are 12,944,
49,808, 197,264, and 787,088 per operation. The raw measured batch
`peak_live_delta` is 6,672, 25,104, 98,832, and 393,744 bytes respectively;
it is a maximum and is not divided by repeat. `parse` uses 79, 275, 1,047,
and 4,123 allocation calls per operation, while `parse-evaluate` uses 91,
289, 1,063, and 4,141. These results show input-scaled work and bounded
observations for flat shallow calls; they do not characterize deeply nested
expressions.

The fixed candidate `evaluate` cases returned their exact preflight values.
They measured about 415–676 ns/op, with four or five allocation calls and
400–496 requested bytes per operation. Lazy `IF` and `IFERROR` cases preserved
the unselected branch behavior, including the reference-containing branch.
The invalid negative, overflow, and overflowing-shift inputs produced
`ScalarError::Number`; invalid text and wrong arity produced
`ScalarError::Value`. These outputs are semantic checks, not a claim that an
invalid formula is a process refusal.

## Hardware counters

The [counter receipts](perf-stat/) use
`cycles,instructions,branches,branch-misses` with whole-process scope,
including setup, preflight, warmups, and measured loops. All seven captures
completed with status 0 and carry binary, stdout, and stderr hashes.

| Binary and case | Cycles | Instructions | Branches | Branch misses |
| --- | ---: | ---: | ---: | ---: |
| Baseline `control-utf8-left-4096` evaluate | 6,277,279 | 7,707,089 | 1,939,154 | 23,310 |
| Candidate `control-utf8-left-4096` evaluate | 4,984,186 | 7,715,541 | 1,939,942 | 20,332 |
| Baseline `logical-iferror-4096` parse | 99,114,339 | 606,343,108 | 156,665,306 | 28,199 |
| Candidate `logical-iferror-4096` parse | 108,913,537 | 606,387,783 | 156,672,128 | 31,817 |
| Candidate `BITAND-4096` evaluate | 154,705,776 | 530,467,041 | 83,134,117 | 26,507 |
| Candidate `BITLSHIFT-4096` evaluate | 158,231,099 | 549,315,677 | 86,393,867 | 26,353 |
| Candidate overflowing-shift error evaluate | 3,048,960 | 3,423,710 | 642,300 | 19,397 |

The logical-iferror pair has nearly identical instruction counts but about
9.9% more candidate cycles in this whole-process capture, consistent with the
remaining parse-only p50 cost in that lane. The counters do not isolate one
evaluator operation or establish causality. The [binary size receipt](binary-sizes.json)
records 1,249,608 bytes for baseline and 1,259,168 bytes for candidate; size
alone does not explain process RSS variation.

## Limits and disposition

The comparable group is a regression control; the bitwise group is
candidate-only capability evidence and must not be combined with it into a
speedup number. RSS is a process-level maximum and varies with process startup
and shared-host state. p95/p99 values are based on 15 measured samples per
lane. Timings include checksum consumption, and raw `peak_live_delta` values
are batch maxima. Exact preflight values, typed error checks, and status checks
support semantic correctness; timing checksums alone do not.

The focused outline removed the earlier large UTF-8 evaluate flags. The final
interleaved repeat retains four parse p50 flags, including both 4096-term
logical parser lanes, and the counter pair shows a corresponding tested
whole-process cost for `logical-iferror-4096`. The evidence supports the
tested bitwise behavior, input-scaled bounded allocator observations, and no
tracked comparable heap growth. It does not support a general performance
improvement claim; the remaining parse latency flags and whole-process RSS
variability are disclosed for future work.

The >5% checks are review triggers under `docs/GOAL.md`. This batch accepts
the four disclosed parse-only costs for the added capability after removing
the existing text evaluation/end-to-end regressions. The 22 retained parser
symbols have identical names and sizes in baseline and final binaries, but
that does not prove identical layout, execution cost, or a cause for the
measured cycle difference. Further tuning should use fresh representative
measurements rather than arbitrary code placement changes.
