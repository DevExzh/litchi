# 0939 — capture validation integration and sample chronology

Batch 0938 established same-binary variability but exposed an integration
error: the capture runner requested an output digest absent from the report
validator's return value. Every child exited successfully, yet all original
validation records failed. Its synthetic controls had bypassed that result
construction. The failed packet and subsequent offline recovery are immutable
historical evidence.

This batch addresses the missing integration coverage and examines the
retained reader samples in chronological order. It does not rerun 0938,
qualify its original runner, alter the 0937 rejection, relax a threshold or
claim a production speedup. iWork is excluded.

## Evidence basis

Review HEAD is `6887be9087c5cd0ad981631fc8f8b367e3193ed2`. The retained
0938 original and moved packets each contain 678 files and 11,554,110 bytes,
with the recorded 0938 tree SHA-256
`46d163d8b4b7dbb977cfb4a830c0cace69371abecedb8b3b8b112b33cd9376e6`.
They contain 96 fresh-process observations, four routes and twelve A/B pairs
per route, with five warmups and fifty measured samples per process.

The harness's `statistics` function enumerates samples before sorting by
`(elapsed, original_index)`. Thus `sample_order[rank]` gives the original
index. Phase vectors are reordered with that same permutation. Chronological
reconstruction assigns `chronological[sample_order[rank]] = ranked[rank]`;
using the permutation directly as an inverse would be incorrect.

All six phases must still sum to each corresponding total after inversion.
The recorded sample order concerns the fifty measured iterations, not the
five discarded warmups. Whole-process CPU, faults and RSS remain whole-child
observations and cannot locate a memory event within those iterations.

## Scope and constraints

The accepted ADRs are unchanged from the preceding reviewed batches. This
work changes performance tooling and evidence only. Production ownership,
validation, preservation, security and public APIs remain unchanged. No new
benchmark binary, source snapshot or Office workload capture is required.
All large evidence stays outside the repository; committed evidence must
remain below the per-file limit.

## Chronological findings

All 96 non-identity permutations were inverted, retaining 4,800 measured
samples and their phase sums. An exploratory 100 µs preparation split
separates reader children with markedly different observed timings:

| Reader children | Count | Range of child preparation medians | Median of child medians |
| --- | ---: | ---: | ---: |
| At least 25/50 preparation samples ≥100 µs | 9 | 235.891–247.081 µs | 239.761 µs |
| Fewer than 25/50 preparation samples ≥100 µs | 15 | 32.451–39.600 µs | 34.451 µs |

The split was chosen for descriptive analysis after observing the data; it
is not a prospective threshold or production gate. There are 429 high
preparation samples among the 1,200 reader samples. Each majority-high child
has 42–50 such samples; three have all fifty. The pattern therefore includes
persistent differences through the measured sequence, rather than only
isolated extreme samples. It does not reveal what happened during warmups.

The high-child counts are five A and four B, and five A-first and four
B-first. Neither label nor alternating order uniquely identifies the group.
This small serial cohort cannot rule out time-varying host or process effects.

For the six reader pairs whose total p50 crossed 5%, mean differences provide
an additive phase accounting. Values below are B minus A in microseconds:

| Block | Total mean | Preparation mean | Open mean | Publication mean |
| --- | ---: | ---: | ---: | ---: |
| 3 | -221.329 | -207.716 | 1.704 | -12.223 |
| 6 | -219.982 | -198.461 | -8.793 | -10.020 |
| 9 | -195.190 | -172.678 | -1.368 | -15.417 |
| 10 | 221.040 | 192.483 | 1.689 | 22.598 |
| 11 | 186.586 | 197.056 | -12.726 | 5.463 |
| 12 | -272.697 | -11.185 | -41.198 | -214.677 |

The complete result includes query, edit and commit, and verifies that all
six mean differences sum to the total. Preparation accounts for 87–106% of
the net mean difference in five pairs; a share above 100% means other phases
offset part of it. Block 12 instead differs principally in publication and
open. These are arithmetic decompositions, not causal attributions.

Matching measured indices across independent processes is bookkeeping, not
matched physical events. Mean decomposition does not require that pairing;
chronological inversion is needed for within-child sequence analysis. Phase
medians remain nonadditive, and whole-process fault/CPU observations cannot
explain a particular phase without a dedicated measurement.

The first offline analysis attempt incorrectly expected each phase vector to
be independently sorted. It failed before producing results. Source review
corrected that assumption to shared elapsed-rank alignment; the successful
analysis and a separate checker then verified the inverse mapping and phase
sums for all samples. The first attempt's terminal and error log remain
available, but its overwritten script bytes were not separately frozen.
The final analyzer is hash-bound and reproducible. No capture was repeated.

## Bounded attribution follow-up

Source review confirms that both owned routes clone the 2,101,427-byte input
in preparation. Reader-only preparation also creates two read observers;
sink reservation is shared. During open, the vector route transfers its
existing allocation, whereas the generic reader route fills a new vector
through `read_limited`. That is an actual ownership and allocation difference,
but it does not by itself explain the preparation-time groups.

The [draft attribution protocol](results/change-0939/attribution-experiment-plan-v1.json)
uses matched vector/reader route observations, separate preparation/open
allocation segments, existing read counters and explicitly measured probe
overhead. Instrumented timings must remain diagnostic and separate from
normal-release qualification. The workload, clone placement, validation and
output identities must remain intact. The plan neither authorizes a capture
nor establishes an allocator, page-fault or memcpy cause.

## Capture integration change

[`tools/perf_child_validation.py`](../../tools/perf_child_validation.py)
provides a callable single-child adapter and a CLI. It accepts an explicit
command, environment, validator path and SHA-256, report contract and output
locations. It writes the child's terminal record before parsing the report
or invoking the semantic validator. After validation, it extracts the result's
nested output digest, checks the expected identity, and includes that digest
in the result returned to the capture driver. It no longer assumes that a
legacy validator returns a field merely because it checked that field.

Artifact paths must be distinct and unused. Oversized reports retain bounded
metadata and fail validation without reading the whole report. A descendant
left after the child exits triggers group cleanup and a failed result; cleanup
cannot turn that child green. There is no child retry.

Seven targeted integration tests pass. They demonstrate the original missing
return key with the real retained 0937 validator, then exercise the new CLI
with valid, missing-digest and changed-digest reports. Other tests check
terminal-before-validator ordering, duplicate output paths, oversized reports,
and cleanup of a deliberately orphaned child. The fixture children copy or
construct reports; they do not execute an Office workload or benchmark binary.

This POSIX process adapter is usable by future drivers, but does not replace
or alter the frozen 0938 driver, define a cohort, freeze a schedule, or provide
a performance qualification by itself. Callers remain responsible for corpus,
binary, source and schedule provenance. No Rust build or production regression
run is claimed for this Python-tooling change.

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B -m unittest -v tools.test_perf_child_validation
python3 -B tools/perf_child_validation.py --help
```

## Chronology evidence and reproduction

The [compact chronology result](results/change-0939/docx-chronology-summary-v3.json)
binds the input packets, source semantics, analyzer, failed/successful attempt
logs and independent checker. The full external result is
`/home/zhuhe/litchi-performance-0939/chronology-analysis-v2.json`,
1,301,641 bytes, SHA-256
`44e9f43df08641cd5a5ba4a2c5a7bfb060aeac508b1fd9a0b3fd5afd7bc04c8c`.
It is deliberately not committed. The final external analyzer is 33,483 bytes,
SHA-256 `b27a527f62a410eb88b4bffccc082d5c455e5031d52c9cccc60186b2f8b8112b`.

Reproduce the analysis into new external outputs with:

```sh
python3 -B /home/zhuhe/litchi-performance-0939/analyze-0938-chronology.py \
  --packet /home/zhuhe/litchi-performance-0938/formal-packet-v1 \
  --output /home/zhuhe/litchi-performance-0939/chronology-reproduction.json \
  --compact-output /home/zhuhe/litchi-performance-0939/chronology-reproduction-summary.json
```

This recreates the numerical analysis. The committed v3 summary additionally
records the subsequent independent checks and corrected interpretation text;
it is not claimed to be byte-identical to the analyzer's generated summary.

## Decision

Retain the checked child adapter and its integration tests as a measurement
infrastructure repair. The next capture driver must explicitly use and bind
that entry point; historical drivers remain evidence rather than corrected
executables. The [independent review](results/change-0939/independent-review-v1.json)
binds the final implementation and test results.

The chronology changes the next investigation: measure preparation and open
separately, preserve within-child sequences, and assess instrumentation cost
before interpreting a route contrast. Do not discard high-preparation
children, replace them with retries, or use this analysis to waive previous
candidate regressions. No optimization, threshold change or Office-wide
completion is claimed.
