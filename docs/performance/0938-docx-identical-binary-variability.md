# 0938 — identical-binary DOCX route variability

This diagnostic compares two labels for the same executable, rather than a
production candidate. Reader-pair p50 flags appeared in 0935, 0936 and 0937 at
different blocks. The purpose is to measure variation under unchanged code
and inputs before designing another comparison. It cannot retroactively
qualify the rejected 0937 cohort or prove that its candidate caused no
regression.

## Prospective design

The fixed cohort contains 96 fresh processes: four ingress routes, twelve
paired blocks per route and two identical control labels. Both labels use
the retained 0933 normal-release executable, SHA-256
`8a81df641151e5ed06b8f700c5d4cf068d0fee9c9e131fa147bf707a799c4c66`,
with the same historical source revision and command arguments except the
report output path. Report labels and paths have equal lengths. No candidate
binary, allocator-instrumented lane, rebuild or production edit is included.

Each child uses five warmups and fifty measured samples on CPU 12. Route order
reverses on alternating blocks, and control-label execution order alternates
too. Every route has twelve A/B pairs. The environment, corpus, semantic and
output identities remain fixed. Original terminal records precede analysis;
failed children cannot be selectively retried or replaced.

Per-route summaries retain pooled percentiles separately from paired ratios.
Paired median A/B p50 ratios use 10,000 bootstrap draws with seed 938. Every
pair retains p50, p95, p99, spread and RSS ratios. Counts above 1.05 are reported
in both directions because the control labels are arbitrary. The one-sided
B/A counts are also retained for comparison with earlier diagnostic flags.

Whole-process `/usr/bin/time -v` records provide user/system CPU time, major
and minor faults, voluntary and involuntary context switches, and maximum
RSS. These process-wide counters are descriptive companions to the timed
scenario; they do not isolate a parser phase or establish causality.

Each fresh process runs all five warmups and fifty samples in its own shared
allocator state. Preparation includes the owned-reader input clone; open,
query, edit, commit and publication have separate timing boundaries.
Correctness checks and report serialization fall outside those timings but
inside the whole-process RSS observation. Phase summaries can locate timing
differences without attributing their cause. Affinity and alternating order
do not eliminate address-layout, allocator, frequency or host-load effects.

## Identity and interpretation

Review HEAD is `3e4ea43db4a5c2bd2d2fd8e16ba18d9c887b7809`. Its changes from
`e3b5acdd91280b9926f94a61e5d736ad8063422e` are performance documentation and
evidence only. Runtime and harness equivalence to the retained binary remains
separately bound; the current review commit is not presented as its historical
build revision. No existing test or benchmark binary is committed.

The benchmark uses synthetic ordinary section fragments, with no
`w12:footnoteColumns` extension. Results apply to this executable, fixture,
machine, protocol and capture window. They do not establish variability for
all Office documents, other ingress implementations or allocator-instrumented
measurements.

Observed flag counts are not population false-positive probabilities.
Twelve serial pairs per route are a limited sample, and alternating order
does not guarantee independence or stationarity on a shared host. Bootstrap
intervals summarize the observed paired cohort under their resampling
assumptions; individual results and order effects remain visible.

If a same-binary pair crosses 5%, it demonstrates that this protocol can
produce such a flag without a code difference. It does not identify whether
allocator state, address layout, scheduling, faults or another factor caused
it. If no pair crosses 5%, that does not establish zero variability or prove
that the earlier candidate was responsible. Neither outcome changes the
0937 rejection or authorizes an optimization here.

Accepted ADR preservation, bounds, deterministic-output and evidence
requirements remain in force. No production API, ownership, validation,
security or execution behavior changes. iWork is excluded.

## Preflight

Before any formal child started, runner controls checked the fixed schedule,
identical normalized commands and clean environment, report identity, and
phase-sum rejection. An independent checker exercised a complete synthetic
96-child packet and rejected changed warmup arguments and a changed phase
sum. Synthetic packets are checker fixtures, not performance observations.

The final prospective manifest is external
`formal-manifest-unauthorized-v5.json`, SHA-256
`45656ab6002af33f16e4274c4fdfb41ea73188db462f5d99c45cc6ad44072e43`.
Earlier preflight revisions remain external: schedule assertions, summary
definitions, bootstrap intervals and replay provenance wording were corrected
before capture. The independent checker was also adapted to the actual
single-lane A/B schema, including complete command equality, process metrics,
and alternating label order. The final preflight checker SHA-256 is
`aa4a1f6a6a2d17e5442893af52f7784d58607a7a0b6c9cff2655960859bea3b6`.
Compact tooling and independent-preflight receipts bind the final inputs and
synthetic controls. None of those preflight revisions reran a formal child.

## Capture and offline validation incident

The single formal cohort completed all 96 children with exit code zero, no
deadline or process-group failure, and no retries or recaptures. However, the
frozen runner marked every child failed after report validation: its result
construction requested `output_sha256` from the reused 0937 validator's
return value. That validator checks output identity but does not return that
key. Every original validation record therefore contains
`KeyError: 'output_sha256'`; the original capture state is **failed**.

The synthetic preflight exercised report checking and packet auditing but
did not execute the runner's post-validation result construction, so it
missed this integration error. Original scripts, manifests, reports,
terminals, validation failures and state remain unchanged. Any recovered
analysis is a separately versioned offline revalidation of those same raw
reports, not a successful original runner validation or a new capture.
Moved-packet checks also required an offline path-normalization correction:
the original terminal command correctly retains its original absolute report
path. Earlier checker failures remain recorded; relocation must not rewrite
that command history.

## Recovered descriptive results

All 96 raw reports pass offline validation of the fixed corpus, output,
sample count, route, phase sums and correctness gates. The original and moved
packets retain the same 678 files, 11,554,110 bytes and tree SHA-256
`46d163d8b4b7dbb977cfb4a830c0cace69371abecedb8b3b8b112b33cd9376e6`.
Each route has twelve pairs and 600 measured samples per label.

| Route | Median paired p50 B/A | Bootstrap 95% interval | B/A > 1.05 | A/B > 1.05 |
| --- | ---: | ---: | ---: | ---: |
| Generic read-at | 1.000356 | [0.999552, 1.001300] | 1/12 | 0/12 |
| Owned vector | 0.999880 | [0.998645, 1.001156] | 0/12 | 0/12 |
| Owned reader | 0.997161 | [0.937461, 1.003149] | 2/12 | 4/12 |
| Short range | 1.000696 | [0.999593, 1.001077] | 0/12 | 1/12 |

Reader p50 crosses the descriptive 5% threshold in six pairs, with excursions
in both directions despite identical code. Generic read-at and short-range
each have one crossing, so excursions are not exclusive to the reader route.
These counts describe the observed cohort; they are not estimates of a
population false-positive rate. Near-unity median ratios do not hide the
individual excursions.

Tail, spread and RSS counts are also retained in both directions. Each table
cell below is **B/A count / A/B count**, out of twelve pairs; a crossing means
the ratio exceeds 1.05.

| Route | p95 | p99 | Spread | Whole-process RSS |
| --- | ---: | ---: | ---: | ---: |
| Generic read-at | 1 / 0 | 0 / 1 | 6 / 6 | 0 / 0 |
| Owned vector | 0 / 0 | 0 / 0 | 4 / 8 | 0 / 0 |
| Owned reader | 1 / 4 | 2 / 3 | 5 / 7 | 0 / 2 |
| Short range | 0 / 1 | 0 / 0 | 7 / 4 | 0 / 0 |

Spread ratios are particularly sensitive to individual extreme samples and
are not substitutes for latency ratios. Reader whole-process RSS crosses in
two A/B pairs, reaching 1.090179; no other route crosses the RSS threshold.
These observations do not identify a memory phase or mechanism.

For the six reader pairs with total-p50 crossings, selected phase-median B/A
ratios are:

| Block | Preparation | Open | Publication |
| --- | ---: | ---: | ---: |
| 3 | 0.1375 | 1.0019 | 0.9914 |
| 6 | 0.1404 | 1.0073 | 0.9892 |
| 9 | 0.1417 | 0.9934 | 0.9842 |
| 10 | 6.9247 | 1.0036 | 1.0119 |
| 11 | 6.9084 | 0.9785 | 1.0085 |
| 12 | 0.9476 | 0.8922 | 0.8412 |

Five pairs have a large preparation difference, while block 12 has a
different pattern involving open and publication. Preparation includes the
input clone and other setup. This distinction makes preparation a useful
follow-up target, but does not isolate cloning or allocator behavior as the
cause. Phase medians need not sum to the median total, and whole-process
fault and CPU counters cannot assign those differences to a phase.

## Decision and next action

Retain the descriptive control and the complete validation-failure history.
Adopt no optimization, change no existing threshold, and keep 0937 rejected.
The control demonstrates that same-code excursions can occur under this
protocol, especially on the reader route. It does not establish that any
particular historical excursion was noise.

Before another candidate comparison, exercise the complete child-to-validation
integration, including result construction, in its preflight. Use the retained
phase and process observations to formulate a prospective reader-variability
investigation with explicit controls. Any attribution to allocator state,
address layout or faults needs new evidence; this cohort supplies no such
causal conclusion. Office-wide work in `docs/GOAL.md` remains unfinished.

## Evidence and reproduction

The compact [independent result](results/change-0938/independent-normal-aa-audit-v2.json)
contains every pair, both directional flags, phase summaries, process metrics,
order grouping and external artifact descriptors. It is 672,666 bytes,
SHA-256 `9000bccdcbd3cf8e7a6978d61cdd43de74cb83c972a0608b263b5d67e14d21dc`.
The final postcapture checker is external `independent-normal-aa-audit-v3.py`,
SHA-256 `184b32eba43a114d02799f7f4c1ca303dd6a5bd0d2b5c1fe1a1927c80e0b9861`.
Its original/moved result is 564,815 bytes, SHA-256
`4c604440cbd401b2188e80a482bdc79b9c0f7961d028db64856e6627e2008eeb`.

The [postcapture receipt](results/change-0938/comparison-postcapture-validation-v1.json)
binds the authorized manifest, original outer terminal, failed packet state,
offline recovery scripts, hashes and commands. The
[independent custody review](results/change-0938/independent-normal-aa-recovery-audit-v1.json)
records the separate raw-report and terminal checks. Source/build equivalence
is reused from retained receipts; no new build or production test run is
claimed for this documentation and evidence batch.

Offline reproduction, using a new output filename and the retained external
root, requires no benchmark execution:

```sh
python3 -B /home/zhuhe/litchi-performance-0938/independent-normal-aa-audit-v3.py \
  --packet /home/zhuhe/litchi-performance-0938/formal-packet-v1 \
  --peer-packet /home/zhuhe/litchi-performance-0938/formal-packet-v1-moved \
  --output /home/zhuhe/litchi-performance-0938/independent-reproduction.json
```

The external canonical archive is
`/home/zhuhe/litchi-performance-0938/archives/canonical-evidence-v1.tar.zst`,
604,061 bytes, SHA-256
`2fc5e91775e246002fdfb40bf937ef10fe70049336135f3b2a6521ca712491d3`.
An explicit 4,886-file allowlist covers 72,640,788 logical bytes, including
both unchanged formal packets, synthetic controls, failure/recovery history,
checkers and a compact-evidence snapshot. Verification checked every member,
file hash, regular-file type, exact member list and compressed integrity.
No source, build, target or binary tree is archived. Archive-generation and
cleanup receipts are retained separately from this frozen snapshot.
The [archive binding](results/change-0938/canonical-evidence-archive-v1.json)
records the regeneration commands and external manifest descriptors. The
[cleanup receipt](results/change-0938/canonical-evidence-cleanup-v1.json)
records removal of 76,258,515 bytes of uncompressed tar and owned Python
cache after verification. Raw evidence and the compressed archive remain;
the shared checkout target and other batches were untouched.
