# 0504: reuse direct ODG transition values during one resolution pass

ODG open resolved every page's style by scanning the full source XML again,
even when those pages referenced the same style. The fresh metadata-large
Callgrind profile attributed 557,394,708 inclusive instruction references to
`find_transition_style_owner`, 30.52% of the whole child. This motivated a
private operation-local cache of successful **direct** transition values.
Inheritance still traverses each page's chain: depth and cycle checks, style
shadowing, duplicate detection, and lazy error ordering remain unchanged.
Neither errors nor fully inherited results are cached. Names borrow the
already parsed definitions; cached values are dropped when resolution ends.

The prior turn was progress: 0503 reconciled retained evidence and prioritized
this measured regression. This batch changes the production parser, adds
focused tests, and retains a new comparison rather than treating the audit as
completion of the broader program.

## Protocol and retained evidence

[The evidence directory](../results/change-0504/) contains the unchanged 0502
probe's four deterministic corpora, raw samples, process RSS receipts, source
and binary hashes, bootstrap summary, exact candidate patch, and profiling
traces. Before is clean `aeb9e229ae7129a08a77f3db85ee4823185cbe69`; after
adds only the recorded production patch. Both use the committed probe lockfile,
Rust/Cargo 1.95.0 and the same release build command. A final rebuild matches
the timed candidate binary exactly. The AMD EPYC 9R45 host and environment
are recorded in `environment.json`.

A 400-sample pilot preceded implementation acceptance. The formal capture runs
serial A1/B1/B2/A2 blocks pinned to CPU 2, with 25 warmups and 200 samples per
corpus per child: 16 children and 3,200 measured samples. This is two reversed
repeat pairs, not promotion into the separate strict claim registry. Fixture
construction and hashing are outside the timer; the timer includes owned-byte
open, allocation/drop, the probe's page/shape traversal and its checksum check.
That checksum covers the original traversal fields, not all transition or
auxiliary metadata. The focused tests and source review supply transition
correctness evidence. Corpus identities and traversal checksums match in all
formal reports. These are generated owned-byte inputs, not native-producer,
filesystem, range-provider, edit/save, or concurrency measurements.

## Results

Open-plus-traversal p50 milliseconds before → after:

| Corpus | R1 | R2 |
| --- | ---: | ---: |
| plain-small | 2.139 → 2.129 (−0.43%) | 2.145 → 2.123 (−1.00%) |
| plain-large | 66.078 → 65.259 (−1.24%) | 65.844 → 65.587 (−0.39%) |
| metadata-small | 1.951 → 1.760 (−9.77%) | 1.959 → 1.760 (−10.15%) |
| metadata-large | 38.340 → 25.139 (−34.43%) | 37.931 → 24.914 (−34.32%) |

Metadata-large p95 falls 34.38%/34.34% and p99 falls 34.25%/33.96%.
Metadata-small R1 p99 rises 1.57%, while R2 falls 15.09%; both observations
remain in the summary. None of the eight paired rows has an adverse change
above 5% in p50, p95, p99, mean, p50 input throughput, or whole-child RSS.
RSS changes range from −3.51% to +1.42%; these process maxima include fixture
setup and report serialization and are not operation-local memory values.

The deterministic 10,000-resample within-run p50 intervals are retained in
`summary.json`. Metadata-large intervals are −34.47% to −34.39% (R1) and
−34.40% to −34.24% (R2). These estimate within-child sampling uncertainty only;
they do not measure host-to-host or day-to-day variability. Plain differences
are small and are not presented as a separate practical improvement.

Linux disallows hardware counters (`perf_event_paranoid=4`). Callgrind is an
instruction-counting fallback, not a hardware latency or cache measurement.
Matched whole-child instruction references fall 1,826,378,401 → 1,303,615,049
(−28.62%). The direct owner scan falls 557,394,708 → 34,807,680 (−93.76%);
`parse_content` remains roughly 850 million inclusive references and is the
next measured hot path. Profiles include fixture setup, an untimed checksum
open, one measured open, output hashing, and report serialization.

Within the `measure_once` call subtree, the before owner scan represents
34.51% of instruction references. A hypothetical complete removal gives an
Amdahl-style instruction-work bound of 1.527× for that subtree. This is not a
wall-clock or parallel-scaling prediction. The candidate retains one scan per
reached style, and subtree instruction references fall 32.40%.

Separate matched heaptrack children report 509,402 → 507,836 allocation calls
and the same rounded peak heap display, 6.98M. The raw compressed traces are
retained. These children include setup, two opens, hashing, and instrumentation;
no operation-local allocation count or peak-live-byte claim follows. The cache
adds one cloned direct value per distinct reached style and map-node overhead,
bounded by the existing parsed-style limit and released at operation end.
Single-use/unique-style-heavy memory tradeoffs are not separately benchmarked.
There is no new shared cache, lock, executor, or worker behavior.

## Correctness and architecture

| Constraint | Implementation and evidence |
| --- | --- |
| ADR 0001/0002/0023/0024: semantic facade and ownership | Private ODG resolver only; no public API, dependency, or archive ownership change. |
| ADR 0003: immutable source and atomic publication | Cache is invocation-local; existing exact no-op, source-bound edit, inverse and publication suites remain applicable and pass. |
| ADR 0005: finite resources and evidence | Only reached definitions are cached under existing style limits; before/after timings and whole-child memory/profile boundaries are explicit. No ambient capability or parallel execution is added. |
| ADR 0006: preservation and deterministic refusal | Source XML remains authoritative. Successful direct values alone are reused; every page retains inheritance depth/cycle checks and source shadowing. No error is cached or malformed owner admitted. |
| ADR 0008: verification | Scoped ODG tests, malformed/limit cases, retained LibreOffice fixtures, downstream umbrella check, warnings-denied lint/docs, formatting, and boundary checks are recorded with their actual scope. |

The new tests cover shared direct values and exact no-op bytes, inherited alias
namespaces and sound fields, content/named-style shadowing, independent sources
with the same style name, malformed referenced owners, cycles, and a deeper
second page after a valid first page populated the direct-value cache.
See `gates.json` for final commands and results. This is scoped crate validation,
not a fresh all-workspace, native-application, sanitizer, or fuzz campaign.

The earlier 0502 comparison used an older parser with less exposed metadata.
This matched richer-parser improvement does not prove that every 0502
regression is closed. General attribute scanning, unique-style owner scans,
metadata parsing, broader CRUD/provider coverage and the full non-iWork
`docs/GOAL.md` objective remain open.
