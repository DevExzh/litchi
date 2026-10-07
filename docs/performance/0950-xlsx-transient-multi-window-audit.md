# 0950 — transient multi-window XML audit with reusable parser buffers

Status: adopted under the frozen gates after independent review. All 12 applied
source files match the tested and measured candidate. Archive extraction and relocated replay pass. The known unrelated workspace formatting drift remains recorded below.

## Mechanism

[0949](0949-xlsx-multi-window-allocation-attribution.md) showed that the rejected
0948 candidate retained large window records and allocated three parser buffers
for every replacement window. This candidate removes the retained window and
ancestor vectors. An observer keeps one pending source element, proves its exact
start/end boundaries, and audits the corresponding replacement slice as that
source element closes. It accumulates checked provisional counters.

The complete original audit remains authoritative. Provisional replacement work
can happen during that scan, but cannot return a replacement verdict or suppress
an original error. Only after the original succeeds can the aggregate proof
succeed. Any ambiguous span, failed window check, allocation failure, or aggregate
limit failure uses the canonical complete replacement audit. Failed scratch is
dropped before that fallback begins.

One quick-xml reader and one auditor state retain their scratch capacities across
windows. Each window must begin with markup, end with markup, and be balanced at
its inherited depth. At such a boundary the input slice is replaced without
asking the reader for EOF; cumulative parser positions are checked and rebased
to the current slice. A synthetic EOF charge preserves the prior per-window
event-limit check. End-name checking remains enabled. The auditor resets every
state field using an exhaustive struct literal, retaining only empty scratch
vector capacities and the next inherited namespace snapshot.

Normal scans use the same token checks through a constant-generic non-segment
path. The existing zero/one-span route is retained. Exact gap matching, source
identity guards, the 4,096-hint limit, half-replacement byte limit, and cumulative
8 MiB conservative namespace admission remain. That admission is a logical
charge, not an RSS bound. Namespace snapshots are now transient; this does not
claim that nonempty namespace clones have no allocation cost.

## Correctness evidence

The final XML implementation is hash-bound to the reviewed source. The debug XML
suite passes 161 tests with one existing asset test ignored; focused release
suites pass 11 new multi-window tests and 31 existing source-replacement tests.
The new tests compare outcomes with complete original/replacement audits and
cover namespace changes, varied window depth, reused reader boundaries, budgets,
malformed input, original-error precedence, and generated disjoint windows.
Default and all-feature XML Clippy for library/tests pass with warnings denied.
Both OPC debug configurations pass (632 default and 652 all-feature unit tests,
plus their integration targets). Both XLSX debug configurations pass (1,397
default and 1,404 all-feature unit tests, plus their integration/example targets).
XML doctests and warning-denied rustdoc pass in both feature configurations.
All 30 broader quality checks passed with an unchanged source tree. Counting
unit, integration, and example suites together, OPC passed 1,019/1,041 tests
(default/all features), with one ignored in each configuration; XLSX passed
2,116/2,136. Selected release integration targets passed 61 OPC and 143 XLSX
tests. OPC/XLSX Clippy and rustdoc passed in both feature configurations with
warnings denied. The boundary check passed for 65 workspace packages and 244
internal dependency declarations, retaining 11 explicit debt items.

Review caught and fixed an empty-token boundary omission: an untrusted range
starting at an empty element must also end exactly at that element. The new
regression supplies two hints and attempts to remove an ancestor closing tag;
the final implementation returns the canonical replacement error. Early compile
failures and an incorrect test expectation for legal interior U+FEFF text are
retained as preparation history. No benchmark ran against those source versions.
The known unrelated workspace formatting drift is reported separately.

## Prospective comparison

The 0948 matrix and gate constants are retained before new data: two primary
one-percent XLSX edits (medium and dense-sparse), ordinary one-edit medium, and
managed one-edit medium. Baseline runtime sources are byte-equivalent to current
HEAD df6ee5583b and the retained b6c5c80 source closure. The candidate manifest
contains 10,790 files, including the new test, versus 10,789 baseline files.

The planned schedule has 16 qualification, 48 native, and 24 allocator children,
with serial workloads on CPU 12 and a fixed C/UTC environment. Native comparisons
use six alternating paired blocks, 10 warmups and 100 samples. Allocation
comparisons use three alternating repeats, two warmups and five samples. No
failed child is retried. All ratios mean candidate divided by baseline.

Both primary publication-p50 paired-median ratios must be at most 0.95 and their
95% bootstrap upper bounds below 1.0. Every case/block must meet the 1.05 safety
limit for publication and workflow p50/p95. Allocation requests, requested bytes,
region peak/live bytes, and maximum RSS must meet the 1.05 per-pair and paired
limits. Missing metrics fail closed; zero baselines allow only equal-zero results.
Bootstrap uses 10,000 resamples and the retained seed 948009. Native timings and
allocation-instrumented timings are not mixed.

## Measurement verifier corrections

All 88 workload children completed without retries. Qualification reduction
first stopped because the reducer expected a top-level semantic-validator
descriptor while the frozen plan bound that descriptor under `adapter_helpers`.
A versioned reducer correction uses the already-bound descriptor; the plan and
validator identity were retained.

Full reduction then exposed the same numeric-type bug in the outer and nested
validators: Python returns an integer when an integer sample vector has an
exactly integral mean, causing the validators to select exact equality instead
of their existing floating-point comparison. One harness report contained
31,047,832.000000004 ns versus the recomputed 31,047,832 ns. The harness computes
means through floating-point accumulation. The corrections apply the existing
relative tolerance 1e-12 and absolute tolerance 1e-9 to means regardless of
Python's return type, while retaining exact integer percentile comparisons.
Negative controls reject an altered percentile and a materially wrong mean.
Original helpers, failed invocations, reports, and correction receipts remain in
the evidence packet. No raw data, acceptance threshold, or workload was changed.

## Results and decision

All 88 children passed capture and semantic/custody reduction: 16 qualification,
48 native, and 24 allocator. All frozen latency and allocation safety gates pass.
Independent review accepted adoption with the tail flag disclosed; the candidate
is applied. Archive verification is recorded below.

| Case | Baseline publication p50 (ms) | Candidate publication p50 (ms) | Paired publication ratio [95% CI] | Paired workflow p50 change |
| --- | ---: | ---: | ---: | ---: |
| One-percent edit, medium | 6.5124 | 5.5329 | 0.849059 [0.847750, 0.850429] | -6.20% |
| One-percent edit, dense-sparse | 13.2138 | 11.9327 | 0.885302 [0.842911, 0.905068] | -4.88% |
| One-edit control, medium | 1.6370 | 1.6293 | 0.995713 [0.991027, 0.999856] | -0.13% |
| Managed one-edit control, medium | 1.6438 | 1.6274 | 0.992978 [0.983994, 0.998248] | -0.12% |

The two millisecond columns are medians of six block p50 values. Ratios and
percent changes are medians of paired candidate/baseline ratios, so they need
not equal the ratio of those two columns. Confidence intervals use the frozen
10,000-resample bootstrap. Timings include the defined warmup/sample protocol;
reopen verification is outside the timed workflow.

| Case | Publication allocation calls, before → after | Requested bytes, before → after | Region peak live bytes, before → after | Largest RSS ratio across three pairs |
| --- | ---: | ---: | ---: | ---: |
| One-percent edit, medium | 406 → 390 | 2,495,613 → 2,495,373 | 26,760,213 → 26,764,060 | 1.028106 |
| One-percent edit, dense-sparse | 406 → 392 | 2,765,840 → 2,765,392 | 35,247,619 → 35,254,138 | 0.993807 |
| One-edit control, medium | 304 → 304 | 1,002,065 → 1,002,209 | 23,300,695 → 23,300,926 | 1.019666 |
| Managed one-edit control, medium | 306 → 306 | 1,002,145 → 1,002,289 | 23,301,311 → 23,301,542 | 1.023034 |

Allocation counts/bytes and region peaks shown above are identical across the
three paired repeats. Requested bytes fall slightly for both primary cases;
region peak live bytes rise by 3,847/6,519 bytes, and post-region live bytes rise
by 3,655/6,327 bytes. Both one-edit controls add 144 requested bytes without
adding allocation calls. These small increases remain within every frozen
memory threshold. RSS is measured separately per child; no RSS reduction is
claimed from these three repeats.

One native review flag remains explicit: ordinary one-edit medium, block 2,
workflow p99 rises from 4,057,411 to 4,286,762 ns (+5.65%). The same pair's
workflow p50/p95 ratios are 1.043673/1.046660, within the 1.05 hard gates;
publication p50/p95 ratios are 0.989221/0.990744. The p99 threshold is a
prospective review trigger. It is not evidence of a proven noise source or a
reason to remove the pair.


## Architectural constraints and scope

| Constraint | Candidate behavior and evidence |
| --- | --- |
| ADR 0003 — snapshots, edits, patches | Existing source identity, exact no-op, publication freshness, and reversible patch paths remain. Integration checks exercise stale and repeated edits. |
| ADR 0005 — I/O, memory, performance | Hint count, replacement fraction, and cumulative namespace admission remain bounded; retained window records are removed. All prospective native and allocation safety gates pass. |
| ADR 0006 — validation, security, compatibility | Original errors take precedence over provisional replacement results. Exact boundaries and outside gaps, inherited namespaces/depth, aggregate limits, and canonical fallback are checked. Differential and adversarial XML tests pass. |
| ADR 0011 — physical OOXML package ownership | OPC owns replacement validation; XLSX provides untrusted range hints. |
| ADR 0030 / 0031 — lazy OPC access and execution budgets | Existing fallible payload access and fallback remain. No ambient I/O or parallel execution is introduced. |
| ADR 0032 — small derived-value memos | Hints remain specific to the edit operation and never authorize skipping the original audit. No new memo policy is introduced. |

The change targets publication validation within four named in-memory XLSX edit
and save cases. It does not establish cold-cache, remote-source, concurrent,
cross-platform, or scaling improvements. Allocation-instrumented timings do not
establish native speedups. Namespace cloning can still allocate for other inputs;
the 0949 attribution applies only to its two measured fixture shapes. Final
acceptance must use complete workflow and publication measurements together.

The p99 review decision is to adopt with this limitation recorded: no pair is
excluded, every hard gate passes, and the ordinary one-edit workflow-p50
bootstrap upper bound is 1.022677. The isolated p99 increase is retained as a
follow-up tail-latency observation; no cause is claimed.

## Build, host, corpus, and reproduction

The host is AMD EPYC 9R45, Linux 7.0.0-1013-aws, 132,553,801,728 bytes of
reported memory, 4,096-byte pages, CPU affinity 12. Each timed child sees one
available logical CPU. The binaries use Rust 1.95.0, release opt-level 3,
fat LTO, one codegen unit, debug level 1, incremental disabled, and the system
allocator; the separate allocator binary enables `allocator-metrics`. Build
receipts supply compiler/source provenance where runtime environment fields
are null. Two candidate builds use two Cargo jobs each and finish before
measurement; all workloads are serial and follow the frozen C/UTC environment.

The ordinary and managed controls edit one cell of the medium synthetic XLSX.
The primary cases update 93 cells in medium and 178 in dense-sparse. Deterministic
corpus/output hashes, preservation identities, exact gates, all per-pair ratios,
and full artifact descriptors are in
[the bounded summary](results/change-0950/summary.json). The full verified
analysis is 5,811,362 bytes and remains outside Git. Baseline sources are the
retained 0948 closure at `b6c5c80ffbcb60736d6c4c2f3b590838640008a9`; the
pre-change measurement checkout is `df6ee5583b409f2547e2f824896d87feb8d50630`.
All 12 applied files match the candidate manifest. Production copying therefore
does not require rebuilding identical bytes for another measurement.

From a restored evidence packet, regenerate only into fresh output paths:

```sh
python3 -B tools/reduce_0950_v5.py all --output regenerated-analysis.json
python3 -B tools/summarize_0950_v1.py --output regenerated-summary.json
```

Build commands are `python3 -B build-tools/build_candidate_native_v1.py` and
`python3 -B build-tools/build_candidate_allocator_v1.py`. Fresh capture commands
are `python3 -B tools/capture_0950_v2.py qualification`, then `native`, then
`allocator`; existing capture paths are create-only and must not be overwritten.
For a relocated packet, first verify its files against the external archive
manifest, then use the committed [relocation adapter](results/change-0950/replay.py):
`python3 -B replay.py all --packet /absolute/restored-packet --output /absolute/fresh-analysis.json`.
The strict retained 0948 build/source dependency is also required. The archive
contains exact commands, fixed environment, source snapshots, complete raw
reports and terminals, verifier corrections, quality logs, and review receipts.

The unrelated workspace formatting failures remain at
`soapberry-zip/src/office.rs`, `litchi-iwa-archive/src/package.rs`,
`litchi-iwa-archive/src/zip.rs`, and
`litchi-opc/tests/source_fingerprint_quality.rs`. Scoped candidate formatting
passes; no global format-clean claim is made and no unrelated source is changed.

## Evidence archive and cleanup

All large artifacts remain outside Git under
`/home/zhuhe/litchi-performance-0950-archive/` (archive/manifest) and
`/home/zhuhe/litchi-performance-0950/` (canonical sources/raw evidence).

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `archive-0950-v1.tar.zst` | 464,187,123 | `cee0b5d2c9e878ca577994d4a4f8e56032c1d119fe6e27ac3b3d8a8e51daa363` |
| `manifest-v1.json` | 8,774,266 | `6efa98e164432849a31dd11485160afc4f5f94976a531542d3a03776ff3fd1ab` |
| `analysis-v5.json` | 5,811,362 | `bbec30403559154534a2b8d185d4929fb32dbe10d246e4d8f97edb1128b22b13` |
| `candidate-source-manifest-v1.json` | 2,037,928 | `5b16ef65a17f868a2d941355baed30dbbf9e338b817365c02bfb7133fedc80d4` |
| `baseline-source-manifest-v1.json` | 2,033,638 | `dd66b044540d185c448cae78b16e9aa61556a10f01252e27f2075f48aa5121c2` |

The archive contains 33,312 files totaling 1,897,047,763 bytes, four allowlisted
ELF executables, and the required empty working directory. Extraction verifies
every file's bytes, hash and mode, the directory census, and absence of symlinks.
The retained 0948 archive dependency is identified in
[the archive closure](results/change-0950/archive-closure.json).

The frozen v4 replay adapter initially confused the historical packet root with
its `workload-cwd-v1` subdirectory. A derived adapter corrects only relocation
of that directory; an initial wrapper-hook omission is also recorded. The
original adapter and failed attempts remain preserved. The final adapter's
full replay matches `analysis-v5.json` exactly, including SHA-256. Negative
controls reject tampered historical cwd/argv and a missing physical artifact.
The derived adapter and [replay closure](results/change-0950/replay-closure.json)
are committed with this record because they postdate the sealed archive.
No workload was launched during replay.

After verification, cleanup removes 18,826 build/temporary files and 2,892
directories, totaling 39,880,216,865 logical bytes. Both baseline executables,
both candidate release executables, source trees, raw evidence, and receipts
remain. The temporary extracted replay tree is also removed. Logical byte
counts do not claim physical disk-block savings from reflinked files.

The archive was generated with the external helper:
`python3 -B /home/zhuhe/litchi-performance-0950-archive/archive-0950-v1.py --seal --go`.
Its manifest selects the retained evidence while excluding build intermediates;
regeneration requires a fresh archive destination. The candidate source manifest
comes from the archived freeze helper; the baseline manifest is retained from
the hash-bound 0948 dependency. Build/capture/reduction commands above regenerate
the underlying measurements from the bound sources.
