# 0915 — DOCX section namespace resolver experiment

**Decision: do not adopt the candidate.** The retained eight-pair result fails
the predeclared 1.05 ratio threshold. Early children also overlap a boundary
check on other CPUs, so the capture does not establish an isolated optimization
benefit. Production code remains unchanged.

This batch investigates the source-backed existing-section layout edit/save
case baselined in [0914](0914-docx-section-layout-baseline.md). The candidate
borrows the XML reader’s namespace resolver for one event instead of cloning
its namespace buffer and binding vector. Its scope is only
`ensure_source_section_inventory_xml` in `litchi-docx`; two helper-call arguments
are adjusted to pass the borrowed resolver directly.

## Mechanism and preservation boundary

The reader is not mutated while an event is validated. Its resolver is read-only
throughout that interval, and the next XML read occurs after the borrow ends.
The candidate keeps event order, root checks, namespace and attribute validation,
DTD/entity/MCE refusals, and input/event/depth limits. It introduces no cache,
retained inventory, public API, dependency, or unsafe code.

Complete-source fingerprinting and source-version fences remain mandatory.
A generic caller-supplied `ReadAt` is not necessarily immutable; replacing
fingerprinting with version equality would weaken exact-source authorization.
Publication still performs the existing patch checks, output fingerprinting,
and preservation plan. See the [source and ADR review](results/change-0915/design-review.md).

## Profile evidence

The [Callgrind packet](results/change-0915/profile-README.md) isolates the query
snapshot and the current-snapshot subcall inside publication. Each selected
snapshot runs the section guard once. At the resolver clone site, both selected
owners show 1,838 process-event calls, 3,676 allocation calls, 3,674 deallocation
calls, and 1,838 copy calls; an independent repeat reproduces these counts.

The guard accounts for about 1.81–1.82% of selected-snapshot guest instructions.
These are Valgrind simulation counts. The host denies hardware perf access,
and guest software SHA execution differs from native execution, so those
instruction shares do not establish native cost or predict end-to-end benefit.
The publication selection excludes patch application, writing, and output
hashing. The packet retains exact caller contexts and external artifact hashes;
`profile-replay.py` verifies those artifacts and recomputes the selected counts.

## Native experiment

The [protocol](results/change-0915/native-protocol.json) was frozen before native
capture. Eight fresh matched process pairs alternate baseline/candidate order
on CPU 12; each child executes 30 warmups and 200 measured iterations. The
baseline executable is the retained 0914 release binary, and both sides use
the same generated corpus, operation, phase clocks, and correctness controls.
The experiment does not compare against historical 0914 timings.

Admission requires the median of the eight paired baseline/candidate child-p50
ratios to be at least 1.05, and the lower bound of a seeded 10,000-draw paired
bootstrap interval to exceed 1. Pooled p50/p95/p99 and setup-inclusive maximum
RSS increases above 5% are review flags. Pooled samples are descriptive;
process pairs are the resampling unit.

The total measured interval is the sum of six phase clocks and excludes
inter-phase gaps, verification, and teardown. Each operation uses a fresh
in-memory source within a warmed process. Logical read counters do not measure
physical I/O, and whole-process maximum RSS does not measure operation peak
live memory. Native allocation counts and hardware-cycle attribution are not
available from this capture.

## Results and disposition

All 16 child reports pass the 0914 semantic, byte-preservation, reversible patch,
typed-refusal, source/output identity, phase-sum, cache, logical-source, and sink
gates. There are 1,600 measured samples per side. The
[retained reduction](results/change-0915/reduced.json) reports:

| Metric | Baseline | Candidate |
| --- | ---: | ---: |
| Pooled total p50 | 4.813090 ms | 4.747636 ms |
| Pooled total p95 | 4.851850 ms | 4.782320 ms |
| Pooled total p99 | 4.876500 ms | 4.810723 ms |
| Median whole-child maximum RSS | 37,676 KiB | 38,378 KiB |
| Measured samples | 1,600 | 1,600 |

The median paired ratio is **1.013824842**, with bootstrap interval
**[1.013222680, 1.014816515]**. The lower bound exceeds one, but the ratio
misses the required **1.05**. Every individual pair ratio is below 1.016.
None of the predeclared pooled-tail or RSS review flags fires. RSS is 1.86%
higher by the median process statistic; it is not an allocation or peak-live
measurement. These are observed distributions, not an admitted speedup claim.

The [capture limitation](results/change-0915/capture-limitation.json) records
an unintended overlap: the restarted boundary checker on CPUs 14–15 remained
active during the first seven children pinned to CPU 12. The approximate
checker window is 15:41:01.977–15:41:41.871 UTC, from log timestamps and the
observed interruption; there was no contemporaneous wrapper timing receipt.
This weakens isolation even though it used different cores. All samples remain
in the frozen reduction; no rows were removed, no favorable replacement cohort
was collected, and no alternate admission rule was applied. The candidate is
not adopted and the exact patch remains solely as an experiment artifact.

## Reproduction and verification

The [manifest](results/change-0915/manifest.json) binds the protocol, capture
script, 0914 validator, source/build receipts, and both executables before the
first child. Committed deterministic gzip reports contain the exact raw JSON
bytes, with separate raw and compressed hashes. Process logs and receipts are
also committed. The [independent review](results/change-0915/review.md)
recomputes statistics and rejects corrupted copies of reports and summaries.
Replay the retained native packet without the external binaries:

```sh
python3 -B docs/performance/results/change-0915/native-compare.py replay
```

The baseline source archive is the 0914 export at
`a4ca85e28cdea151176534b7f8dc912253a1ca1e`; the candidate is that archive plus
[`candidate.patch`](results/change-0915/candidate.patch). The independent
[source comparison](results/change-0915/candidate-source-verification.json)
checks all 9,357 archive files byte-for-byte and mode-for-mode: exactly the
intended source file differs, with no mode differences. The patch has three
hunks and changes three lines. Its SHA-256 is
`1bebd5ee09bfbce6f7d79ea33caf5ad35ef24ef033365a2e5fe74b611033ea4a`.

The candidate executable is 69,938,512 bytes, SHA-256
`b3b3279d839b5dd92fed0533a528934b3c901a356f79ee694db3826079f0c053`.
It remains external under `/home/zhuhe/litchi-performance-0915-candidate/`.
The source archive is 171,120,640 bytes, SHA-256
`7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`.
The [source receipt](results/change-0915/receipts/source-receipt.json) and
[build receipt](results/change-0915/receipts/build-receipt.json) retain exact
paths and identities. The build receipt candidly records its end time from
log mtime; its wrapper did not record a start timestamp. Both receipt byte
identities were frozen before capture.

To rebuild, recreate the 0914 pinned source export, apply the retained patch
to `crates/litchi-docx/src/source_backed.rs` (the patch headers contain original
absolute paths), and use an external target directory:

```sh
CARGO_TARGET_DIR=/new/external/target cargo build --release --locked -j2 \
  --bin litchi-perf-baseline --manifest-path /new/source/tools/perf-baseline/Cargo.toml
```

For a new capture, use a separate packet directory/check-out and external root;
the helper deliberately refuses to overwrite frozen files. Set
`LITCHI_PERF_0915_ROOT` to a new external directory, then run `freeze` with
`--candidate`, `--baseline`, `--source-receipt`, and `--build-receipt`, followed
by `run`, `reduce`, and `replay`. Keep quality work stopped throughout that new
capture. This batch does not perform that replacement experiment.

The candidate passed 1,911 DOCX tests with one ignored test, 78 doctests with
31 ignored, warning-denied Clippy, warning-denied documentation, and DOCX
formatting checks. The corrected boundary-check mirror passes for 65 packages
and 244 internal dependency declarations, with 11 recorded debt items. Test setup initially lacked the workspace lockfile and
compile/runtime fixtures. The corrected test export uses all 1,426 tracked
`test-data` files plus two separately identified external SVG fixtures. Failed
setup attempts and mistaken broad-scope/tool invocations are retained in the
[quality record](results/change-0915/quality-receipt.json); no full-workspace
quality claim is made.

## Scope

This is one generated DOCX fixture on one shared host. It does not establish
real-producer diversity, filesystem/cold/range performance, parallel scaling,
or completion of the broader non-iWork goal. iWork remains outside this batch.
