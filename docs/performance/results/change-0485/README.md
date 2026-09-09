# 0485: bounded OPC splice consumption

This bundle compares the retained 0484 DOCX tail-append executables with a
private OPC adapter change. Bytes already consumed by the XML parser are
hashed and emitted together within the existing bounded window. The separate
source XML audit, replay authentication, source freshness policy, and
per-fragment authored Work charges remain in place.

The [implementation review](implementation-review.md) describes the invariants
and ten focused regression tests. The [methods](methods.md) define 18 arms
across three workloads, two executable roles, and two reversed process
repeats. Formal samples exclude pilots and external profiling. GNU time RSS
is one whole-child observation; allocator operation peak is measured in a
separate instrumented executable.

## Measured results

All 144 formal processes and 4,320 samples pass. Authored-heavy file-input
p50 falls from 443.916 / 440.383 ms to 239.499 / 238.865 ms across R1/R2,
a 45.76–46.05% reduction. Source-heavy owned p50 falls from 482.325 / 476.215
ms to 385.058 / 385.006 ms. Candidate archive bytes remain identical.

The [results and regression review](results-review.md) retains all 18 normal
arms, three adverse latency quantiles, and nine distinct adverse whole-child
RSS observations. Operation heap peaks remain effectively unchanged. The
[full tables](comparison-summary.md), [raw summary](comparison-summary.json),
and [chart](latency-comparison.svg) keep the two process repeats separate.

The [profile summary](profiles/profiles1-summary.md) validates 12 diagnostic
children. Authored-heavy file `statx` calls fall from 3,735,939 to 1,475,055,
while source-heavy file calls rise from 15,415 to 25,219. The latter is a
recorded metadata tradeoff despite lower elapsed time. Both comparisons keep
their `pread64` counts unchanged.

## Validation

All-feature OPC/DOCX tests pass with 1,989 tests and 32 ignored. The
no-default-feature configuration passes with 1,948 tests and 32 ignored.
Formatting, Clippy with warnings denied, rustdoc with warnings denied, the
five replay benchmark tests, and nine Python helper tests also pass. The
crate-boundary check passes for 64 packages and 239 internal dependencies.
The sanitizer fuzz lane passes 27 required-success cases and two completed
10,000-run campaigns with no crash artifacts.
Existing fixture tests include real LibreOffice, POI, and Microsoft Office
packages. This is not a new native Microsoft Office certification run.

Development attempts are retained. The first focused compile exposed test
sink borrowing errors; the first full compile exposed a missing test import;
the first Clippy attempt exposed two syntax lints. Later checks fixed those
issues. `tests-all-features-final3` encountered the host `/tmp` quota;
`tests-all-features-final4` passed with an explicit dedicated `TMPDIR`.
`boundaries-dev1` passed its command but observed source edits, so it does not
serve as the final boundary gate. The 21 selected final gates in `validation-summary.json` require unchanged
source hashes that match the measured after build.
`summary-profiles-final1` exposed a helper inventory error that rejected the
receipt file itself; the corrected `summary-profiles-final2` validates all
12 retained profiles.

## Custody and scope

`machine.json` records the actual host and CPU 2 setup. Build receipts bind
compiler flags, source manifests, gates, and retained executable hashes.
`comparison-protocol.json` freezes the matrix and report validators before
measurement. Captures retain command lines, raw reports, GNU time output,
artifact hashes, and empty file-store cleanup receipts. Historical 0484
inputs remain immutable.

The coordinator serializes heavy work with the existing CPU lock. CPU pinning
and this lock do not reserve the entire host; point observations in
`host-observations/` document visible compiler/profiler activity and load.
Two process repeats support descriptive comparisons, without strong
confidence intervals. Every change above the five-percent review threshold
is retained for review, including adverse tail and RSS rows.

`cleanup.json` records removal of owned fuzz scratch, bytecode, and empty test
scratch, plus all 36 pilot/formal file-store cleanup receipts. Three copied
executables remain as explicit external evidence. The final
[seal contract and commands](seal-methods.md) describe independent read-only
verification; `seal-verification.json` is the sole excluded verification
receipt. Build inputs copied for fuzz reproducibility use `.txt` suffixes.

The full non-iWork goal remains open. This bundle does not close cold-cache,
concurrent, atomic-save, all provider/input intersections, or broad native
Office validation requirements.
