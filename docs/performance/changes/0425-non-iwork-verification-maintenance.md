# 0425: non-iWork verification maintenance

Rust 1.98.1's strict Clippy gate rejects constant-width chunk iterators across
the non-iWork codecs. This batch replaces them with borrowed fixed-array
views, retaining full-chunk iteration, checked lengths, terminators, padding,
mutable-buffer behavior and explicit remainder rejection. Runtime-sized widths
remain dynamic. It also closes mechanical assertion/setup findings, removes
impossible fixed-array conversion fallbacks, and moves an unchanged DOC test
module after production items. No ownership layout, dependency, public API,
lint allowance, parallelism or performance claim is introduced.

The [bundle](../results/change-0425/README.md) binds source hashes, exact
commands, raw output, Rust 1.98.1, four build jobs and one test thread. Cargo,
tests and verification workloads run serially. Its package inventory derives
45 leaf/shared packages from 64 workspace manifests, excluding 17 iWork owners
and the Python binding that hard-enables iWork. The root facade uses an
explicit non-iWork feature closure; standalone tools have separate checks.
Read-only reviews cover shared substrates, DOC/PPT, XLS/XLSB, crypto/fonts/
images/OLE, benchmark CFB helpers and the subsequent test corrections.

The full 45-package test run records 15,775 passes and three failures, with 98
ignored tests. All three failures reproduce at baseline `340cc91ae` with the
same dependency lock. Test-only corrections preserve ODP's pre-materialization
size-limit oracle, strengthen ODT malformed-descriptor refusal and atomicity,
and distinguish XLSB's six parsed drawing anchors from its five supported
worksheet transfers. All three complete integration targets then pass 27
tests. Source comparison proves that only those test files changed after the
full run: composite coverage is **15,778 passing tests**, not a second complete
run. Earlier repeated tests are not added to that count.

Strict Clippy passes for 44 selected packages and the three corrected package
targets. ODF's unchanged 312-byte/eight-byte enum layout warning remains open;
the 45-package command with that lint exempted is diagnostic only. Boxing the
reader would add allocation and needs a separate measured ownership decision.

The facade has 455 passes, six failures and 11 ignored tests. Its six failures
reproduce at baseline in the two affected targets (360 passes, six failures):
ODT/OOXML arbitration and input limits, deferred XLSX access, XLSB wrong-format
errors and legacy XLS formula extraction remain unresolved. Its strict gate
also retains 18 findings in unchanged code. The standalone harness retains 29
lint findings outside the modified code; native-resave stops before compilation
because its unchanged tracked lockfile needs updating. None is suppressed or
reported as a passing gate.

The eight affected XLS/CFB harness tests pass. Warning-denied rustdoc passes
for all 45 packages and the facade. Crate boundaries, the 15-category/32-selector
CRUD index, all nine registered strict claim replays and formatting for 194
changed Rust files pass. The portable verifier checks all 34 command receipts
and the composite source transition after temporary-checkout cleanup.

This is verification maintenance, with no before/after timing, allocation,
peak-memory or I/O inference. It does not recapture native Office, fuzz,
sanitizer, hardware-counter or performance evidence. Facade correctness and
strict-gate debt need follow-up alongside caller-drop snapshots, near-limit
budgets, native/cold/range coverage, streaming, bounded scaling and the broader
CRUD matrix. The full non-iWork goal remains active and incomplete.
