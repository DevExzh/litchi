# Independent review of the 0476 evidence boundary

Reviewed 2026-09-08 against the frozen `protocol.json`, the schemas emitted by
`common.py`, `prepare.py`, `build.py`, `capture.py`, and `gate.py`, and the
current `analyze.py`, `verify.py`, and `report_checks.py`. This review does not
assess the Rust implementation or rerun the capture workload.

The final verifier now has a separate provenance boundary in `custody.py`.
`verify.verify_bundle()` calls `custody.verify(root)` before it reads the
protocol rows for report analysis. The custody pass on the current bundle
authenticates 24 main lanes, 4 pilots, 4 counter lanes, 4 guard lanes, and 23
validation receipts. The remaining final action is the root-owned sealed
replay and acceptance receipt; the findings below are no longer open code
omissions.

## Root final acceptance

Final live verification passes all 36 lanes, 28 validation receipts, both
binaries per arm and the 7,034/7,035 source manifests. All eleven Python tests
and exact summary replay pass. After owned runtime cleanup, a fresh sealed
copy passes standalone verification and restored replay. Five resealed probes
reject a modified derived summary, omitted receipt command, omitted producer
output, omitted source manifest and duplicated guard shape. Initial verifier
and fixture failures are retained. The shared guard's one paired XLSX p99 flag
and two candidate tail-drift flags remain explicit in `guard-review.md`.

## Finding dispositions

1. **Auxiliary-suite coverage — resolved.** `custody._make_lanes()` and
   `_verify_protocol()` require the exact frozen `order`, `pilot_order`,
   `counter_order`, and `guard_order` arrays, including lane, arm, repeat,
   mode, and shape metadata. The final entrypoint therefore cannot fall back to
   the formal 24 rows when an auxiliary suite is missing.

2. **Capture custody — resolved.** `_verify_capture()` requires the exact
   started/finished receipt field sets, shared-field equality, lane identity,
   command argv, environment, source and binary bindings, monotonic timestamps,
   clean status, exit code zero, and the complete artifact map. It authenticates
   report, catalog, resource, stdout, stderr, and counter artifacts against the
   retained files.

3. **Source, prepare, build, and binary identity — resolved.** Custody checks
   both frozen source manifests and fixtures, prepare receipts, candidate build
   receipts, the reused control build and seal selection, mode-specific binary
   identities, and the command binding chain. Candidate capture starts are
   ordered after preparation and build completion; control build preparation is
   ordered after control preparation.

4. **Producer report identity — resolved in the current report layer.**
   `report_checks.validate_report()` and the pilot/guard validators now require
   the producer schema, tool and binary identity, environment, configuration,
   corpus dimensions, source/PPTX oracle, sink counters, operation metrics,
   status/scope fields, and deterministic catalog references. Formal summary
   replay uses the validated deterministic projection. Root still needs to
   record the final full replay result.

5. **Frozen corpus oracle — resolved.** Custody compares main and pilot report
   corpus objects to the frozen tiny/medium/large protocol objects and binds
   their output archive identities to the canonical catalog. The report layer
   validates the complete corpus/source oracle and guard corpus projections.

6. **Guard shape coverage — resolved.** Guard reports require the exact
   selector × `{tiny, large}` result set once each, and the analyzer checks the
   same identities and deterministic projections across the ABBA guard lanes.

7. **RSS and operation-peak review flags — resolved in the current analyzer.**
   Paired summaries retain process RSS and allocator operation-peak deltas with
   the declared five-percent review threshold. R1/R2 repeat summaries retain
   operation-peak drift and review flags, and `review_required` includes timing,
   RSS, and operation-peak observations.

8. **Validation gate custody — resolved.** Custody requires every retained
   started/final validation pair, exact field sets, typed command argv, absolute
   cwd, environment and driver/common hashes, monotonic timestamps, unchanged
   source snapshots, and stdout/stderr artifact hashes. The Rust validation
   ledger separately binds the required successful labels and retained nonzero
   attempts, so retained rejected attempts cannot substitute for required
   passing gates.

9. **Portable path redirection — resolved for the final verifier.** Custody
   uses its frozen lane table and bundle-contained paths, rejects traversal and
   symlink components, and validates regular files. The sealed verifier also
   requires exact `SHA256SUMS` coverage. External source-tree and binary paths
   remain behind the explicit live check.

10. **Verifier independence — reclassified as nonblocking.** The provenance
    verifier is a self-contained stdlib module and does not import or execute
    capture/build drivers. The final verifier may share sealed pure report and
    arithmetic helpers with the analyzer, as established by the task scope;
    duplicating those helpers is not required. The focused custody mutations
    below exercise the independent source, command, lane, artifact, and
    canonical-hash boundary.

11. **Elapsed-statistics and unsupported-counter contracts — resolved in the
    current report layer.** `check_elapsed()` requires `unit == "ns"`, an
    explicit complete sample permutation, retained statistic keys, and the
    producer confidence-interval method and bounds. Metric vectors require
    explicit status/scope and preserve `not_supported`, `not_counted`,
    `unavailable`, `not_applicable`, and `overflow` states without coercing
    them to zero.

## Independent mutation evidence

Each mutation was applied to a temporary copy of the bundle; no retained
artifact was changed.

- Swapping the first two protocol lanes was rejected by the frozen lane-order
  check.
- Replacing the candidate source revision was rejected by source identity.
- Changing a capture argv token was rejected by started/final receipt binding.
- Changing a capture artifact byte count was rejected by artifact metadata.
- Changing catalog canonicalization while recomputing the catalog file hash
  and receipt metadata was rejected by the canonicalization check.
- Replacing a validation argv array with a scalar was rejected by typed argv
  validation.
- Replacing a capture exit code with JSON `false` was rejected as a non-integer
  status; replacing protocol `workers` with JSON `true` was rejected likewise.

`custody.py --root docs/performance/results/change-0476` passes on the complete
bundle. `custody.py` compiles, and the lightweight evidence suite passes all 11
tests. No Rust build, capture, profiler, or other heavy workload was run for
this review.

The root-owned final step is to run the sealed portable replay, optional live
source/binary check, and final mutation/seal receipt after all documentation
and acceptance artifacts are frozen.
