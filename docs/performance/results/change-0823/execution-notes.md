# Execution and repairs

Production stayed at `b76786208d` during baseline quality and probe preparation.
All six baseline production gates passed. The baseline test log contains 85
successful summaries: 1,241 passed, zero failed, three ignored. Boundary auditing
was allowed to run to its terminal exit; no timeout was treated as completion.

Three pre-freeze probe-quality attempts failed and remain intact:

1. `quality-probes-0/03.log`: default-feature synthetic Clippy rejected unused
   allocator-only helpers. `probe-repair-0/` retains both complete probe trees
   before the repair. Callback APIs were feature-gated; test-only counter and
   split-region helpers were gated to allocator tests. No counter arithmetic,
   corpus generator or timed workflow changed.
2. `quality-probes-1/03.log`: default-feature test Clippy exposed the split-only
   `segment_before` field. `probe-repair-1/` retains the source revision. Its
   existing narrow unused-field annotation was aligned with the test/feature
   condition. A link to a feature-conditional helper became plain code text so
   native rustdoc can resolve its documentation.
3. `quality-probes-2/15.log`: the real probe's synthetic counter tests received
   actual test-harness allocator callbacks. The first failure compared a
   synthetic expected peak with a larger process peak; six subsequent failures
   were consequences of the poisoned test mutex. `probe-repair-2/` retains the
   source revision. The real probe now follows the existing synthetic probe:
   unit tests call the wrapper directly, while actual observer executables
   install it as the global allocator. Direct counter tests share the observer
   test lock, and the real edit parity test takes that lock in allocator test
   builds. Release allocator installation and callbacks are unchanged.

Every attempt has its original complete console log, per-command logs,
timestamps, exit codes and source hashes. No measurement was run from a failed
probe-quality revision. The successful probe-quality receipt, freeze and build
receipts identify the final sources; failed sources were not overwritten.

The baseline production quality census included extra, untested probe source
hashes before probe formatting and repair. Those extra entries are not a claim
that the production test command tested the probes. The dedicated successful
probe-quality census binds their final revision. The production census omitted
`.cargo/config.toml`; freeze separately verifies its exact base bytes. That
file defines a `cargo lint` alias unused by the recorded commands.

Candidate and probe formatting happened before freeze. `format-*.log` and
format receipts retain these preparation commands. The candidate's old scanner
oracle was independently compared with the baseline loop, allowing only the
test method and position-helper renaming. The final patch is regenerated from
the formatted before and after source files.

## Capture interruption, analysis, and disposition

The root applied the frozen candidate only after 19 baseline qualifications. Both
production six-gate suites and all four builds per leg passed. The first native
lane stopped when the offline analysis agent unexpectedly committed two reader
files, changing HEAD from the frozen base. Root retained nine report artifacts,
all logs, the commit object text, and its patch in `native-interrupted-0`; no
timing analysis selected a subset. Root mixed-reset only that owned local commit
and restarted the complete frozen native matrix. The admitted native and
allocation lanes subsequently completed with 228/76 reports respectively.

Analysis attempt 0 failed on `production_changed` versus the actual
`production_changed_at_freeze` metadata key. The log and all four reader sources
are in `reader-attempt-0` and `reader-failures.json`; hash-only lock descriptors
were also handled according to their actual frozen schema. Attempt 1, the
independent raw audit, and pre-cleanup validation passed. A later ad-hoc report
inspection requested the nonexistent `spread` key and returned KeyError in tool
output; it changed no file or analysis result. Its tool transcript is retained in
`report-inspection-failure.json`, not represented as an original process log.

The candidate misses the 3% benefit threshold; root restored the exact archived
before source. Allocation metrics match all rows, and no latency veto or RSS
review trigger fires. Cleanup removes 9,687 files/3,238,577,403 logical bytes after
verifying all eight executable identities. Baseline and candidate production
test totals are 1,241 and 1,249 passed respectively, zero failed, three ignored.
