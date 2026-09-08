# Evidence and schema review for change 0480

This is a bounded read-only review of the retained 0480 evidence boundary. It
covers the current `analyze.py`, `verify.py`, `report_schema.py`, gate receipts,
capture protocol, source manifests, and profile CSVs. The production source
review is in `source-review.md`; this document reviews custody and comparison
evidence rather than the Rust implementation.

The two-line source handoff is now independently reconstructible. `verify.py`
hashes `before-source.txt`, requires the one expected old block, computes the
candidate file by replacing that block, and checks the declared candidate
digest. The retained patch has exactly two deletions and two additions with the
expected `Arc::clone` and shared OPC method. The two flat source manifests each
contain 7,048 entries, bind the production file to the before/after digests,
and the manifest difference is exactly the production path. The candidate
binary records and final validation receipts use the candidate manifest; the
control builds and pilots use the control manifest. This closes the earlier
source-custody gap.

`analyze.expected_captures()` and `verify._check_argv()` establish all 48
captures in A1/B1/B2/A2 order. They bind `/usr/bin/time -v -o`, CPU 2 via
`taskset`, arm/instrumentation binary, mode, count, sample and warmup counts,
and label-specific resource/report paths. The current required gate set uses
`format-docx` and `clippy-all-features`; the older workspace `format` and
default-feature `clippy` failures remain retained development attempts and are
not required final gates. The verifier now parses the literal `run-gates.py`
table, checks required receipt argv against it, checks build/pilot argv against
their frozen bindings, and binds terminal receipts to their `.started.json`
fields. It also checks the recorded environment on every retained validation
receipt against the protocol. The ledger then binds each attempt's path,
digest, argv, exit status, and source-unchanged flag to the receipt.

The report schema is reused through `analyze.py` aliases and `_sync_root()`;
the analyzer does not maintain a second report validator. The report and
capture checks enforce fixed corpus/XML/member oracles, source-read histogram
accounting, sink size/digest limits, allocator live-byte conservation, phase
boundary continuity and final release, and no summation of phase peaks into a
total peak. The profile verifier parses the retained CSV rows and requires the
raw counter values and running percentages to equal `profiles.json`; the two
profile runs are explicitly excluded from the 1,440 formal samples.

## Findings requiring resolution before the evidence is sealed

1. **The baseline freeze is not ordered against the runs.** The current
   `analyze._baseline_corpus()` checks the baseline file's schema, timestamp,
   independent corpus derivation, and the `initial-state.json` digest, but it
   does not compare the historic `frozen_utc` with the arm-specific pilot
   receipt finish times and first formal capture starts. The old ordering code
   in `report_schema._corpus_manifest()` is not on this path: it refers to the
   legacy `corpus-manifest.json` name, while the 0480 analyzer reads
   `baseline-corpus.json`. A stale baseline can therefore pass the current
   analyzer. The retained facts are ordered correctly: the byte-exact historic
   freeze is `2026-09-08T17:19:20.976467+00:00`, control pilots finish before
   the first control formal capture, and candidate pilots finish before the
   first candidate formal capture. Add assertions for
   `historic_freeze < first_formal_start` and, for each arm,
   `all_arm_pilots_finished < that_arm_first_formal_start`. Do not regenerate
   the historic baseline timestamp or digest; rerun the analyzer and gates
   after adding only these ordering checks.

2. **The global peak label is semantically wrong.** The frozen schema's
   `_metric_rows()` sets `total_peak_live_bytes` from
   `region_peak_live_bytes`, then labels it “absolute process allocator
   live-byte high-water mark.” The global allocator boundary is
   `peak_live_bytes_after`; the region value is an operation-region high-water
   mark. For example, the retained control 131,072 allocator row has a region
   peak near 42.5 MiB and a global peak near 118.1 MiB. The separately derived
   `incremental_peak_live_bytes = region_peak_live_bytes - live_bytes_before`
   is the correct scoped comparison, and the phase `retained_delta` values are
   not summed, but the generated summary still carries the misleading global
   label. Rename/remove that field or derive it from `peak_live_bytes_after`,
   and add a regression fixture where the region and global peaks differ.

3. **Optional profile receipt custody is incomplete.** `check_profiles()`
   authenticates the profile receipt file's digest and validates the CSV
   values, event set, and report, but it does not open the receipt and require
   its exit status/source-unchanged result or bind its argv to the expected
   `taskset`/`perf stat` command, binary, CSV path, and report path.
   `check_validation_receipts()` also treats `profile-control` and
   `profile-candidate` as optional attempts, so it does not enforce those
   fields for them. The current receipts happen to be successful and use the
   expected commands, but a mutated receipt could leave apparently valid CSV
   values without proving how they were collected. Either make the profile
   check enforce the receipt status and exact command binding or explicitly
   exclude profiles from the retained evidence rather than presenting their
   values as verified observations.

The source manifests, 48-capture order, formal report identity, PMU raw-value
comparison, required gate names, final source arm, and required gate
argv/environment custody otherwise match the stated protocol. Existing
`test_evidence.py` tests cover XML/output mutations, read histogram and sink
limits, allocator phase continuity/release, normal-allocation absence, sample
indices, protocol order, path traversal, and GNU RSS parsing. They do not
exercise the three findings above; the baseline-order, peak-scope, and profile
receipt mutations are the meaningful missing regression cases.

## Root resolution after review

All three findings are resolved in the final helpers. The baseline remains
byte-exact 0479 evidence; the active verifier checks its freeze precedes the
formal runs and each arm's pilots precede that arm's first formal capture.
The analyzer overrides the inherited peak label with the explicit
operation-region scope, and a new fixture distinguishes region and process
peaks. Profile verification checks successful unchanged-source receipts and
the exact `taskset`/`perf stat` command and provenance paths. Pilot and profile
path checks also work after copying the bundle to a different directory.

Root's final evidence-test and analysis gates pass against the frozen
candidate source. All 14 evidence tests pass; the final analysis validates
48 captures and 1,440 samples. The data-only verifier passes all current
pilots, captures and both profile records. Earlier successful development
test/analysis receipts and `summary-initial.json` remain retained.
