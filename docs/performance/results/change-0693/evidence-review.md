# 0693 independent evidence review

This is a read-only packet review. I did not run Cargo, a build, a profiler,
or a workload. I checked the retained JSON, raw TSV, trace, binary, source,
and restoration receipts, and ran only derived Python checks outside the
packet. iWork is outside this review.

## Disposition

The frozen 0693 evidence is internally consistent and has no substantive
evidence blocker. The coordinator's final `audit.py` receipt also passes after
normalizing explicit trace-run paths. The measured result remains correctly
scoped to the retained capture, phase, allocator, refusal, profile, and
diagnostic trace observations; it does not establish a general performance
claim.

## Independent checks

- `builds-candidate.json` binds all 602 production source files to the live
  candidate byte-for-byte. The candidate native, allocator, and refusal probe
  binaries, together with all retained baseline binaries, match their recorded
  SHA-256 values. The
  13 semantic case bindings, generated input receipt, real input receipts,
  and the 103-member/43-replacement mechanism control agree with the raw
  outputs and current corpus hashes.
- The native matrix has 78 successful process records and 7,800 timed rows
  (13 cases, 100 samples, five warmups per process). Every raw output and
  stderr digest matches its run receipt. All phase columns are present and
  numeric. No-op records carry `commit_is_changed=false`,
  `revision_identical=true`, and `output_identical=true`; edit and two-edit
  records carry the expected post-edit semantic/archive and reopened-target
  fields. The retained native summaries have 468 rows, 312 paired comparison
  rows, and 28 phase triggers.
- The allocator matrix has 26 successful files and 78 timed rows (13 cases
  per binary, three samples and two warmups). Its 49-column schema includes
  both allocation and reallocation counters. The retained comparison has 78
  case/phase groups. The requested-byte field already includes the new size of
  successful reallocations, so the report's subtraction does not double-count
  `realloc_requested_bytes`; net live bytes do not change and no measured
  peak increases.
- The refusal matrix has six successful legs, ten fixture blocks per leg, and
  6,000 samples. Every fixture has 100 rows. The expected and observed typed
  errors match for all malformed cases; valid cases retain their expected
  success form. The graph, metadata, and duplicate-name preservation receipts
  are bound in `refusal-bindings.json` and the 60-row summary/40-row paired
  comparison.
- The completed trace run is diagnostic-only, has six successful probe runs,
  no verification errors, and restores all three instrumented sources exactly.
  Real and control captures have 18 MCE calls; generated captures have 19,
  including the two disclosed unmapped generated-resource calls. The trace
  records 24 bytes per proof entry and 48 bytes per capture entry. The pointer
  records are explicitly process-local. The failed qualification attempt under
  `trace-initial` has no probe runs and also restored its sources exactly.
- The profile counters reproduce the stored `(210 - 10) / 200` slopes exactly;
  profile file bindings all match. The retained RSS values are 5,540 and 5,632
  KiB (+1.66%), and the native text sizes are 2,588,542 and 2,595,586 bytes
  (+7,044), matching `profile-summary.json` and `binary-sizes.json`.
- The seven integration results, six repository evidence gates, and quality
  totals are successful: 918 default tests plus 2 ignored, 932 all-feature
  tests plus 2 ignored, and 45 facade tests. No fuzz or Miri execution is
  claimed.

The `expanded-refusal-baseline` and `initial-identity-helper` directories are
historical receipts. Their restorations and raw files are self-consistent.
The expanded-baseline manifest's archived candidate bytes predate the final
candidate in `opened/tests.rs` and `notes/package.rs`; this is visible in its
recorded hashes and does not bind the final measurements. Final candidate
builds and measurements use the current 602-file candidate map.

## Seal bookkeeping

The post-fix `trace-summary.py` path normalization changed a driver file after
the first script receipt. Before sealing, refresh `script-checks.json` for the
current `trace.py`, `trace-summary.py`, and `audit.py` bytes (the raw evidence
and successful audit are unaffected), then retain the post-cleanup seal. This
is a packet receipt refresh, not a measurement or source-custody blocker.

Owner disposition: refreshed all top-level Python syntax/hash receipts before
the final seal; see `script-checks.json` and `final-validation.json`.
