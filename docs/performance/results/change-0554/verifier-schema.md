# 0554 verifier contract

`verify.py` is a read-only, fail-closed checker for the OLE2 directory-name
handoff campaign.  Its CLI envelope is:

```json
{
  "schema": "litchi.ole2.verification.0554.v1",
  "status": "pass | incomplete | fail",
  "scope": "selected component",
  "result": "present only for pass"
}
```

`--strict` returns 2 for missing later evidence and 1 for malformed or
contradictory evidence.  The verifier never runs Cargo, Rust, a capture
driver or cleanup. It replays the bound pure metrics, profile and instruction
analysis functions in isolated Python interpreters, comparing their full JSON
results exactly; it never invokes their writing CLI entry points.  It also never writes an evidence artifact.

The frozen input map binds `plan.json`, `run.py`, the ignored workspace lock
binding, and the ADR manifest.  The plan binds the Git revision, source file,
candidate origin, CFB/XLS matrix, ABBA order, owner profile jobs, and all
admission rules.  Baseline source must equal the frozen Git tree and have an
empty patch.  Candidate source must have an exact changed-path patch and a
retained byte witness for every changed path; `candidate-preparation/file.rs`
is the witness for the planned CFB change.  A final source is accepted only if
it is the exact bound candidate on adoption or the exact baseline on reject,
and the live source has the same manifest.

Each build and capture receipt is checked for its actual command, source and
execution-stage hashes, binary hash, script and plan hashes, artifact map,
UTC interval, and non-overlap with the other receipts.  The stage inventory is
an exact set of filenames.  The four retained binary descriptors are checked
against their build receipts, bytes, source manifest, and retained binary.
After cleanup, descriptors may refer to missing binaries only when the
checked cleanup record has an empty process-reference list, binds the passing
precleanup verifier output by hash, and carries exactly the descriptor and
binary hashes observed before deletion.

The quality validator requires `quality.json` to be byte-identical to one
passing `quality-attempts/*/result.json` for `final`/`commands`.  It checks
the quality input binding, every command receipt, stdout/stderr hashes, and
serial intervals.  The metrics report is `metrics-analysis.json` with schema
`ole2_name_handoff_metrics_0554_v1`; its `source_manifests` must equal the
retained stage manifests and its `main_gates.all_frozen_main_gates_pass` value
is the numerical gate.  Profile evidence is mandatory in both stages and both
repeats.  The stage reports use
`ole2_name_handoff_0554_profile_analysis_v1`; the canonical
`profile-comparison.json` uses
`ole2_name_handoff_0554_profile_comparison_v1` and retains a boolean
`comparison.mechanism_gate` for every mechanism rule, including the
many-small owner-Ir rule.  A false gate is valid evidence for a rejected
candidate; it is never converted into a skipped profile.

The review document has schema `ole2_0554_adverse_review_v1`, status `pass`,
boolean `adoption_allowed`, and mandatory metric, profile, instruction and final
quality hashes. It retains one-for-one canonical copies and individual
interpretations for all 83 adverse comparisons and 43 repeat-drift rows.  The decision document has exactly these fields:

```json
{
  "schema": "ole2_0554_decision_v1",
  "status": "pass",
  "observed_utc": "timezone-aware ISO-8601",
  "scope": "0554 full disposition",
  "disposition": "accepted | rejected",
  "adoption_allowed": true,
  "plan_sha256": "...",
  "metrics_sha256": "...",
  "profile_sha256": "...",
  "quality_sha256": "...",
  "review_sha256": "...",
  "source_manifest_sha256": "..."
}
```

The verifier derives `adoption_allowed` as the conjunction of the metrics
gate, every profile and instruction mechanism gate, and the review's `adoption_allowed` value.
The decision disposition must match that result and its source hash must be
the exact final source manifest hash.  A failed mandatory gate therefore
requires a rejected decision and restored baseline.

The cleanup document has schema `ole2_0554_cleanup_v1` and the same exact
field inventory used by the verifier: `observed_utc`, `plan_sha256`,
`target`, `removed`, `owned_paths_absent`, `accessible_process_references`,
`process_reference_scope`, `python_cache_absent`,
`binary_sha256_by_kind`, `binary_descriptor_sha256_by_kind`,
`precleanup_verification_sha256`, and `scope`.  Its four binary maps use the
keys `baseline/normal`, `baseline/alloc`, `candidate/normal`, and
`candidate/alloc`.  `precleanup-verification.json` is the exact pass envelope
printed by `--component precleanup`; cleanup binds its byte hash before target
deletion.  The documentation manifest uses schema
`ole2_0554_documentation_manifest_v1`, sorted external `docs/` paths, and
must include `docs/performance/0554-ole2-name-handoff.md`.

Terminal sealing validates the exact post-cleanup target state and a sorted
recursive `SHA256SUMS` inventory that excludes only the seal itself.  Symlinks,
Python bytecode caches, duplicate inventory paths, and non-file evidence fail
closed.

The frozen profile consumer has one narrowly admitted correction: baseline
repeat 2 executes under the candidate checkout while retaining its baseline
binary and source identity. The original script and failed receipt are retained
and hash-bound by `analysis-amendment.json`; the verifier accepts only the exact
recorded old/new digests. No measurement or admission threshold changes.
Instruction consumers are honestly bound after capture in `instruction-inputs.json`.
Their stage reports must equal the comparison's embedded reports, and all six
instruction mechanism gates participate in the decision.

After strict cleanup custody succeeds, metrics replay permits only the four
`binaries/{stage}/{kind}/present` observations to change from true to false.
All hashes, byte counts, source bindings, measurements and gates must remain
exactly equal to the precleanup canonical report. The precleanup verifier
version is retained under `verifier-attempts/before-postcleanup-presence`.
