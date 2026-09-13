# 0555 verifier contract

`verify.py` is the read-only, fail-closed custody checker for the OLE2
physical-sector marker accounting campaign. Its CLI envelope is:

```json
{
  "schema": "litchi.ole2.verification.0555.v1",
  "status": "pass | incomplete | fail",
  "scope": "selected component",
  "result": "present only for pass"
}
```

`--strict` returns 2 when later evidence is absent and 1 for malformed or
contradictory evidence. The verifier never builds, captures, runs quality
commands, applies source, removes the owned target, or writes an evidence
file. It replays the pure metrics and profile consumers in isolated
`python3 -B` interpreters and compares their complete JSON values.

The frozen input envelope is `ole2_0555_frozen_inputs_v1` and contains the
frozen UTC time plus hashes for `plan.json`, `run.py`, `workspace-lock.json`,
and `adr-manifest.json`. `analysis-inputs.json` is the separately frozen
`ole2_0555_analysis_inputs_v1` map of the metrics/profile consumers and every
transitive helper. Every listed path is checked against its current bytes.
The retained same-turn consumer amendments are
`ole2_0555_analysis_amendment_v1` for the two `set(STAGES)` type fixes in
`analyze_metrics.py` and `ole2_0555_profile_amendment_v1` for the per-dump
`claim_sector` state fix in `analyze_profiles.py`. Each amendment records its
original bytes, the failed receipt, and preserved stdout/stderr, and binds the
exact source transformation. A 0554 admission supplement or amendment is not
admissible.

The plan schema is `ole2_physical_marker_0555_plan_v1`. It binds the frozen
Git revision, the single private CFB source file, the six physical roles
(`FAT`, `DIFAT`, `Directory`, `MiniFAT`, `MiniStream`, and `RegularStream`),
semantic and allocation invariants, the nine XLS cases, three CFB shapes,
ABBA order, native and allocator vectors, profile jobs, assembly owners,
admission text, source/lock/ADR bindings, and the owned target
`/home/zhuhe/litchi-goal-0555-target`. The physical accounting scope has no
public API or semantic format change and keeps OLE2/OOXML ahead of deferred
ODF work.

Baseline source must equal the frozen Git tree and have an empty patch.
Candidate source must have the exact changed-path patch and an independent
byte witness. The corrected candidate qualification is retained alongside
the original failed `qualification-01` source, stage manifest, patch witness,
candidate binding, application, review, and failed targeted quality result.
`candidate-correction.json` binds the original and corrected source hashes and
the test-only qualification change; `candidate-correctness.json` binds the
passing corrected targeted checks. A final source is accepted only when its
manifest equals the candidate for an accepted disposition or the baseline for
a rejected disposition, and the live source has that same manifest.

Every build and capture receipt uses `ole2_0555_run_receipt_v1`. The verifier
checks the exact command, output stage, execution stage, source and execution
manifest hashes, binary hash, workspace lock hashes, driver and plan hashes,
the nine-field environment map (including `CARGO_TARGET_DIR`), artifact
digests, UTC interval, and non-overlap. Baseline repeat two stays in the
baseline output folder while its execution stage is candidate. The exact
stage inventory includes both normal and allocator builds, native and
allocator captures, profile artifacts, and assembly artifacts.

Assembly evidence is mandatory in both stages. `assembly-index.json` uses
`ole2_0555_assembly_v1`, binds its normal binary, source manifest, plan,
inspection script, requested owners, and the producer's static-instruction
scope. Every emitted row has a unique requested CFB symbol, an exact
`objdump` receipt, and hashed stdout/stderr/host artifacts. Static bounds do
not replace positive timed owner ancestry; absent or inlined physical leaves
remain indeterminate until same-binary mapping identifies their instructions.

The metrics report uses
`ole2_physical_marker_metrics_0555_v1`. It retains all native, RSS,
allocator, identity, adverse, and repeat-drift rows. Its six numerical gate
groups are primary XLS p50, primary XLS mean, native XLS controls, native CFB
controls, native RSS, and allocation. The four primary XLS p50 rows must
improve by at least 3% in both repeats; all control, RSS, and 72 allocation
rows keep the 5% ceiling. `many-small` is a required CFB control and has no
candidate-specific positive gate. `main_gates.all_frozen_main_gates_pass`
is the numerical result; profile, correctness, review, and quality remain
independent gates. After cleanup, deterministic metrics replay permits only
the four `binaries/{stage}/{kind}/present` observations to change from true
to false, with every other field byte-equal.

Profile stage reports use
`ole2_physical_marker_0555_profile_analysis_v1`; the comparison uses
`ole2_physical_marker_0555_profile_comparison_v1`. Both stages and both
repeats are required. There are eight jobs, 40 positive timed constructor
dumps, and six setup dumps. Owner self, direct, inclusive, calls, physical
targets, setup ancestry, inline context, and moved-work rows stay separate.
Missing, inlined, or non-positive targets carry nullable values and
`indeterminate_assembly_required`, never zero. Profile evidence is
diagnostic and cannot make a native, allocation, physical-I/O, or adoption
claim. No separate instruction-analysis artifact is required by the fresh
0555 consumer freeze.

Quality uses `ole2_0555_quality_plan_v1`,
`ole2_0555_quality_inputs_v1`, `ole2_0555_quality_receipt_v1`, and
`ole2_0555_quality_v1`. Every retained attempt, including the failed
qualification attempt, is checked against its source snapshot, exact command
prefix, receipt, stdout/stderr hashes, and serial intervals. `quality.json`
must be byte-identical to exactly one passing `final`/`commands` attempt.

The adverse review uses `ole2_0555_adverse_review_v1`, retains every
`comparisons.adverse_over_five_percent` and
`repeat_drift_over_five_percent` row one-for-one with a nonempty
classification, interpretation, and disposition, and binds metrics, profile,
and final quality hashes. The decision uses `ole2_0555_decision_v1` and
contains exactly:

```json
{
  "schema": "ole2_0555_decision_v1",
  "status": "pass",
  "observed_utc": "timezone-aware ISO-8601",
  "scope": "nonempty disposition scope",
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

The verifier derives adoption from the numerical gate, the required profile
gate, and the review decision. A failed mandatory gate requires a rejected
disposition and restored baseline source.

Cleanup uses `ole2_0555_cleanup_v1` and the exact fields checked by the
verifier: target, removal, process-reference scope, Python-cache state, the
four binary and descriptor hash maps, the passing precleanup verifier hash,
plan hash, timestamp, and cleanup scope. Descriptors may refer to missing
binaries only after the owned target is absent, process references are empty,
and the precleanup record proves the exact hashes before deletion.

Terminal sealing checks the absent target, no symlinks or Python bytecode
caches, and a sorted recursive `SHA256SUMS` inventory excluding only the seal
itself. Every retained path is therefore covered by a schema, a digest, an
independent source/receipt binding, or the final seal.
