# 0553 verifier contract

`verify.py` is a read-only, fail-closed checker for the commit-local XLSX
source-cell experiment. Its output envelope is:

```json
{
  "schema": "litchi.xlsx.verification.0553.v1",
  "status": "pass | incomplete | fail",
  "scope": "selected component",
  "result": "present only for pass"
}
```

An incomplete component means that required evidence has not arrived. It never
creates synthetic measurements, treats missing rows as zeros, or turns a
pending analyzer into a pass. `--strict` returns exit 2 for incomplete evidence
and exit 1 for malformed or contradictory evidence. The verifier does not run
Cargo, Rust, capture drivers, or candidate code. It may replay a candidate
patch into a temporary Git index for custody validation.

The frozen 0553 inputs are the exact mappings in `frozen-inputs.json`,
`analysis-inputs.json`, and `supplemental-inputs.json`. The verifier checks
their hashes, the `8aa0c5baf0616d16c79eba0c6c28dc1716338ad6` baseline tree, the
ADR manifest, the ignored `Cargo.lock` binding, the 0553 plan and quality
matrix, and the host envelope. The baseline manifest must equal the frozen
Git tree and have an empty `source.patch`. A candidate must have a changed
source manifest, a nonempty diff whose paths equal the manifest delta, and a
matching live, `candidate-sources/`, or retained-attempt snapshot witness for
every changed source. The witness lookup selects bytes matching the bound
candidate digest, so a later baseline restoration cannot make retained
candidate custody fail.

`baseline-correctness.json` is the only historical reuse accepted by this
contract. It uses schema `xlsx_0553_baseline_correctness_v1`, binds the current
baseline to the committed 0552 final manifest, binds the prior passing quality
result and all eleven prior check receipts/artifacts, and reports exactly 59
test groups, 1313 passed, 0 failed, and 0 ignored. Its scope explicitly says
that those commands were not rerun for 0553 and that it does not validate
candidate source. No 0552 public-test bundle, draft-attempt assumption, or
analyzer-amendment path is accepted.

## Candidate custody before capture

Performance reports cannot be admitted from a source patch, capture matrix,
and quality result alone. `validate_metrics_analysis()` and
`validate_guard_cap_analysis()` first require `validate_candidate_custody()`.
Missing candidate custody is therefore `incomplete`, with no canonical
performance or disposition outcome.

Every retained `candidate-attempts/draft-NN` is an immutable iteration record.
The first draft uses schema `xlsx_0553_candidate_draft_binding_v1` and exact
files `binding.json`, `application.json`, and `candidate.patch`. Later drafts
also bind their parent binding digest and `transition.patch`; the transition
must replay from the parent candidate to the full candidate patch. Draft
bindings must include:

* the frozen revision, plan digest, baseline digest, and workspace-lock digest;
* `source_hashes` whose changed paths and content hashes exactly match a
  temporary-index replay of `candidate.patch`;
* `limits` exactly equal to
  `{"source_bytes": 8388608, "events": 131072,
  "logical_proof_bytes": 2097152}`; and
* a freeze time before the corresponding `application.json` time.

An application record uses schema `xlsx_0553_candidate_application_v1` and
binds its draft binding digest, patch digest, `source_verified: true`, and
the actual application time. A recorded preapplication attempt must either use
the detailed unchanged-parent/old-receipt envelope or the bounded
no-check-started observation envelope (`recorded_utc`, `status`,
`observed_commands`, `observation`, `checks_started`). The latter uses the
exact status `previous global compiler guard stopped before application`, a
nonempty command list and observation, `checks_started: false`, and a recorded
time between the draft freeze and application. Neither form can count as
validation of the later draft; the verifier checks both forms even when a
later iteration supersedes them.

The final `candidate-binding.json` is required after the selected iteration is
known. It uses schema `xlsx_0553_candidate_binding_v1` and exactly binds:

```text
schema, frozen_utc, recorded_utc, scope, plan_sha256, base_revision,
baseline_manifest_sha256, selected_attempt, selected_binding_sha256,
application_sha256, candidate_manifest_sha256, candidate_patch_sha256,
source_hashes, limits, first_application, check_manifest, status
```

`frozen_utc` is the selected draft's source/cap freeze time. `recorded_utc` is
the time the aggregate binding was actually written and must be at or after
the selected application. `first_application` names and hashes the selected
application record and repeats its application time and `source_verified`
boolean. `check_manifest` names and hashes a retained
`check-attempts/*/source-manifest.json` whose exact map is the complete
candidate source map plus the bound ignored `Cargo.lock`. This keeps the
preapplication freeze claim separate from the later canonical aggregate
document and prevents backdating it.

The final `candidate-correctness.json` is also mandatory. It uses schema
`xlsx_0553_candidate_correctness_v1`, status `pass`, and binds the final
candidate-binding digest and candidate source-manifest digest. Its
`independent_review` object has exactly `path`, `sha256`, `status`, and
`reviewer`; the review is a retained regular file with matching digest and
`status: "pass"`. Its `checks` rows have exactly:

```text
kind, name, attempt, receipt_sha256, source_manifest_sha256,
exit_code, source_stable
```

Each row points under `check-attempts/` and is validated with the 0553
check-attempt schema against the complete candidate source map plus the bound
ignored `Cargo.lock`. At least one row of each kind is required:
`compact-success`, `compact-fallback`, and `resource`. Every receipt must pass,
be source-stable, and finish before the correctness record. These rows are
direct correctness evidence; a failed or unchanged-source preapplication
attempt cannot substitute for them.

## Captures and analyzers

The complete main matrix is 16 preflight, 32 native, and 32 allocator jobs per
stage. Supplemental custody is 12 normal guard, 12 allocator guard, and 10 cap
jobs per stage. Every receipt binds the exact 0553 run/plan hash, source and
execution manifests, retained binary, command, local artifacts, and serial
interval. The metrics analyzer must replay to
`xlsx_multisource_edit_metrics_0553_v1`; the guard analyzer must replay to
`litchi.xlsx.guard-cap-analysis.v1`. Existing canonical aliases are accepted
only when byte-identical to deterministic replay. The frozen analyzer hashes
must match `analysis-inputs.json`; there is no post-capture amendment path.

The independent baseline R1 validation record is supportive custody only. It
does not replace the complete baseline/candidate capture matrix or either
canonical analysis report.

## Conditional profiles and review

The profile pilot is exactly:

```text
main.main_gates.all_frozen_main_gates_pass
AND guard.comparison.guard.admission_passed
AND guard.comparison.cap.admission_passed
```

`validate_profiles()` accepts both the verifier's `(metrics, guards)` call and
the decision consumer's `pilot_expected=` call; the latter recomputes the
frozen reports and checks that expectation. A false pilot requires
`profile-decision.json` schema `xlsx_0553_profile_decision_v1`,
`status: "skipped"`, `profile_required: false`,
`profile_gate_passed: true`, empty `profile_rows`, and no profile receipts. A
true pilot requires all sixteen profile receipts and a strict `Ir` reduction
for every matched shape/repeat row.

`adverse-review.json` is mandatory before disposition. It uses schema
`xlsx_0553_adverse_review_v1`, binds the metrics and guard report digests, has
`status: "complete"`, `complete: true`,
`all_diagnostic_rows_retained: true`, a boolean `adoption_allowed`, and six
one-for-one reviewed groups covering the two adverse/drift arrays from the
metrics report and the four guard/cap adverse/drift arrays. Every source row
must be retained exactly once with nonempty `id`, `classification`,
`interpretation`, and `disposition` fields.

## Quality and disposition

Quality uses the 0553 `quality.py` and `check_attempt.py` schemas. A passing
canonical `quality.json` must be an exact copy of one passing final-source
quality attempt, with all eleven commands successful and source-stable. The
decision consumer compatibility API is explicit: `stage_manifest(stage)`
returns `(manifest, digest)`, `validate_profiles(pilot_expected=...)` accepts
the conditional pilot expectation, and `validate_final_source(disposition,
candidate_manifest, candidate_digest)` validates the final checkout. The
verifier also requires the final decision's source digest to be present and
exact.

Adoption is the conjunction of main, guard, cap, conditional profile, quality,
and adverse-review gates. `accepted` requires the final source manifest to
equal candidate. `rejected` requires the final source to equal the restored
baseline with an empty patch. The live source must equal the selected final
manifest. OLE2/OOXML remains the active priority; ODF is deferred until that
optimization goal completes.

## Terminal cleanup, documentation, and seal

The verifier has separate `precleanup` and `all` components. `precleanup`
requires `/home/zhuhe/litchi-goal-0553-target` to be a present, non-symlink
directory and validates the complete campaign, including exactly ten retained
binary descriptors and live binary hashes: `baseline` and `candidate` crossed
with `normal`, `alloc`, `guard-normal`, `guard-alloc`, and `cap`. It also runs
the complete report, conditional-profile, quality, review, disposition, and
external-documentation checks. This is the only stage that can establish live
owned-target binary custody before removal.

The root cleanup step is outside this read-only verifier. It removes only the
owned target after a passing precleanup run, checks accessible process cwd,
executable, and open-descriptor references, and records the result as
`cleanup.json`. The exact schema is:

```json
{
  "schema": "xlsx_0553_cleanup_v1",
  "observed_utc": "ISO-8601 timestamp with timezone",
  "plan_sha256": "sha256(plan.json)",
  "target": "/home/zhuhe/litchi-goal-0553-target",
  "removed": ["/home/zhuhe/litchi-goal-0553-target"],
  "owned_paths_absent": true,
  "accessible_process_references": [],
  "process_reference_scope": "Accessible /proc cwd, executable and open file descriptors; cleanup process ancestors excluded.",
  "python_cache_absent": true,
  "binary_sha256_by_kind": {
    "baseline/normal": "...",
    "baseline/alloc": "...",
    "baseline/guard-normal": "...",
    "baseline/guard-alloc": "...",
    "baseline/cap": "...",
    "candidate/normal": "...",
    "candidate/alloc": "...",
    "candidate/guard-normal": "...",
    "candidate/guard-alloc": "...",
    "candidate/cap": "..."
  },
  "binary_descriptor_sha256_by_kind": {
    "baseline/normal": "sha256(baseline/binary-normal.json)",
    "baseline/alloc": "sha256(baseline/binary-alloc.json)",
    "baseline/guard-normal": "sha256(baseline/binary-guard-normal.json)",
    "baseline/guard-alloc": "sha256(baseline/binary-guard-alloc.json)",
    "baseline/cap": "sha256(baseline/binary-cap.json)",
    "candidate/normal": "sha256(candidate/binary-normal.json)",
    "candidate/alloc": "sha256(candidate/binary-alloc.json)",
    "candidate/guard-normal": "sha256(candidate/binary-guard-normal.json)",
    "candidate/guard-alloc": "sha256(candidate/binary-guard-alloc.json)",
    "candidate/cap": "sha256(candidate/binary-cap.json)"
  },
  "precleanup_verification_sha256": "sha256(precleanup-verification.json)",
  "scope": "Remove only owned change0553 build/retained-binary target after passing precleanup and exact binary hash checks."
}
```

Both hash maps must contain exactly those ten `stage/kind` keys. Every
descriptor remains bound to its stage source manifest and build receipt. If a
binary is absent after cleanup, the capture and profile validators accept its
descriptor digest only through this validated cleanup record; no missing
binary is accepted by an unconditional allow-missing branch. In particular,
post-cleanup profile validation obtains the normal-binary digest from the
validated descriptor custody and keeps receipt commands source-bound.

`documentation-manifest.json` is required for both terminal phases and uses
the exact envelope:

```json
{
  "schema": "xlsx_0553_documentation_manifest_v1",
  "scope": "Files outside this evidence bundle included in completed change0553",
  "files": {
    "docs/performance/0553-xlsx-commit-local-compact-proof.md": "sha256"
  }
}
```

`files` must be a nonempty, sorted map of regular, non-symlink files under
`docs/` outside `docs/performance/results/change-0553/`, with live hashes. The
0553 outcome document above is mandatory; any additional completed-work index
or HOTSPOTS documentation can be listed with its exact hash.

The passing precleanup CLI output must be retained as
`precleanup-verification.json` inside this bundle before target removal. Its
envelope is the normal verifier envelope with `scope: "precleanup"`; its
result additionally contains `observed_utc`, `phase: "precleanup"`,
`target_present: true`, and the complete campaign result. `cleanup.json` binds
its exact digest through `precleanup_verification_sha256` and rechecks all ten
binary rows against that report. This preserves the before-deletion custody
record in the sealed bundle.

After cleanup and the terminal campaign checks, root creates `SHA256SUMS`.
Each line is `<lowercase-sha256>  <safe path relative to this bundle>`, sorted
by relative path, with a final newline. The inventory must contain every
regular file recursively under `docs/performance/results/change-0553/` except
`SHA256SUMS` itself, including `verify.py`, `verifier-schema.md`, custody
attempts, reports, cleanup, and the documentation manifest. Symlinks,
non-files, and `__pycache__` entries fail the seal. The verifier compares the
manifest with the live recursive inventory and never creates or repairs it.
The pre-edit verifier custody is retained at
`verifier-attempts/initial/verify.py` with matching hash metadata in
`verifier-attempts/initial/inputs.json`; both are part of the recursive seal.

The root commands are read-only verifier invocations; redirect output outside
the evidence bundle:

```sh
# Retain this passing output in the evidence bundle before deleting the target.
python3 -B docs/performance/results/change-0553/verify.py \
  --component precleanup --strict \
  > docs/performance/results/change-0553/precleanup-verification.json
# Root removes only /home/zhuhe/litchi-goal-0553-target and records cleanup.json.
python3 -B docs/performance/results/change-0553/verify.py \
  --component cleanup --strict
python3 -B docs/performance/results/change-0553/verify.py \
  --component all --strict
```

`cleanup` checks only the terminal cleanup record and descriptor custody;
`seal` checks only the recursive seal. `all` requires cleanup first, then
revalidates all campaign evidence against validated cleanup descriptors, and
finally validates the seal. These commands do not build, capture, delete,
write reports, or modify the checkout.
