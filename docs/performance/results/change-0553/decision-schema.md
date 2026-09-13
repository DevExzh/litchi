# 0553 decision contract

`decide.py` consumes completed evidence for the commit-local XLSX compact
source-cell collector campaign. It is outcome neutral and fail closed: a
missing, stale, or contradictory artifact produces no `decision.json`.
`--schema` is the read-only contract inspection mode.

The frozen bindings are:

| input | SHA-256 |
| --- | --- |
| `plan.json` | `17d810a24912065fde8c71de6109be983f7c5b3bdcea5fd6584f4725c6499ea4` |
| `run.py` | `f1398fcc87dc7bccfc05930276e26a4a8a21106480613238eab49310ae1c23f0` |
| `capture.py` | `21b0153923e87a331906552457a5f77bb444686d4ce9c9c99e3efcb2dbe5d2c1` |
| `analysis-inputs.json` | `2d95091698893cf402ce85003f407e0352a80b3fd70fa5c4c5e3110f8332a38a` |

The campaign revision is
`8aa0c5baf0616d16c79eba0c6c28dc1716338ad6`. Its scope is matched source
backed XLSX `MultiSourceEdit`; OLE2/OOXML has priority, ODF remains deferred,
and iWork is excluded.

## Required analyzer evidence

The verifier must return completed-pass references to these canonical reports:

| lane | report schema | decision gate |
| --- | --- | --- |
| main | `xlsx_multisource_edit_metrics_0553_v1` | `main_gates.all_frozen_main_gates_pass` |
| guard/cap | `litchi.xlsx.guard-cap-analysis.v1` | `comparison.guard.admission_passed` and `comparison.cap.admission_passed` |

Each reference includes a path and SHA-256 that match the regular file in this
bundle. Main gate inventory remains the frozen five groups plus
`all_frozen_main_gates_pass` and the pending external-control record. The
correctness identity and every exact check row are retained. The canonical
metrics report remains the authority for all numeric comparisons, repeat drift,
source/output/semantic identities, and allocation fields. The decision carries
the complete report reference and the comparison, drift, allocation, and field
path data needed to locate every metric; it does not replace those reports with
a selected statistic.

The guard/cap report must retain native and allocation comparisons, every gate
row with an explicit `passed` boolean, adverse arrays, same-build drift arrays,
and ABBA evidence. Allocator-instrumented elapsed time is validated as receipt
data and is not used as latency evidence.

## Conditional profile contract

The pilot is exactly this conjunction:

```text
main.main_gates.all_frozen_main_gates_pass
and guard.comparison.guard.admission_passed
and cap.comparison.cap.admission_passed
```

`verify.validate_profiles(pilot_expected=...)` must receive that computed
boolean. When it is false, `profile-decision.json` must use schema
`xlsx_0553_profile_decision_v1` and contain exactly these fields:

```json
{
  "schema": "xlsx_0553_profile_decision_v1",
  "status": "skipped",
  "scope": "nonempty explanation",
  "pilot_passed": false,
  "profile_required": false,
  "profile_gate_passed": true,
  "main_analysis_sha256": "<metrics report SHA-256>",
  "pilot_gates": {"main": false, "guard": false, "cap": false},
  "profile_rows": [],
  "reason": "nonempty explanation"
}
```

The three booleans in `pilot_gates` are the independently computed values, so
the example values above are illustrative. Profile receipts are forbidden
after a skipped pilot. When the pilot is true, profiles are required for every
frozen shape and repeat; each retained profile row has an explicit `passed`
boolean and the profile decision binds the main report digest. The decision
serializer preserves verifier capture interval endpoints as timezone-aware
ISO-8601 strings using `datetime.isoformat()`. It rejects naive datetimes and
every unknown object type, so a true-pilot profile row cannot disappear at
JSON output.

## Exact adverse and drift review

`adverse-review.json` uses schema `xlsx_0553_adverse_review_v1`, has
`status=complete`, `complete=true`, `all_diagnostic_rows_retained=true`, and a
boolean `adoption_allowed`. It binds `comparison_sha256` to the main report
and `guard_analysis_sha256` to the guard/cap report. It retains six groups:

1. `metrics-analysis.json:comparisons.adverse`
2. `metrics-analysis.json:repeat_drift_over_five_percent`
3. `guard-cap-analysis.json:comparison.guard.adverse_flags_over_five_percent`
4. `guard-cap-analysis.json:comparison.guard.same_build_drift_over_five_percent`
5. `guard-cap-analysis.json:comparison.cap.adverse_flags_over_five_percent`
6. `guard-cap-analysis.json:comparison.cap.same_build_drift_over_five_percent`

Each group must have the same length as its source array. Every source row is
matched exactly once by an `original` object and carries nonempty `id`,
`classification`, `interpretation`, and `disposition` fields. A review cannot
drop a row or make an incomplete report admissible.

## Source and quality custody

The verifier dependencies are:

```text
validate_metrics_analysis()
validate_guard_cap_analysis()
validate_profiles(pilot_expected=bool)
validate_final_source(disposition, candidate_manifest, candidate_sha)
validate_quality()
stage_manifest(stage)
current_source_manifest()
```

`validate_final_source` is authoritative for patch and source custody, and the
decision rechecks its returned final manifest against the stage manifests and
live source. For acceptance, the final manifest and digest must equal the
candidate manifest exactly. For rejection, the final manifest must equal the
baseline exactly, or differ only by an explicitly verifier-bound retained test
set whose paths and hashes are present in the returned custody record. There is
no implicit prior-campaign public bundle or serialized-analyzer allowance;
the baseline revision already tracks its seven public guard tests.

`quality.json` must use schema `xlsx_0553_quality_v1`, have `status=pass`,
`source_stage=final`, and bind `source_manifest_sha256` to the final manifest.
The verifier must identify the passing final quality attempt. Quality is
required for both an accepted candidate and a rejected/restored baseline.

The final disposition is `accepted` only when every main, guard, cap,
conditional profile, quality, and adverse-review gate passes. Otherwise it is
`rejected` only after final source restoration and final quality custody pass.
No synthetic decision or new performance/baseline claim is emitted.
