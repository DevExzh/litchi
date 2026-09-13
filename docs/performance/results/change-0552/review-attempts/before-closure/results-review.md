# XLSX batch 0552 results review

**Status:** complete diagnostic review; candidate rejected and production restored.

The canonical main and guard/cap analyzers completed, but the candidate is not admitted. The main gate is false because two repeat-2 dense-sparse RSS rows exceed the 5% workflow-memory limit. Guard admission is false because every valid native p50/mean check fails its 1.05 ratio limit. Cap admission is false because six size-2/size-160 p50 or mean checks exceed the 1.05 ratio limit. The retained public regression tests and all captured evidence remain available for the baseline restoration and final-quality workflow.

This review does not claim a speedup, production adoption, exact-commit profile result, or large-tag guard result. The final-source quality lane completed with all eleven checks passing, including 1,313 XLSX tests; that separate quality result does not override the failed performance and admission gates. The exact commit profile lane and prospective large-tag guard are unmeasured because the main pilot failed. ODF remains deferred until the OLE2/OOXML optimization goal completes; iWork is excluded.

## Outcome-scope cross-check

The outcome-scoped record independently states that seven public regression tests are retained and that no optimized candidate source remains in production after baseline restoration. Its profile decision is explicitly skipped because the pilot failed, and the supplemental large-tag/large-cell guard decision is also skipped with an empty measurement set. Accordingly, this review makes no profiled or large-tag performance claim.

## Gate findings

| Gate | Result | Evidence |
| --- | --- | --- |
| Main primary one-percent | pass (32/32) | `main_gates.primary_one_percent` |
| Main one-cell latency | pass (32/32) | `main_gates.one_cell_latency` |
| Main allocation | pass (32/32) | `main_gates.allocation` |
| Main correctness identity | pass | `main_gates.correctness_identity.identity_equal=true` |
| Main workflow memory | **fail (62/64)** | two RSS rows below |
| Guard admission | **fail** | valid native p50/mean: 0/8 checks pass |
| Cap admission | **fail (14/20)** | six rows below |

The two main workflow-memory failures are:

| Case | Shape | Repeat | Metric | Baseline | Candidate | Delta | Change |
| --- | --- | ---: | --- | ---: | ---: | ---: | ---: |
| `xlsx_source_backed_cell_values_one_edit_save` | `dense-sparse` | 2 | `time_v.max_rss_bytes` | 88363008 | 94523392 | 6160384 | +6.971678% |
| `xlsx_source_backed_managed_cell_values_one_percent_edit_save` | `dense-sparse` | 2 | `time_v.max_rss_bytes` | 87863296 | 95232000 | 7368704 | +8.386555% |

Guard gate inventory (the individual guard adverse and drift rows are retained below in the JSON):

| Guard gate | Checks | Failures |
| --- | ---: | ---: |
| `invalid_incremental_peak` | 8 | 0 |
| `invalid_native_p50_mean` | 16 | 0 |
| `valid_native_p50_mean` | 8 | 8 |
| `valid_noop_incremental_peak` | 4 | 0 |

The six cap gate failures are:

| Size | Repeat | Metric | Baseline | Candidate | Delta | Change |
| ---: | ---: | --- | ---: | ---: | ---: | ---: |
| 2 | 1 | `p50` | 168316 | 180716 | 12400 | +7.367095% |
| 2 | 1 | `mean` | 170136.005 | 182376.515 | 12240.51000000001 | +7.194544% |
| 160 | 1 | `p50` | 11378446 | 14305619 | 2927173 | +25.725596% |
| 160 | 1 | `mean` | 11393173.045 | 14307843.79 | 2914670.744999999 | +25.582608% |
| 160 | 2 | `p50` | 12303734 | 14224180 | 1920446 | +15.608644% |
| 160 | 2 | `mean` | 12304949.03 | 14231062.195 | 1926113.165000001 | +15.653158% |

## Individual row retention

The companion [`adverse-review.json`](adverse-review.json) uses the schema `xlsx_0552_adverse_review_v1` required by `decide.py`. It has six groups. Every canonical source row is copied value-for-value into `original` and receives its own nonempty `id`, `classification`, `interpretation`, and `disposition`; no row is grouped away or removed.

| Group | Canonical source | Rows | Interpretation | Disposition |
| --- | --- | ---: | --- | --- |
| Main adverse | `metrics-analysis.json:comparisons.adverse` | 1085 | candidate comparison; phase/allocator/RSS scope retained per row | retain unresolved |
| Main repeat drift | `metrics-analysis.json:repeat_drift_over_five_percent` | 176 | same-build repeat variation; phase scope retained per row | retain diagnostic |
| Guard adverse | `guard-cap-analysis.json:comparison.guard.adverse_flags_over_five_percent` | 439 | candidate guard comparison; no admission credit | retain unresolved |
| Guard repeat drift | `guard-cap-analysis.json:comparison.guard.same_build_drift_over_five_percent` | 18 | same-build guard variation | retain diagnostic |
| Cap adverse | `guard-cap-analysis.json:comparison.cap.adverse_flags_over_five_percent` | 35 | candidate cap comparison; no admission credit | retain unresolved |
| Cap repeat drift | `guard-cap-analysis.json:comparison.cap.same_build_drift_over_five_percent` | 20 | same-build cap variation | retain diagnostic |
| **Total** | six exact source arrays | **1773** | every row individually interpreted | — |

Negative changes are retained with the same care as positive changes. Main RSS rows are process-lifetime high-water observations. Reopen/open/planning and allocator metrics retain their analyzer scope, and allocator-instrumented elapsed time is excluded from latency interpretation. These scope statements prevent a diagnostic row from being treated as an operation-local attribution.

## Evidence bindings

| Artifact | SHA-256 |
| --- | --- |
| `metrics-analysis.json` | `c0dd9f49b7006c5d426211624e36217fc680af1ad37d1febaa5d82a6f10c2a3c` |
| `guards-analysis.json` | `b59becde85b91b30303a765441450c073ab2167feebb7427423f0509cde002bb` |
| `analyze_metrics.py` | `af6dc11cd31160ac267f695d76ff2422b19cce80a5376b8e704fdff9addca5cc` |
| `analyze_guards.py` amended | `6714a465703be541c969a1256ad3c1d5d21f81387f94aa65d613b03890a79dbc` |
| `analyze_guards.py` original frozen | `c30dbe68e0d0db4ff15ab214f75eeda0a4e5d3ca042f646c6e46d79f5d7bf817` |
| `analyzer-amendment.json` | `decfe401fb348e2e674fad25064546daeed38ef29ac931470d79c13d62ff2dd0` |
| `analysis-inputs.json` | `353a12870da60b4e6ee978168c4dd28eee04129386598b3ebf99c979a3f4adea` |
| `plan.json` | `0fa7066ad4649b07d0064574014215a0372ec0b98ab5eec3466635085b84eb42` |
| `run.py` | `2dac95818ffdc24ced07a87bc8603e490ea1b62b976bfb36e7dcb0ec95a46709` |
| `capture.py` | `21b0153923e87a331906552457a5f77bb444686d4ce9c9c99e3efcb2dbe5d2c1` |
| `guarded_capture.py` | `1e44a1a95d23622b4b3346caef20a6987f620cdbe58b3bb1e827418e71aa7092` |
| `docs/performance/0552-xlsx-compact-source-proof.md` | `e2bff34b02ae9987d5fd564ec9c7261ff5d984211b5cd0a51d25a8db8b84a1c4` |
| `quality.json` (all eleven checks pass) | `8a47d3a452e2db1cb6db3b1e64c70b4d1dd9caa1f5d399bff165a9b61b6beaaa` |
| `profile-decision.json` (skipped) | `b243cd63327246e60bf69e9c8efee4b23f9076bee38788e119e3547fe4298b4a` |
| `supplemental-guard-decision.json` (skipped) | `7bd0161248b82b6505119f428b2f0bbe0ee21b704f972500523d0aa947dd9b44` |

The guard analyzer amendment is preserved as an explicit metadata-only serialization correction: the original frozen analyzer and failed receipt remain bound, while the amended analyzer serializes only the two interval timestamps as ISO strings. The amendment record states that numerical results, gate logic, receipt validation, and ordering were unchanged.

No build, test, capture, or Rust source operation was performed for this review. The root coordinator owns those receipts and final source restoration; the bound final-quality artifact reports all eleven checks passing.
