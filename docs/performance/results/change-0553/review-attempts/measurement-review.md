# XLSX batch 0553 results review

**Status:** complete diagnostic review; candidate rejected by mandatory main gates; final-source quality pending.

The canonical main metrics report completed with status `pass`, but its frozen admission result is false. Three unmanaged vendor-extension one-percent workflow rows fail the required 3% improvement threshold, and one unmanaged dense-sparse repeat-2 RSS row exceeds the 5% workflow-memory limit. The guard and cap canonical reports both complete with admission `true`; that does not override the failed main gates, so `adoption_allowed` is false.

Final-source quality is still pending in this review. The conditional profile lane is skipped because the main pilot is false, and no profile or prospective large-tag result is claimed. Baseline restoration and the quality receipt remain root-owned follow-up evidence. OLE2 and OOXML remain the performance priority; ODF is deferred and iWork is excluded.

## Gate findings

| Gate | Result | Evidence |
| --- | --- | --- |
| Main primary one-percent | **fail (29/32)** | three vendor-extension workflow rows below |
| Main one-cell latency | pass (32/32) | `main_gates.one_cell_latency` |
| Main allocation | pass (32/32) | `main_gates.allocation` |
| Main workflow memory | **fail (63/64)** | one RSS row below |
| Main correctness identity | pass | `main_gates.correctness_identity` |
| Guard admission | pass | `comparison.guard.admission_passed=true` |
| Cap admission | pass | `comparison.cap.admission_passed=true` |

The three primary one-percent failures are:

| Case | Shape | Repeat | Metric | Baseline | Candidate | Delta | Change | Criterion |
| --- | --- | ---: | --- | ---: | ---: | ---: | ---: | --- |
| `xlsx_source_backed_cell_values_one_percent_edit_save` | `vendor-extension` | 1 | `timing.workflow.p50` | 20420591 | 19898506 | -522085 | -2.556660% | `candidate change <= -3.00%` |
| `xlsx_source_backed_cell_values_one_percent_edit_save` | `vendor-extension` | 2 | `timing.workflow.p50` | 20286930 | 19976046 | -310884 | -1.532435% | `candidate change <= -3.00%` |
| `xlsx_source_backed_cell_values_one_percent_edit_save` | `vendor-extension` | 2 | `timing.workflow.mean` | 20287773.200000014 | 19983204.10000001 | -304569.1 | -1.501245% | `candidate change <= -3.00%` |

The workflow-memory failure is:

| Case | Shape | Repeat | Metric | Baseline | Candidate | Delta | Change | Criterion |
| --- | --- | ---: | --- | ---: | ---: | ---: | ---: | --- |
| `xlsx_source_backed_cell_values_one_percent_edit_save` | `dense-sparse` | 2 | `time_v.max_rss_bytes` | 88113152 | 92663808 | 4550656 | +5.164559% | `candidate change <= 5.00%` |

## Individual row retention

The companion [`adverse-review.json`](adverse-review.json) follows the exact `xlsx_0553_adverse_review_v1` envelope required by `verify.py` and `decide.py`. Each canonical row is copied value-for-value into `original` and receives its own nonempty `id`, `classification`, `interpretation`, and `disposition`. The six source arrays retain all 346 rows individually; no adverse or over-five-percent drift row is hidden behind an aggregate.

| Group | Canonical source | Rows | Interpretation | Disposition |
| --- | --- | ---: | --- | --- |
| Main adverse | `metrics-analysis.json:comparisons.adverse` | 55 | matched candidate comparison with phase scope | retain unresolved |
| Main repeat drift | `metrics-analysis.json:repeat_drift_over_five_percent` | 234 | same-build repeat variability | retain diagnostic |
| Guard adverse | `guard-cap-analysis.json:comparison.guard.adverse_flags_over_five_percent` | 2 | candidate guard comparison | retain unresolved |
| Guard repeat drift | `guard-cap-analysis.json:comparison.guard.same_build_drift_over_five_percent` | 22 | same-build guard variability | retain diagnostic |
| Cap adverse | `guard-cap-analysis.json:comparison.cap.adverse_flags_over_five_percent` | 4 | candidate cap comparison | retain unresolved |
| Cap repeat drift | `guard-cap-analysis.json:comparison.cap.same_build_drift_over_five_percent` | 29 | same-build cap variability | retain diagnostic |
| **Total** | six canonical source arrays | **346** | every row individually interpreted | — |

Positive and negative changes are retained alike. RSS is process-lifetime high water, allocator-instrumented elapsed is excluded from latency, reopen/open and planning retain their phase scope, and same-build drift is diagnostic variability. These interpretations preserve the analyzer limits without turning diagnostics into operation-local attribution.

## Evidence bindings

| Artifact | SHA-256 |
| --- | --- |
| `metrics-analysis.json` | `147abc506cde809dfe7f20368a5cfc567629b5f58feffeeba09d23ad1c9fb02a` |
| `guard-cap-analysis.json` | `3dfc8654c49ef8aa2bd8cd000471b961337fc9b34c201d148c24e2cdc35fc85b` |
| `analyze_metrics.py` | `69042472302568c42241fbfcbdad8a8c683f3134e7e7c0ebf9c8ed5ec8f42426` |
| `analyze_guards.py` | `89c9ed81e66ff3ea35505cfa9aaf0c9ffefdfa4292ccff3757f88c2b64edb2d0` |
| `analysis-inputs.json` | `2d95091698893cf402ce85003f407e0352a80b3fd70fa5c4c5e3110f8332a38a` |
| `plan.json` | `17d810a24912065fde8c71de6109be983f7c5b3bdcea5fd6584f4725c6499ea4` |
| `run.py` | `f1398fcc87dc7bccfc05930276e26a4a8a21106480613238eab49310ae1c23f0` |
| `capture.py` | `21b0153923e87a331906552457a5f77bb444686d4ce9c9c99e3efcb2dbe5d2c1` |
| `guarded_capture.py` | `ba3a2dc8c8f43531dfdd099310e19aeb9ce6ce160b5ce89dc577e21fc85f4119` |
| `adr-manifest.json` | `c4e7331b91816d3752917c3cb4b320ed56c35efede61baa9170d215546e10ad2` |
| `profile-decision.json` (skipped) | `71bdc6c947a99389c3cb5e9c314cf179134ffd92c0753189e03170e018ca5328` |

The metrics analyzer binds the four main failure rows to the frozen plan and matched source identity. The guard/cap analyzer binds its passing admissions to the same plan and guarded capture. No analyzer, capture, source, or gate was changed for this review. Final quality remains pending and must be evaluated on the selected restored baseline source.

No build, test, capture, or Rust source operation was performed for this review; the root coordinator owns those receipts and the final quality decision.
