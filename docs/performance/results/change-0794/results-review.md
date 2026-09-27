# Independent numeric review — change 0794

Status: **REJECT; restore the baseline source.** The primary PPTX campaign
does not satisfy the frozen useful-benefit requirement, and the separate
DOCX/XLSX control lane triggers its frozen regression veto. This is an
offline review of the retained raw reports and custody summaries. It adds no
measurement, does not pool historical timings, and does not change the
adoption policy.

## Evidence and scope

| Lane | Reports | Measured samples | Scope |
| --- | ---: | ---: | --- |
| Qualification | 15 | 15 | One baseline sample per primary case |
| Primary native | 180 | 5,400 | Six paired blocks × 15 cases × 30 samples |
| Primary allocation | 60 | 180 | Two paired blocks × 15 cases × 3 samples |
| **Primary total** | **255** | **5,595** | |
| Heap profile | 4 | 4 | Two alternating large-capture pairs; 8 decodes |
| Cross-format qualification | 1 | 8 | One before-only report |
| Cross-format native | 12 | 2,880 | Six paired blocks × eight rows × 30 samples |
| **Cross-format total** | **13** | **2,888** | Separate regression-veto lane |
| Helper supplement | 2 | 84 iterations | Fourteen counts × three repeats × two legs |

The primary report walk retains deterministic output and semantic readback
identities for all 15 cases. The six quality gates pass with 12,865 tests
passed, 0 failed, and 89 ignored. The candidate changed five canonical
`xml_attributes.rs` copies; the formula helper remained unchanged. The live
five-file production set now hashes exactly to the before source census after
the rejected disposition.

## Primary timing recomputation

I recomputed nearest-rank p50 values directly from the 180 native reports,
formed six after/before block ratios per case, and bootstrapped the median
with `random.Random(794079)`, 10,000 resamples, and sorted zero-based
endpoints 250 and 9,749. The table reports the median of the six per-block
p50 values in milliseconds; the interval is the ratio interval.

| Shape | Mode | Before p50 | After p50 | Change | 95% ratio CI |
| --- | --- | ---: | ---: | ---: | --- |
| tiny | capture | 0.235557 | 0.236321 | +0.3245% | [0.998714, 1.008161] |
| tiny | commit | 0.211391 | 0.213806 | +1.2188% | [1.007403, 1.013573] |
| tiny | lifecycle | 1.409536 | 1.414007 | +0.4200% | [1.000503, 1.005516] |
| medium | capture | 0.457612 | 0.461272 | +0.8796% | [0.999242, 1.012941] |
| medium | commit | 0.295856 | 0.298271 | +0.7505% | [1.006624, 1.009021] |
| medium | lifecycle | 2.008955 | 2.020000 | +0.3752% | [1.001574, 1.006802] |
| large | capture | 19.087100 | 19.294511 | +1.0042% | [0.992390, 1.016934] |
| large | commit | 1.293866 | 1.319511 | +2.1220% | [1.011487, 1.024562] |
| large | lifecycle | 28.493757 | 28.893127 | +1.3220% | [1.009677, 1.015075] |
| vendor | capture | 0.540872 | 0.550922 | +2.0283% | [1.015033, 1.025702] |
| vendor | commit | 0.323942 | 0.328487 | +1.2706% | [1.008746, 1.017250] |
| vendor | lifecycle | 2.159711 | 2.180686 | +1.0229% | [1.008927, 1.011287] |
| unicode-vendor | capture | 0.544068 | 0.557863 | +2.5576% | [1.020974, 1.031271] |
| unicode-vendor | commit | 0.323791 | 0.328402 | +1.2258% | [1.008420, 1.016516] |
| unicode-vendor | lifecycle | 2.165860 | 2.188035 | +0.9008% | [1.003329, 1.013694] |

There are no primary latency violations: no row has a paired p50 ratio above
the frozen 1.05 threshold. There are also no eligible benefits: no capture or
lifecycle row reaches the required 3% improvement with a bootstrap upper
bound below 1.0. The primary adoption guard is therefore false before the
cross-format veto is considered.

The allocation reports retain raw `allocation_calls`, `allocated_bytes`,
`net_live`, and `peak_above_entry` semantics separately from elapsed time.
Comparing paired block medians gives no after increase for any of the four
frozen resource guards. This passes the resource guard but supplies no
benefit evidence.

## Cross-format veto

I independently read the 12 native cross-format reports, reconstructed each
30-sample p50 as the integer midpoint of sorted samples 14 and 15 (zero-based),
and recomputed the six-ratio bootstrap with seed `794080`. The frozen cross
gate rejects when the p50 ratio exceeds 1.05 and the lower confidence bound
exceeds 1.0.

| Case | Shape | Ratio median | 95% ratio CI | Veto |
| --- | --- | ---: | --- | --- |
| docx_semantic_full_text | large | 0.975554 | [0.960779, 0.996907] | no |
| docx_semantic_full_text | tiny | 0.960545 | [0.945358, 0.970282] | no |
| docx_semantic_open | large | 0.989853 | [0.965642, 1.001852] | no |
| docx_semantic_open | tiny | 0.988869 | [0.979907, 0.997013] | no |
| xlsx_full_cell_scan | tiny | **1.055109** | **[1.050283, 1.081791]** | **yes** |
| xlsx_full_cell_scan | dense-wide | 1.005361 | [0.998136, 1.006945] | no |
| xlsx_open_owned | tiny | 1.008009 | [1.002853, 1.011804] | no |
| xlsx_open_owned | dense-wide | 1.002026 | [0.977090, 1.042027] | no |

The `xlsx_full_cell_scan/tiny` row is the sole veto row. The cross lane is a
regression control, so it adds no benefit requirement and is not pooled with
the primary PPTX timings. The independent [cross-root-audit.json](cross-root-audit.json)
and [cross-analysis.json](cross-analysis.json) agree on the row and decision.

## Profile and helper boundaries

The four heap-profile reports qualify all four owner/whole-stack checks. The
owner cost matches the declared `allocation_calls` counter: 72,106 before and
10,781 after, an observed reduction of 85.048%. Whole-process stack totals
are 608,290 and 241,608 respectively. Nested stack costs overlap, and the
profile packet makes no latency or RSS claim; these counts are therefore
diagnostic and outside the adoption guards.

The helper supplement has 84 count/layout iterations and reports no increase
in the guarded allocation fields. Its iterator-size observation changes from
120 to 192 bytes, but the supplement declares no latency claim and cannot
substitute for the 15-case policy result.

## Disposition

The frozen primary policy requires at least one named capture or lifecycle
case to improve by 3% with bootstrap high below 1.0, rejects any primary p50
regression above 5% with bootstrap low above 1.0, and forbids increases in
the four paired allocation medians. The primary campaign has no qualifying
benefit, while the separate cross-format lane has one significant veto row.
The candidate is consequently rejected. The five production files are
restored to the exact baseline source census, while candidate before/after
copies, patch, failed quality attempts, reports, and replay audits remain
retained for custody review.

The result applies to the frozen generated warm borrowed-input workflows and
the named DOCX/XLSX controls. It does not establish a general workload gain,
a cold-start or concurrency result, a profile-derived causal attribution, or
a production adoption claim.
