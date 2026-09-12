# 0543 measured planning scope and residual priority

The conditional planning profile was admitted after the native, allocation,
and refusal-guard gates passed. This review covers the frozen two-shape,
two-repeat matrix in [`plan.json`](plan.json) and the direct symbol
`litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets`. Callgrind
Ir is retained as an attribution diagnostic; it is not converted to latency.

## Profile custody and edge checks

The frozen plan digest is
`91532f97396a98f5fc0693117126db9819de647ddcad4989a936b5fd82db4fcd`. The
symbol observation digest is
`a161bdd3710bab0f459963dd6e69c65e49a68b48c32d3a5fe062e5078869bfec`, and its
baseline and candidate `nm -C --defined-only` outputs each contain two
matching `edit_sheets` symbols. The eight profile receipts all exit 0, have
the plan digest, and have complete artifact inventories with matching hashes.
Baseline profiles intentionally report the baseline source manifest while
running under the candidate checkout; candidate profiles report the candidate
manifest for both fields. The eight receipt intervals are serial and do not
overlap.

An independent raw replay checked all eight profile jobs. Each job has four
numbered `callgrind.N` dumps and a base termination dump. Parts 1–3 contain no
positive edge from
`litchi_perf_baseline::run_xlsx_cell_values_edit_save` to the selected owner;
each has exactly one lifecycle-parent call whose inclusive Ir equals its
summary. Part 4 has exactly one measured-parent call whose inclusive Ir equals
its summary, and is the only selected part. Every dump uses the Ir event and
the exact owner trigger. Each base dump has `Program termination` and zero Ir.

The retained annotations also replay deterministically:
`analyze_planning.analyze(False)` equals
[`planning-profile-analysis.json`](planning-profile-analysis.json), whose
digest is
`d66ee3cf569401665aebf3c0ba821f6af4ded567fa908727da76242f260ed1c1`. The
report status is `pass` and every pair clears the frozen 5% Ir gate.

| Shape | Repeat | Baseline Ir | Candidate Ir | Reduction |
| --- | ---: | ---: | ---: | ---: |
| medium | 1 | 125,629,180 | 94,323,634 | 24.9190% |
| dense-sparse | 1 | 237,110,670 | 176,995,199 | 25.3533% |
| medium | 2 | 125,633,048 | 94,329,103 | 24.9170% |
| dense-sparse | 2 | 237,096,082 | 176,995,263 | 25.3487% |

The retained normal profile JSON preserves shape identity across both stages
and repeats. The medium output digest is
`fc742a7ad139e2ce2e75fabfff8d2a33408dd5eeb10997717a2e2656d8423e6c` with 205
reads, 4,233,109 read bytes, 93 updates, and four selected worksheets. The
dense-sparse output digest is
`07eee3f980949f031776087ed61f06125edd82cb29394a608a3d8d5a4b28374d` with 206
reads, 4,258,543 read bytes, 178 updates, and four selected worksheets. The
normal profile's planning, commit, and publication allocation vectors are
all explicitly `unavailable`; allocation conclusions come from the separate
allocator lane.

## Remaining planning owners

Before the change, `Snapshot::from_source_selected` had separate direct
children for raw worksheet parsing (59.76–60.10% of planning Ir) and
`validation::validate_xml` (37.13–37.86%). After the change, the candidate
shared `worksheet_xml_and_parse_source` boundary accounts for 95.63–96.17%
of planning Ir. Its inclusive descendants overlap and must not be summed.

| Candidate owner | Inclusive share across pairs | Self Ir, medium / dense-sparse |
| --- | ---: | ---: |
| `cell_values::validation::worksheet_xml_and_parse_source` | 95.63–96.17% | 6,529,928 / 12,582,872 |
| `raw::worksheet::…Parser>::transition` | 36.29–37.15% | 3,029,332 / 5,838,772 |
| `raw::worksheet::…Parser>::start` | 23.34–23.98% | 1,313,940 / 2,531,108 |
| `quick_xml::reader::Reader::read_event_impl` | 14.57–14.94% | 4,882,593 / 9,398,081 |
| `raw::worksheet::…Parser>::start_cell` | 14.43–15.12% | 2,027,520 / 3,914,240 |
| `cell_values::validation::validate_element` | 9.63–9.91% | 1,696,675 / 3,267,699 |
| `raw::worksheet::…Parser>::finish_parse` | 8.30–8.41% | 645,756 / 1,246,076 |
| `raw::worksheet::semantic::materialize` | 6.12–6.17% | 1,428,480 / 2,757,760 |

The shared traversal is therefore the remaining planning boundary. Within it,
`Parser::transition` and its `start`/`start_cell` work are the largest visible
residual paths. The candidate `Reader::read_event_impl` inclusive share is
14.57–14.94%, down from the baseline 21.86–22.30%, while the new transition
path carries 36.29–37.15%. This is consistent with the measured traversal
consolidation, but the nested inclusive values describe attribution rather
than independent costs.

## Native phase context and next priority

The native primary rows in [`comparison.json`](comparison.json) show planning
p50 reductions of 28.50–31.60% and whole-workflow p50 reductions of
5.89–11.91%. The reported phase shares exclude the separately reported reopen
phase and sum the open, planning, commit, and publication phases:

| Phase | Baseline share | Candidate share |
| --- | ---: | ---: |
| open | 0.24–0.41% | 0.31–0.52% |
| planning | 32.62–32.96% | 24.46–26.13% |
| commit | 34.98–35.61% | 38.52–39.48% |
| publication | 31.15–32.13% | 33.99–36.70% |

For the overall OOXML workflow, the next priority should be a separately
profiled commit path: commit is now the largest candidate phase, while its
primary p50 change is mixed (−1.63% to +3.72%). Publication is the next
largest phase and is also mixed (−5.00% to +6.73%). The current planning
owner profile cannot attribute either phase, so an exact commit/publication
owner should be frozen and measured before changing those paths.

For a planning-only follow-up, the bounded target is
`worksheet_xml_and_parse_source`, starting with `Parser::transition` and its
`start`/`start_cell` descendants. Any follow-up must retain the validation,
source eligibility, x14ac, error-order, allocation, and refusal-guard checks.
No production optimization or latency conclusion follows from this review;
the native phase measurements and the Ir attribution remain separate pieces
of evidence. The profile covers only the two frozen shapes and two repeats,
so it makes no cold-cache, range, scaling, or Office-producer claim.
