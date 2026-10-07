# 0963 reduced-readback fusion review

Status: independent, read-only review. No Rust source, build, benchmark, or
workload was run. This note records the route census, the arithmetic check for
the 0963 residual, and the implementation disposition.

## Route and removable-work census

The current route in `cell_values/snapshot.rs` validates the complete rewritten
worksheet, then `try_rewritten_value_cells` in the same file constructs a
reduced byte buffer, parses it with `Repairs::Refuse`, and calls
`Store::merge_omitted_cells` (`raw/worksheet/edit/package.rs` and `cell.rs`).
The validator/parser source route is already fused in
`cell_values/validation.rs::worksheet_xml_and_parse_source`; that fact does
not make the rewritten reduced route fused.

The 0958 source census gives the admitted measured-route counts:

| case | owners | updates | omitted owners | merge calls/iteration |
| --- | ---: | ---: | ---: | ---: |
| medium 1% | 9,216 | 93 | 9,123 | 4 |
| dense-sparse 1% | 17,792 | 178 | 17,614 | 4 |
| medium one edit | 2,304 | 1 | 2,303 | 1 |

The retained 0958 captures independently reproduce the omitted counts as
`Stored::clone` calls over three samples: 27,369, 52,842, and 6,909. Their
direct merge callsites are 12, 12, and 3 calls, respectively. The exact source
numeric lexical payloads for those omitted owners are only 58,723, 92,907, and
10,933 bytes. The selected allocation report gives 8.28--8.43% of selected
allocation calls but only 0.28--0.38% of selected requested bytes to the
numeric clone subset. These are allocation witnesses, not CPU or elapsed-time
shares.

The direct `merge_omitted_cells` reserve at the retained 0958 callsite requests
5,087,232, 9,821,184, and 1,271,808 bytes over the three-sample captures,
or 1,695,744, 3,273,728, and 423,936 bytes per iteration. A fused parser could
change how the final store is assembled, but an independent immutable store
still needs cells, indexes, and metadata. These requested bytes therefore are
not a removable-work estimate.

## Independent arithmetic check

The replay helper `reduce_residual.py` was run against the committed 0962
boundary packet. Its input is 9,820 bytes with SHA-256
`14fa5e02723da074f0cf06c494f2ff48e1389bb489e2f189fb09ad6ce08ea477`.
For every child, the checked inequality is
`nested_validator <= rewritten_readback <= selected_stage`, followed by
subtraction. The independent result matches `readback-residual.json` exactly:

| quantity | samples | period |
| --- | ---: | ---: |
| rewritten readback | 4,319 | 18,905,129,394 |
| nested validator | 3,509 | 15,354,265,106 |
| residual | 810 | 3,550,864,288 |
| selected denominator | 12,319 | 52,852,646,778 |

The residual fraction is
`3,550,864,288 / 52,852,646,778 = 0.0671842283114959`, or 6.7184%.
Child fractions range from 4.2524% to 7.2645%, and child residuals conserve
the pooled sample and period totals. **PASS: arithmetic and conservation.**

This 6.7184% is a sampled ancestry remainder after subtracting the nested
validator. It still includes reduced parsing, omission conversion, merge,
drop, and other retained work. It is not a removable-cost bound, a timing
measurement, or an end-to-end speedup prediction. The 0962 deepest-inline-source
attribution leaves 55.13--64.09% of qualified period source-unresolved, so it
cannot turn this remainder into a complete owner ranking.

## Disposition

Decline implementation in this batch. Keeping full validation means a shared
rewritten reader must still resolve and visit every event for the authoritative
validator. It can skip semantic materialization only for exact omitted-cell
subtrees; it must add event-span/range checks and preserve row/duplicate/
collision state. `SourceAdmission::Borrowed` is the safe first gate because
MCE rewriting can shift byte offsets. Provisional parser failure, MCE or
unsupported metadata, stale provenance, allocation failure, cancellation, and
typed error precedence must retain the current reduced-parse/merge and complete
parse fallback.

The source evidence does not isolate enough removable CPU work to pay for that
new contract. The 0602 derived real-producer projections reported small
`reduced_readback` and merge instruction fractions, but those operations,
source revision, and guest-instruction population are different; they are
context only and are not substituted for the current 0962 measurement.
This is a prioritization decision before coding, not a compiled-candidate
regression claim.

The next single measured target should stay on current source admission: isolate
the repeated namespace/event-resolution work in the source-backed traversal,
with explicit native and allocator gates. The 0962 physical table identifies
`NsReader::process_event` and `NamespaceResolver::resolve_event` as recurring
self symbols, but the unresolved-source fraction means this is a target for a
new matched measurement, not an adopted optimization. Older real-producer MCE
figures remain a separate corpus/design question and are not current-route
qualification.

Evidence references: `docs/performance/0962-xlsx-native-cpu-attribution.md`,
`docs/performance/results/change-0962/source-fusion-caveat-v1.md`,
`docs/performance/results/change-0962/source-followup-v1.md`,
`docs/performance/0958-xlsx-edit-commit-allocation-attribution.md`,
`docs/performance/results/change-0958/review/number-ownership-trace-v2.md`,
and `docs/performance/0602-xlsx-real-producer-admission-design.md`.
