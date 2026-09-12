# 0530: attribute XLSX planning and identify the MCE namespace scan

Four fresh profiles of the unchanged XLSX edit_sheets method identify raw
worksheet parsing and full value-only validation as the dominant planning
costs. The batch changes no production or harness source and claims no native
speedup. It selects an exact substring-search optimization for a future pilot.

The previous 0529 batch rejected the XML attribute probe under its native
gates and retained instrumentation/tests. Its final source is byte-identical
to this baseline, so all 14 quality checks and 2,024 successful executions
remain bound. A fresh release build and four output/lifecycle profile oracles
provide the new execution evidence.

Retained 0529 baseline phase vectors show planning at 31.815642% of combined
edit/save time across four 200-sample primary rows. This uses summed phases
and excludes reopen. It motivates attribution but does not convert instruction
shares into native speedups.

The frozen protocol preferred SourceBackedEditor::edit_sheets and allowed a
snapshot-loader fallback only if that symbol was absent. Full nm output shows
two preferred-owner monomorphizations, so the method boundary was captured.
Each profile retains three lifecycle calls and one uniquely proven measured
call from run_xlsx_cell_values_edit_save, followed by a zero-Ir termination
dump. The method includes source loading and edit initialization; selector
construction is outside the native planning timer.

| Parent-specific cost | Aggregate Ir | Share of planning |
| --- | ---: | ---: |
| Planning method | 725,477,066 | 100% |
| Raw worksheet parsing | 434,424,105 | 59.881163% |
| Value-only XML validation | 272,706,313 | 37.589929% |
| Selected worksheet PartView::data | 15,467,269 | 2.132014% |
| Catalog loading | 1,814,338 | 0.250089% |
| MCE preprocessing, nested inside worksheet parsing | 37,572,982 | 5.179072% |

Parsing, validation and PartView::data are disjoint immediate children of
from_source_selected. Catalog loading is its sibling under the snapshot
loader. MCE preprocessing is nested in the parser and is not additive to it.
Instruction totals are 125,643,633 and 125,655,087 for medium, and 237,075,784
and 237,102,562 for dense-sparse. Direct/self equations and raw edge costs
replay exactly against deterministic annotations.

The catalog cost is too small to prioritize reuse for these cases. Source
inspection instead identifies the namespace-presence scan in
process_markup_compatibility: overlapping byte windows compare against the
MCE namespace. The existing memchr dependency offers an exact substring
search with the same predicate. A narrow replacement can retain input/output
limit ordering, borrowed no-MCE output, and the entire MCE parser unchanged.
Its native benefit across XLSX and shared DOCX/PPTX callers remains unmeasured;
the next pilot must establish it. Earlier rejected candidates remain rejected.

The evidence preserves all raw captures, caller scope, source/ADR/dependency
bindings and Valgrind warnings. It does not establish cold/range, scaling,
native-producer or broader CRUD performance. OLE2/OOXML remains the active
goal; ODF optimization is deferred until it completes, and iWork is excluded.

See the [evidence bundle](../results/change-0530/README.md),
[profile review](../results/change-0530/profile-review.md), and
[planning context](../results/change-0530/planning-context.md).
