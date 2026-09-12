# 0530 measured planning scope and attribution

All four fresh normal-release children pass output and lifecycle oracles.
The frozen symbol preference selects SourceBackedEditor::edit_sheets, which
has two exact demangled monomorphizations in the retained nm output. The
snapshot-loader fallback was not selected. Selector-vector construction stays
outside the native planning interval; the selected method includes the
execution check, source snapshot load and empty edit staging initialization.

Every profile has four positive numbered dumps. Raw incoming edges identify
one call from run_xlsx_cell_value_lifecycle_gates in each of parts 1, 2 and 3;
part 4 has exactly one positive edge and one call from the measured
run_xlsx_cell_values_edit_save. Part 4 is the unique final measured call.
Every termination dump has zero Ir. No dump count was assumed before capture.

The root adapter of the sealed 0528 analyzer regenerated both inclusive and
self annotations for each selected dump. It verified every retained direct
edge against raw costs and all selected-owner self-plus-direct equations,
then followed material descendants. The inherited Rust slice-symbol bracket
fix remains in the local adapter. All four analyses pass.

| Shape | Repeat | Planning Ir |
| --- | ---: | ---: |
| medium | 1 | 125,643,633 |
| dense-sparse | 1 | 237,075,784 |
| medium | 2 | 125,655,087 |
| dense-sparse | 2 | 237,102,562 |

Aggregate planning Ir is 725,477,066. Parent-specific edges from
Snapshot::from_source_selected assign 434,424,105 Ir (59.881163%) to raw
worksheet parsing, 272,706,313 Ir (37.589929%) to value-only XML validation,
and 15,467,269 Ir (2.132014%) to PartView::data. These three direct children
are disjoint. Catalog loading, a sibling of from_source_selected under the
loader, is only 1,814,338 Ir (0.250089%). This demotes catalog reuse for the
measured two-sheet workloads; many-sheet/repeated-edit workloads are separate.

The worksheet parser's direct process_ooxml edge is 37,572,982 Ir,
5.179072% of planning, nested within the 59.881163% parser share. It must
not be added to that parser total. Source inspection identifies a naive
windows(NAMESPACE.len()).any equality scan on the namespace-absent path.
An exact substring-search replacement is a concrete next measurement
candidate. Its effect on native workflows and other OOXML consumers is
unmeasured; the existing validation, limits, error ordering and fallback must
remain unchanged.

Historical phase shares are separate context. Neither planning Ir nor its
nested shares are native latency, allocation counts, or an end-to-end speedup.
Collection-off call metadata outside positively scoped edges may include
setup/readback; it cannot establish measured allocation activity. Valgrind
stderr warnings are retained, including the brk-segment limitation note.
No cold-cache, physical/range provider, scaling or native-producer claim follows.
