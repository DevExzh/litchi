# Independent scanner attribution review

The delegated profiler independently reparsed all eight measured 0550 raw
dumps with the retained `parse_raw` helper, without modifying files. Root then
compared the returned costs with `analysis.json`. All eight scanner totals,
self costs, and named child costs agree. Each has one positive rewrite caller;
scanner self plus aggregated direct children exactly equals incoming scan Ir.

The source-backed owner is `MultiSourceEdit::commit`, under the measured
`run_xlsx_cell_values_edit_save` parent. This does not profile the distinct
single-sheet `SourceEdit` API. It also excludes external staging, publication
and final commit destruction. Old eager API profiles are not substituted.

The review supports investigating already-produced addresses and compact
offsets. A full scanner observer would still perform the separately emitted
`Scanner::start_cell` work: its `cell_address` and `wire::cell_tag` children
dominate that function. This nested qualitative finding was independently
reproduced directly from the sealed raw dumps; the canonical `analysis.json`
records scanner-level partitions only, not this additional nested partition.
Names are codegen evidence; an absent symbol is not
proof that its source-level work did not execute.

The report's category names deliberately retain their narrow scope. Named
reader plus namespace dispatch includes only three emitted functions. QName,
namespace-level changes, byte decoding, allocation and any inlined work are
not silently assigned to that category. Disjoint direct scanner partitions
must not be added to overlapping nested parent/child totals.

Disposition: attribution and design selection only. No allocation-count,
retention-size, removable-work, latency, hardware-counter or candidate
admission claim follows. All repeated native/allocator drift analysis remains
in the sealed 0550 evidence; this batch introduces no new measurements.
