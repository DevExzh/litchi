# Follow-up candidates from the read-only goal audit

These are investigation priorities, not performance findings or completion
claims. They require separately declared protocols and compatible oracles.

1. **Matched source-backed media-rich PPTX lifecycle.** Reuse the owned
   lifecycle's source/destination archives and slide positions with the bounded
   source-backed cross-copy implementation in
   `crates/litchi-pptx/src/presentation/source_cross_copy.rs`. Establish equivalent
   ingress, planning, application/publication and sink boundaries. Verify
   dependency closure, semantic output, untouched destination records, source
   read ranges, allocations and RSS. Do not compare the existing plain
   source-backed phase timer with the media-rich owned lifecycle. Notes,
   diagrams, shared owners and extensions outside the admitted closure must
   retain their typed refusals.

2. **Changed and mixed OPC overlay publication.** Measure the existing
   `opc_source_overlay_multi_part_changed` and `..._mixed` selectors over the
   small/large/media-incompressible and 2/8/32-part matrix. Attribute work in
   `crates/litchi-opc/src/source_backed.rs`; retain no-op as an independent
   transport guard. Compare identical mode, shape, part count and publication
   timer boundaries. The 0402 accepted no-op cells do not establish changed
   publication performance.

3. **Matched XLSX scalar edit/save.** Use the existing eager/source-backed
   `xlsx_*_cell_values_one_edit_save` pair, then the matching one-percent pair,
   over their fixed four-sheet corpus. The owner is
   `crates/litchi-xlsx/src/cell_values/`, with OPC publication attribution.
   Add operation allocation and source-range evidence where needed. The
   eager-only 0410 MCE evidence and unrelated cell-removal selectors are not
   substitutes for this comparison.

All three remain opt-in investigations. The representative coverage index,
default matrix, broad producer/cold/remote/native/concurrent/scaling evidence,
and overall non-iWork definition of done remain separate open work.
