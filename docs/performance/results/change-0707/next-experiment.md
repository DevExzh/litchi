# Prospective validator name-storage experiment

This is a design handoff, not an implemented change or an admitted performance
claim. Freeze a new paired measurement plan before coding or comparing results.

The current validator copies each Start local name into a boxed byte slice and
copies each Empty local name into a temporary Vec. Its stack retains only names
needed for parent and exact closing-name checks. A private representation could
retain modeled names statically, retain unknown names in owned byte slices, and
borrow Empty local names only for the duration of validation. No event borrow
may outlive the callback. Classifying a name for storage must not change semantic
admission, namespace resolution, or copied-subtree handling.

Preserve all validation and error ordering, strict/transitional dialect checks,
unbound-prefix refusals, exact closing-name bytes, XML depth limits, allocation
failure handling, opaque bytes, and authoritative fallback. Keep existing
try_reserve checks. A static/owned enum may increase each stack frame from 16
to 24 bytes on this host; measure the actual layout and bounded-depth effect.
Avoid mutable global caches, new public APIs, parser fusion, and proof reuse.

Keep validation_borrow_tests.rs's frozen owned reference independent, including
its Vec<Box<[u8]>> stack. Compare acceptance and complete error strings. Add:

- All modeled worksheet and workbook names, including Start/End and Empty
  forms and rich text is/r/t, under transitional/strict and prefixed/default
  namespaces.
- Misplaced v under row, t under c, c under sheetData, and sheet under workbook,
  preserving the exact contextual refusal.
- A copied unknown subtree with modeled-looking local names in a foreign
  namespace, followed by a real sheetData sibling.
- Unknown empty elements and namespace rebinding inside a copied subtree.
- Copied-subtree close mismatches against the owned reference, without assuming
  the XML reader versus validator supplies the diagnostic.
- Both orders of malformed XML and an invalid composed-span element, preserving
  first-error precedence and shared-traversal fallback.

Run matching fixtures through shared_traversal_tests.rs's admitted/authoritative
comparison. Retain existing depth, source retry, typed refusal, and atomicity
coverage. Focused targets are the borrow_tests, shared_traversal_tests and
facts_oracle_tests library filters and the source_backed_cell_values integration
suite; full XLSX and consumer quality checks are required if a candidate survives.

Use a fresh CPU-pinned native A/A and A/B/B/A matrix for both primary shapes,
plus one-edit, managed-budget, noncompact, vendor-extension, and producer guards.
Separate allocator and planning instruction captures. Require meaningful native
end-to-end gains in addition to allocation reduction, review every greater-than-5%
regression and repeat-drift flag, and preserve exact output/source/resource parity.
Callgrind replaces allocator behavior, so allocator instruction costs cannot
predict native latency. Do not sum nested inclusive instruction costs or phase
allocation peaks. RSS and bounded-stack memory costs need separate observations.
