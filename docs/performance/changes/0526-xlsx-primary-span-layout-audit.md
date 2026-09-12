# 0526: investigate XLSX primary-span layout

The accepted 0525 profiles place 53.26% of measured commit instructions in
worksheet layout scanning. This batch follows that owner with a source-bound
cost decomposition and a draft row-owned primary-span arena. Production is
unchanged and no new latency, allocation or RSS improvement is claimed.

The current scanner builds a separate span vector for each cell, then shrinks
it into a boxed slice. Both the ordinary row writer and the provenance writer
consume those spans only for payload replacement. The draft retains each span
and its order in a row-owned array, with a range per cell. It does not combine
XML parsing and rewriting, skip validation, or revive the rejected 0522
cell-reference/tag scanner.

The design must preserve arbitrary multiple payload spans and unknown markup
between them, empty rows/cells, style-only copies, formula metadata, inferred
addresses, source/version checks and typed failures. Allocation growth and
memory peaks may change, so fewer individual containers are not sufficient
proof of improvement. The retained allocator call graph cannot attribute every
allocator instruction uniquely to span storage.

The [evidence bundle](../results/change-0526/README.md) records current source
and ADR bindings, retained-profile decomposition, design and tester reviews,
and the unapplied candidate patch. A fresh baseline/candidate pilot is the next
step. Admission thresholds must be fixed before capture; useful end-to-end
results and correctness gates are required before production retention.

The existing native, profile and allocation measurements remain those of
accepted 0525. This audit adds no workload, producer, physical-provider,
cold/range, scaling or broad CRUD certification. OLE2/OOXML stays first, ODF
remains deferred, and iWork is excluded. The full performance goal remains open.
