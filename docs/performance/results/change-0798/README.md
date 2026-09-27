# 0798 — OPC attribute consumption in public PPTX workflows

This diagnostic follows the 0797 first-attribute replay preflight. Before choosing
another specialization, it records how many attributes the OPC checked iterator
sees and how much callers consume within representative operation regions.
Production is restored exactly before all captures, and no optimization is adopted.

The probe preserves the 0793 corpus and exact public operation boundaries:
`Package::opened_presentation` for capture; `Transaction::commit` after staging
for commit; and capture/edit/commit/publication/serialization for lifecycle.
Package ingress, fixture generation, and semantic/output verification stay outside
the region. Five shapes cover tiny, medium, large, and ASCII/Unicode vendor
attribute fixtures. Historical source/output identities are compared without
pooling timing.

Fifteen fresh plain controls use one sample and no warmup. Two fresh instrumented
repeats cover the same 15 cases in forward then reverse order, also one sample
without warmup. Root runs all builds and 45 captures serially, pinned to CPU 12.
The baseline and temporary instrumented source are separately compiled and copied;
production restoration precedes both capture lanes.

The diagnostic instruments only `litchi-opc::xml_attributes::CheckedAttributes`,
only on the calling thread, and only within an explicitly enabled census region.
It records iterator instance lifetimes, successful/error/end observations, full
lexical attribute-prefix counts, and lossless element-name bytes. Full-input scans
at drop and histogram allocations perturb work substantially. All recorded elapsed
values are diagnostic leftovers from the inherited probe and are excluded from
performance comparisons. This is not a latency, allocation-count, RSS, instruction,
or speedup measurement. Other helper owners, skipped iterators, and other threads
are outside the census.

Correctness requires exact fixture/source/output/readback parity with the plain
controls and sealed 0794 identities. Census qualification requires no overflow,
zero live instances at finish, instance/count conservation, and exact agreement
between both repeats. Partial consumption, errors, clones, and zero-consumption
instances remain explicit rather than being discarded. Count distributions may
inform the next candidate but cannot adopt production code or predict workflow
speedup by multiplying microbenchmark ratios.

The packet retains the exact diagnostic source, build/test receipts, raw reports,
independent audits, reviews, restoration witness, and cleanup witness. Full
recapture requires a new packet/target because drivers refuse overwrite. Only
the owned temporary target is removed. Unrelated files and worktrees are preserved.
OLE2/OOXML remain active; ODF is deferred and iWork is excluded.
