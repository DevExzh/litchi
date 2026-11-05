# 0360ac — Preserve catalogue sharing while repairing cold reads

The W4 cold-reader prototype removes the measured three extra allocations,
reducing its added bytes from 1,016 to 8 per operation on the reference probe.
That cold-path result alone does not qualify retained edit history.

Source review of the working prototype based on `233e605202` finds that
`SourceInner::defined_names` changes from an Arc slice to a boxed slice.
`open_internal_with_catalog` still validates the newly parsed names against
the shared proof, but now retains that new boxed catalogue in each reopened
owner instead of reusing `shared.defined_names`. Meanwhile `catalog_proof()`
creates an Arc from a clone of the ordinary owner's boxed names. History
owners therefore retain duplicate strings even while their proof Arc is
shared. The existing proof-identity assertion alone cannot detect this.

The prototype is returned for repair before its source commit. Preserve the
ordinary cold-open improvement and the shared catalogue on proof-backed
reopens. Verify the actual public names/string backing storage across
repeated owners, then remeasure retained history and first admission as well
as the cold routes. Include managed and customized policy opens: the prototype
also adds an Arc around the execution context and separate storage for
nondefault semantic/owner options. Source-sheet and cache sharing also need
to remain intact.
No runtime or retained-heap measurement is claimed by this source review.

The reviewed source diff is retained outside the checkout at
`.coord/w4-xlsx/cold-reader-review-before.diff` (21,084 bytes), SHA-256
`4ebad25abd52d818c70bee60cc49f6f6e715a72e592d1050d0a079c4e1a0cacc`.
It was captured with `git diff -- crates/litchi-xlsx/src/workbook/source.rs`
in the worker tree and can be applied to the recorded base for inspection.
The final repair, its remaining byte/peak differences and native performance
qualification remain required. This record grants no performance exception.
