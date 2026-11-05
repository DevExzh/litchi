# Combined PPTX slide-property publication

A unified update containing both slide name and visibility now stages both in
one native operation for eager and source-backed owners. The source primary
transaction exposes `Edit::set_slide_properties(slide, name, hidden)`, matching
the existing eager and focused-source method. Individual name/hidden setters
delegate to it. Omitted properties remain unchanged; an empty update is a no-op.

This removes the duplicate lexical scan, reservation and staging that remained
in 0162. The durable grammar and publication boundary are unchanged. A new
differential test compares combined and sequential edits, exact eager/source
output, equal-value no-ops, failure after earlier successful staging, transported
replay and inverse restoration.

## Validation

Nine native slide-property tests, five unified slide-property tests and nine
runnable API examples passed. Clippy passed with warnings denied for the native
and facade libraries and affected integration tests, and for all harness
targets. Scoped formatting and diff checks passed. The full-feature PPTX glossary
check passed all 70 bans with rustdoc warnings denied. It used a temporary
lockfile selecting cached cc 1.6.0, as in 0162; the inherited workspace
lockfile was unchanged. This is focused validation
of the combined path; the complete 0162 suites were not rerun for this change.

## Paired measurements

The harness now uses a single native combined setter and one commit. Its JSON
schema is version 2 and its routes end in `slide_properties_v2`, distinguishing
this reference path from the two-setter sequence measured in 0162.

[Results](0163-pptx-combined-slide-properties-results.json) retain every sample,
corpus hash, source digest and binary provenance. Three ABBA rounds give six
samples per leg, pinned to CPU 0 with no concurrent Cargo job. All ten cases
produced exactly equal native/unified bytes, and the failure matrix is empty.

| Shapes in slide | Eager unified/native | Source unified/native |
|---:|---:|---:|
| 1 | 1.002 | 1.029 |
| 10 | 1.004 | 1.022 |
| 100 | 1.012 | 0.987 |
| 1,000 | 1.000 | 0.984 |
| 10,000 | 1.000 | 0.985 |

One-shape native/unified medians were 301.8/302.3 microseconds for eager owners
and 46.5/47.9 microseconds for source owners. Ratios below one are measurement
variation. These warmed generated-deck measurements exclude opening,
serialization, allocations/RSS and real-world corpus acceptance. They do not
establish a controlled before/after speedup against the separate 0162 run.

The broader all-format CRUD, Python, ADR migration and performance goal remains
active.


Temporary builds, raw logs and measurement scratch were removed after retaining
all timing samples and the binary name, size, SHA-256 and regeneration command
in the reduced results. No worktrees or borrowed fixtures were created.
The rewritten base remains an ancestor of this branch.
