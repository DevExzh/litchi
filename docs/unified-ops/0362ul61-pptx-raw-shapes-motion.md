# 0362ul61 — PPTX raw shape and motion limits

This leaf carries the PPTX shape, transition, and laser raw-boundary
migration from integration base `a942320908`.

## Public boundary

The borrowed and owned shape readers now expose only `read(..., Limits)`;
the old `read_with` and default twins are gone.  Transition `read`, `write`,
and `write_to` likewise require the transition policy.  Transition output has
its own `output_bytes` budget, separate from opaque-reader retention.

Laser raw operations require a `Limits` policy.  `Usage` is a separate mutable
aggregate ledger passed through slide walks, preserving total input, trace,
and point accounting across slides without mutating the policy.  Failed reads
stage their charges and leave the caller's ledger unchanged.  Laser fragment
writers admit output incrementally through the selected output bound and
restore an existing destination buffer on refusal.

The package editor intersects the retained `OpcPackage::read_limits()` policy
with the laser safety ceilings once, then threads that profile through input
admission, dialect scanning, fragment writing, complete-slide size preflight,
and read-back.  The insertion path computes the complete projected slide size
before allocating its output buffer; it does not fall back to the codec
default when an owner was opened under a tighter profile.

All PPTX source, package, edit, unit-test, integration-test, and example
consumers were migrated to the canonical signatures.  The minimal and full
PPTX glossary profiles register the required shape, transition, laser, and
package-loader signatures.

## Source checks

The source-only checks for this leaf are JSON parsing and static API/call-site
audits.  Cargo, rustdoc, native tools, and performance capture were not run in
this wave; the parent coordinator must perform independent compilation and
validation after composing the neighboring raw-content and raw-properties
leaves.

Focused source tests cover exact and one-byte-under output boundaries, failure
atomicity, aggregate laser usage across multiple slides, and a retained-owner
profile that refuses a complete-slide growth before allocation.  The main
performance consideration is the bounded laser and transition writers' small
incremental sinks; they preserve direct appends and avoid growing a destination
beyond the selected output ceiling.
