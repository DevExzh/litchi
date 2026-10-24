# 0362dvh — DrawingML geometry, color, transform and formatcode2 raw limits

Date: 2026-10-24. Parent: 0362dv and 2f7db7ee7f.

This lane gives the remaining resource-scaled DrawingML fragment codecs one
required module-owned policy. The old static-only entries are removed from the
public surface; every consumer in the shared text, PPTX shape/metadata,
spreadsheet-drawing, XLSB host, focused tests, and tracked report harnesses
passes an explicit default policy.

The color policy bounds XML bytes, depth, nodes and transforms. Transform
`Limits` bounds XML bytes and nodes and is retained by the detached transform
package through edits, commits, replay and inverse replay. Custom geometry
reader and writer APIs use `geometry::Limits`; builders no longer have a
public `Default` escape, and geometry serialization checks its output budget
before appending to the caller's buffer. `formatcode2` exposes one `Limits`
policy for element and attribute reads, shared-source reads, value writes and
sink writes; input, decoded value and output budgets are checked before
allocation or source publication.

The migration preserves existing hard caps and source fidelity. Semantic
color/transform/geometry model constructors remain domain APIs. Historical
measurement snapshots and receipts under `docs/report/spec-gap-validation-evidence`
are unchanged; their live Cargo harnesses are migrated in place and receive
the explicit policy at every codec call.

`cargo fmt --all -- --check` and `git diff --check` pass. With the coordinator's
capped offline target, `cargo check --offline --all-targets -p
litchi-drawingml -p litchi-spreadsheet-drawing -p litchi-pptx -p litchi-xlsb`
passes. The focused DrawingML library plus formatcode2 and transform tests
pass (238 library, 12 element, 17 adversarial, 4 XML-misc and 1 facade test),
and the spreadsheet-drawing shape tests pass (24 tests). No native workload
was run; the full workspace gate, fresh inventory and native qualification
remain pending assembled-source validation.
