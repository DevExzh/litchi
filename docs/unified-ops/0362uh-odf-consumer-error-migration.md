# 0362uh — Migrate native ODF consumers to family errors

The assembled v6 consumer inventory at `a814baa1de` contained 85 diagnostics
in `litchi-odb`, `litchi-odp`, `litchi-ods` and `litchi-odt` after the family
package and facade verbs adopted `litchi_odf_common::Error`. This repair keeps
that boundary intact across every affected native test and internal consumer.

The ODF projection chains no longer convert an ODF error to
`litchi_core::Error` before calling the next ODF projection. Edit helpers and
typed assertions now use the owning family error. Core failures that remain
projected by the family boundary match `Error::Core(CoreError::...)`; ODF
limit and refusal assertions match `Error::Limit` and `Error::Refused` with
their existing dimensions and refusal details. The generic flat ODT tests
retain separate assertions for private core-returning helpers and public ODF
returning methods, preserving the mixed internal/public capability boundary.

No test was weakened to a generic `is_err()` assertion, and no error alias was
introduced. The change is source-only as required by the assembled-consumer
repair lane.

Validation used the pinned Rust 1.95 formatter on every changed Rust file and
`git diff --check`. Cargo builds and tests were intentionally not run under
the lane's no-build instruction.
