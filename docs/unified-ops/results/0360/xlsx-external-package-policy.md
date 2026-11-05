# 0360aea (return to 0360ae): XLSX external source-package policy handoff

This is the compact continuation of the main 0360ae receipt; the source-only
policy handoff and its deferred-publication fence are recorded here so the
coordinator can review the return without importing large evidence artifacts.

Status: source-only repair committed for assembled validation. This record does
not claim native timing or performance acceptance.

The source-backed XLSX facade previously reconstructed a committed owner from
`OpenOptions` alone. `SourceBackedPackage::cache_diagnostics().budget_managed`
could identify a managed package, but the package's actual cache limits and
execution context were not retained in `SourceOwnerOptions`. An externally
opened package could therefore lose its one-entry cache policy or become
unmanaged after a deferred commit was materialized.

The repair adds the package-owned `SourceBackedPackage::cache_limits()`
accessor and carries package policy separately from an optional facade context.
`from_source_backed_package_with_options` keeps the package's cache limits and
package execution context authoritative. A supplied facade context remains an
additional fence. Deferred materialization chooses cache limits and package
execution context from the shared owner policy, so facade options cannot
replace them. Sequential reader opens reuse one shared owner-policy handle to
avoid an additional context allocation.

Focused regressions added to `workbook/source.rs` cover:

- an externally opened one-entry cache surviving a changed commit and both
  worksheet reads, even when facade options specify the default cache;
- an externally managed package retaining cancellation after commit,
  materialization and a successful read; and
- separate package and facade contexts both remaining active after commit and
  materialization.

The deferred publication path now checks the source revision first and then
both execution fences before allocating a reopened artifact or writing to a
caller sink. The focused cancellation regression leaves an uncached committed
owner unmaterialized and verifies both a deferred sheet read and `write_to`
return cancellation without touching the output sink. This preserves the
existing source-changed-before-cancellation precedence; the package-owned
publisher still enforces its own context after the facade fence.

Static validation completed before this handoff:

- `rustfmt --edition 2024 --check` on the two changed Rust files;
- `git diff --check`.

The pre-publication fence and regression were reviewed with direct source
inspection; Cargo validation remains pending coordinator release. No native
timing or performance acceptance is claimed.

Cargo tests and assembled integration validation remain pending coordinator
release. No benchmark binaries or large evidence are part of this record.
