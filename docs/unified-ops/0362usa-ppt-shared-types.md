# 0362usa — PPT/PPTX shared edit types

This source slice completes record 0351 §2.2 for the remaining minimal
PowerPoint edit families. It is based on `173dd5330476713714871007316794b9a3a1bdaa`.

`litchi_ppt::package::metadata` now re-exports
`litchi_core::edit::Diagnostics` and aliases its exact triad as
`Commit<Package, Patch>`. Metadata publication uses `Commit::new` and the
shared no-op/changed constructors; `into_parts()` keeps the core `(Package,
Patch)` contract. Existing property-set errors, exact source patches, and
no-op package reuse are unchanged.

PPTX's changed-only diagnostics module and `opened::Diagnostics` both
re-export the core type. `opened::Commit` remains a native type because it
retains its `Arc<Package>` and opened `Snapshot` for replay and retained
projection accounting; its triad accessors and genuine no-op behavior are
unchanged. All affected constructors and consumers use the shared
diagnostics value without an additional scan or allocation.

This lane was static-only while the coordinator's capped Cargo slots were
occupied; the coordinator's integrated build and API gate provide validation.
