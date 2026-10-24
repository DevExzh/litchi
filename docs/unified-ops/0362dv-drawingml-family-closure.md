# DrawingML detached family closure

Record 0362dv closes the remaining public DrawingML rows in the 0351
contract. The change is intentionally confined to `litchi-drawingml` and the
consumers of its public chart, transform, theme-family, and ink-action APIs.

## Contract decisions

- `chart::Package`, `transform::Package`, and `theme::family::Package` are the
  immutable source owners. The old public `Snapshot` names are removed rather
  than aliased. Their raw `Chart`, `Transform`, and `Family` projections remain
  distinct from the owner.
- Each owner exposes `edit()` and source-checked `apply(&Patch) -> Result<Commit>`.
  The commit is `litchi_core::edit::Commit<Package, Patch>` with shared
  content-free `Diagnostics`; publication uses `package()`/`into_package()`.
  `Patch::is_noop()` and `Diagnostics::changed()` replace public
  `is_changed()`/snapshot accessors while retaining exact no-op bytes and
  reversible patches.
- `chart::Transaction` and `transform::Transaction` become focused `Edit`
  types. Theme-family `Edit::is_changed()` is deleted. The old commit
  `snapshot()`/`into_snapshot()` methods are deleted without compatibility
  aliases.
- Raw canonical ink authoring now requires `AuthoringLimits` at
  `ink::Prepared::from_bytes`. Source-backed action edits require `Limits` at
  `Edit::{new,from_profile}`; `Edit::with_limits` is deleted without an alias.
  Its commit uses the same core triad with `Prepared` as the checked package.

All chart, transform, theme-family, and ink-action tests, examples, DOCX/PPTX
consumers, and durable decoder call sites were migrated in the commits that
change these public signatures. The API glossary profile registers the six
family module paths and pins each shared commit/diagnostics type to the core
triad.

## Validation

The focused DrawingML test target is checked with the repository capped build
shim, one job, offline locked dependencies, and a target outside the checkout.
The dependent DOCX, PPTX, and operations test targets are checked with the
same settings. No generated binaries or large evidence files are part of this
record.
