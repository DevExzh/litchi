# 0362usb — XLSB and XLS shared commit triads

- Date: 2026-10-23.
- Status: static migration complete on frozen base `173dd5330476713714871007316794b9a3a1bdaa`; coordinator integration validation pending.
- Scope: the changed-only XLSB cell-values and XLS metadata edit families required by record 0351 §2.2.

## Shared types

The XLSB cell-values package and structural families now expose the canonical
triad types:

```text
pub type Commit = litchi_core::edit::Commit<Package, Patch>;
pub use litchi_core::edit::Diagnostics;
```

The XLS `package::metadata` family uses the same aliases. Every constructor
and package replay consumer in these families now calls the shared
`Commit::new`; no local `Commit`, `Diagnostics`, or publication helper remains.
The core `Commit::into_parts` contract intentionally returns `(Package, Patch)`;
the package, patch, and diagnostics accessors are consequently the shared
reference-based accessors.

The XLSB structural replay route in `package::mod` is migrated together with
its owner module so undo, redo, direct edits, and patch replay retain the same
triad. Tests and examples continue to use the public package/module APIs.

This migration is deliberately limited to changed-only diagnostics. The
richer XLS and XLSX OOXML metadata diagnostics that carry family and field
information remain their format-specific wrappers; they are not aliases to a
content-free shared diagnostic.

No Cargo command was run in this leaf because the coordinator reserved the
build slot. Static formatting, diff, and API searches are the leaf checks;
the coordinator should run the affected XLSB/XLS tests after composition.
