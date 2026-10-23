# 0362usc — Pages shared Commit and Diagnostics types

- Date: 2026-10-23.
- Base: `173dd5330476713714871007316794b9a3a1bdaa`.
- Status: source migration complete; coordinator validation is pending.

## Pages migration

`litchi_pages::metadata` now re-exports the shared
`litchi_core::edit::Diagnostics` and aliases its exact triad as
`litchi_core::edit::Commit<Package, Patch>`. Both `Edit::commit` and
`Package::apply_metadata` return that triad; the edit path uses
`Diagnostics::UNCHANGED`/`CHANGED`, while replay derives the same state from
the patch. The package-creation consumer now proves changed, no-op and replay
triads through the canonical `package`, `patch`, and
`diagnostics().changed()` accessors. No scan, allocation, or capability path
changed.

## Audit boundary

The audit was performed against the frozen composed source and 0351 §2.2.
Changed-only duplicate diagnostics/plain triads found elsewhere are:

* XLSB cell-values package and structural families;
* PPT, XLS, ODI, ODC, OTH and ODM metadata/facade families;
* the PPTX shared diagnostics helper and opened presentation diagnostics.

Those families are assigned to their format owners. ODI flat is also a plain
triad wrapper, although it already uses core Diagnostics; its derived
`semantic_patch` helper must move with that owner migration.

The following are not false positives: OOXML properties diagnostics retain
`families` and `fields` (`litchi-ooxml-common/src/properties/transaction.rs`),
and ODF metadata diagnostics retain a field count. DOCX, iWork, PPT/PPTX,
RTF, XLS/XLSX/XLSB and ODF families with operation counts, touched components,
parts, domains, fingerprints, or publication state likewise have richer
family state and may keep custom diagnostics/commit types under §2.2.
Generic CFB/OLE substrate transactions remain outside this format-facing
audit unless publicly leaked by a format facade.

Validation in this leaf is source-only (`git diff --check` and static source
inspection); Cargo, rustdoc, performance, and native qualification results
are intentionally not claimed.
