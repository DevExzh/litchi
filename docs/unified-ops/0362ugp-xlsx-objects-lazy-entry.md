# 0362ugp — XLSX object-family lazy edit entry

This correction closes the XLSX rows in the 0362 edit-construction audit for
data-type icons, worksheet OLE objects, page breaks, workbook revisions,
rich-value refresh intervals, and Survey metadata.

Each ordinary `Package::edit_*` entry now retains only the immutable record
source, a bounded worksheet selector where needed, and the explicit edit
policy.  The source snapshot and typed clone-staged draft are prepared on the
first accessor, mutator, or commit.  A failed preparation leaves the edit
unprepared, so a later operation retries without retaining partial state.
Fallible accessors (`before`, typed projections and Survey entries) preserve
the existing family error instead of returning a dummy value or panicking.

The existing source-closure checks, signature policy, exact no-op commits,
replay and inverse patches remain on the same materialized paths.  Snapshot
and patch loaders are unchanged: read-only projections and committed patches
still expose their established infallible semantic accessors.  The source
edit consumers and the XLSX operations page-break writer were migrated to the
fallible staged APIs in the same change.

Cargo validation is intentionally pending coordinator allocation of the XLSX
build slot.  Before assembly, run the pinned offline locked all-target checks,
the API glossary/debt gate, and the repository artifact guards.
