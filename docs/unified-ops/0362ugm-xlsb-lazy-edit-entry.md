# 0362ugm — XLSB lazy edit entry

This slice closes the XLSB edit-construction cost debt.  Package-owned edits
retain only an immutable package owner and their finite policy at entry; source
catalogs, worksheet drafts, native owner snapshots, OPC graph images and exact
before-images are prepared on the first operation or commit that needs them.
Dropping an untouched edit therefore does not scan or copy package payloads.

The deferred paths cover the structural workbook and cell-values families,
generic `part_edit::Staged` families, Calculation Chain, Data Model, Volatile
Dependencies, Custom XML Maps, connections, shared-workbook metadata, Theme,
Custom Data, and threaded comments.  Fallible accessors retain the existing
typed errors where a read-only staged projection must materialize a source
snapshot; no dummy values or panic-based lazy state are introduced.

The existing replay, exact no-op, inverse, signature-policy, bounded-limit and
publication behavior remains on the same first-use/commit paths.  Constructor
and drop tests prove that a cold package projection remains cold for the
structural and generic staged owners.

Cargo validation is intentionally pending coordinator allocation of the XLSB
build slot.  Static `rustfmt`, `git diff --check`, and repository artifact
guards are required before the implementation commit.
