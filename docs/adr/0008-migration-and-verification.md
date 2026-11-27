# ADR 0008: Migration and verification

- Status: Accepted
- Date: 2026-07-31

## Migration policy

The public API migration is breaking. Format-specific crates replace the
legacy monoliths. The shared storage, drawing, text, and spreadsheet crates
own the vocabulary their format crates depend on. XLSB shares spreadsheet
models with XLSX through those crates. There are no released compatibility
shims.

## Support criteria

The [CRUD scenario checklist](../CRUD_Scenario_Checklist.md) defines the
operations covered by a support claim. A supported direction covers the
applicable checklist rows. An inapplicable row has a format-specific reason.
Native Office and iWork compatibility includes opening and editing a
document in its native application and reading the saved result. Expected
producer metadata changes are allowed. Repairs, missing content, and
unexplained semantic changes fail the compatibility contract.

The unified Rust API and its Python bindings share the following contracts.

- An atomic batch failure identifies the failing item and preserves the
  source snapshot and save destinations.
- Applying a patch to its source reproduces the edit. Applying its inverse
  restores the source bytes. Applying it to an incompatible source fails.
- Independent edits made against one snapshot can compose without losing
  either change. Each individual patch checks its source before publication.

Feature matrices communicate support by direction. Generated results belong
outside the repository.
