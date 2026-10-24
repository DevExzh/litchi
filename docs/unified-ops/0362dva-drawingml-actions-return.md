# DrawingML action publication return (0362dva)

This record returns the DrawingML leaf to the coordinator after review found
that the action editor still exposed two publication spellings. The source
state before this return was `9034328f32657bec6b144f607bfe71e5caf7d1f3`.

## Returned findings and repairs

- `ink::actions::Edit::commit` is now the sole source-backed publication
  method. `Edit::finish` and `Edit::prepare` are removed; every DrawingML
  action editor consumer uses `commit`, which always returns the
  `Commit<Prepared, Patch>` triad.
- Explicit `Alias for` compatibility spellings in the action family are
  removed (`from_bytes`, `with_value`, `add_*` builder aliases, `source`,
  `readback`, group-edit aliases, and the old publication aliases). Canonical
  builder methods remain under `ink::actions::builder`; InkML authoring is
  likewise discoverable under `ink::builder`.
- Before this return, `ink::actions::Prepared` had no canonical `edit` or
  `apply` owner route. It now has `edit()`, `edit_with_options(EditOptions)`,
  and `apply(&Patch) -> Result<Commit>`. Prepared edits share the immutable
  profile/source allocation through `Arc`; changed replay is read back before
  publication and no-op replay reuses the owner allocation.
- Chart, transform, and Theme Family raw `Patch::apply(&Package)` helpers are
  private implementation details. Consumers replay through the corresponding
  `Package::apply(&Patch) -> Result<Commit>` owner route, preserving stale
  source checks and exact no-op bytes.
- The action `LimitKind` now names every `EditOptions` budget dimension plus
  retained-source admission, and is wired into the actual budget-error paths;
  the typed dimensions remain visible through stable resource labels in
  DrawingML limit errors.

The coordinator should not infer contract closure merely from the 22 inventory
rows: this return specifically audits publication routes and owner entry
points in addition to those rows. The focused action tests cover exact no-op
replay, changed replay through `Prepared::apply`, stale-source refusal,
source-preserving inverse replay, and bounded output/readback paths.

## Glossary gate additions

`tools/api_glossary.json` bans the removed action-family names at their exact
canonical type scopes, removes the private `actions::edit` family registration,
registers `actions::builder` and `ink::builder`, nominates both InkML and
action `Prepared` owners, and registers action `Prepared::apply` as a commit
result method. No alias or deprecation forwarder remains.

## Validation

The follow-up source changes were validated with the coordinator-capped native
DrawingML test/check command recorded in the subsequent 0362dva validation
receipt. Cargo was not run before the source migration in this record; the
receipt records the exact post-change command, source/lock provenance, and
result separately.
