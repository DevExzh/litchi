# 0362ul87ab — OOXML properties staging repair

Status: source-only follow-up to 0362ul87aa.  No Cargo, rustdoc, native
qualification, performance measurement or final gate was run for this repair.

## Frozen review findings repaired

`properties::Plan::metadata` now reflects a staged core removal instead of
resurrecting the source core projection.  After `clear_metadata`, the getter
returns cleared core fields while retaining the application projection.
`update_metadata` keeps the core removal when an update changes only an
application field, and cancels the removal only when a core field is changed
explicitly.  Preparation consequently publishes a core-part removal together
with any application-only update, without clearing or rewriting untouched app
values.

Custom-property staging now distinguishes an absent custom part from a
present, valid empty custom part.  Staging an empty value removes the existing
empty graph, including its relationship, while an exact empty no-op remains a
no-op when the graph is absent.  The signed changed path still refuses before
publication.

## Regression coverage

The common transaction tests cover:

* clearing core properties, reading the staged app projection, updating an app
  field, and publishing the resulting core removal plus app change;
* removing a present empty `custom.xml`, and refusing the same graph change on
  a signed package.

## Source-only checks

`git diff --check` and scoped `rustfmt --edition 2024 --check` pass for the
changed Rust sources.  Runtime validation remains with the coordinator after
the W4/W4b integration; it should include the common transaction tests and
signed/no-op/replay checks.
