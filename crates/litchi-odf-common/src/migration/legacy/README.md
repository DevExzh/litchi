# Cross-family legacy migration

The files in this directory are source-preserving staging relocations from
the former `litchi-odf` root:

- `flat.rs` is a cross-family flat-document adapter. Its generic wrappers
  still reference the old umbrella and must be decomposed into the owning
  family crates before any public API is wired.
- `drawing/layers.rs` contains the family-neutral drawing-layer XML parser,
  plus old package/flat attachment methods that must be separated before
  publication from common.

They are deliberately not declared by `litchi-odf-common`: common must not
depend on concrete family crates merely to make this migration compile. The
deferred flat tests are colocated under `tests/deferred/flat` and remain out
of Cargo's direct integration-test discovery.
