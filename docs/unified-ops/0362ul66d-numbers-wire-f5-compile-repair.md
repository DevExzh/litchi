# 0362ul66d — Numbers-wire F5 compile repair

Status: source READY for the next scoped F5 retry.

The first compile-backed row after the approved-lock repair exposed two
compile errors that source-only checks could not see:

- the derive and field attributes still named `thiserror` after its direct
  dependency was removed; and
- `table_write::{fields, single_bytes}` returned slices borrowed from `data`
  without naming that lifetime.

The derive/attribute remnants are removed because `limits::Error` now has the
same fixed messages through its local `Display` implementation and the
standard `Error` trait.  The two table helpers now name the source lifetime;
their owner-selected `Limits` borrow remains independent and no allocation or
wire behavior changes.  Rustfmt was run on the two changed files and the
source diff is clean.

The failed row and its 245,581,146-byte target receipt remain in the external
F5 state directory.  The runner must restart from row 0 against the new clean
source commit; no later row was started.
