# OPC support review

The form-control batch adds a bounded exact-byte content-types payload handle
and exposes snapshot metadata for signature infrastructure. It uses the existing
relationship source APIs; the pending synthetic relationship-part map in
`package.rs` is excluded.

Focused tests cover retained content-types bytes, caller limits before budget
charge, managed memory/object/work reservation and release, cancellation before
read, source-version changes during read with cleanup, and metadata-only
signature detection without budget charge.

The OPC reviewer passed these preliminary checks in the shared workspace:

- `cargo test -p litchi-opc source_content_types_handle --lib`
- `cargo test -p litchi-opc source_signature --lib`
- `cargo test -p litchi-opc source_relationship_handle --lib`
- `cargo check -p litchi-opc --tests`

The final isolated full-crate receipts in `gates/` supersede these preliminary
checks for the committed batch. No synthetic `.rels` lookup behavior is added
by this work.
