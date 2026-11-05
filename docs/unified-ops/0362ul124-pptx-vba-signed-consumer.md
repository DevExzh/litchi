# 0362ul124 — PPTX VBA signed-error consumer correction

This source-only repair starts from the frozen OOXML runtime candidate
`a8d32513b8e427c6d234ddd17fb836b0e8d2477b` (tree
`0251ddc8aeef84702604f340406eda671c9b6c54`).  The coordinator's remaining
PPTX runtime receipt reached
`every_family_refuses_a_signed_package_until_it_is_unsigned` with the
shared-owner VBA row as the only failure.  No Cargo, compiler, rustdoc,
native, metadata, or performance command was run in this leaf.

## Runtime finding

The diagnostic is retained outside the repository at
`(deleted-coord-artifact)/0006.log`
(223,115 bytes, SHA-256
`8bd5fef94baf46ff1d07c15e71c36010d6bbe8343d277c1539cf96cc2638837e`).
The changed VBA operation returned PPTX's format-owned `Error::Signed`:

```
hello-world-signed.pptx: VBA project (shared owner): Signed
```

That result is intentional.  The shared producer returns
`litchi_ooxml_common::Error::Signed`, and the PPTX VBA adapter's `map_common`
preserves it as `litchi_pptx::Error::Signed`.  The existing table helper was
stale: it recognized only the lower-level
`Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)` used by the other
package families.

## Repair

The all-families assertion now uses a dedicated `Error::Signed` classifier
only for `VBA project (shared owner)`; every other family remains checked by
the strict OPC refusal classifier.  The no-op byte-equality assertion,
signed-package preservation checks, explicit `unsign`, unsigned changed
publication, and reopened readback remain in the same loop, so the fix does
not turn the test into a first-error-only check or weaken the capability
coverage.

The adjacent public VBA edit and raw-owner documentation now states the
actual crate-owned signed error rather than the lower-level OPC variant.
Production mapping and signature behavior are unchanged.

## 124a compile follow-up

The first source-only tip passed the classification review but passed the
owned `Family` into the helper after the loop had moved its `Change`.  The
helper now takes the family's copied static name instead.  This keeps the
dispatch typed and strict without borrowing a partially moved test value.

## Source-only validation

`git diff --check` passed.  The changed Rust paths were inspected for the
typed producer/consumer mapping and the reduced record contains no generated
evidence, fixture, executable, or `3rdparty/` path.  Cargo, rustdoc, and the
coordinator's runtime rerun remain intentionally pending.
