# 0362uk — Repair v9 OPC consumers

This source-only repair starts at integration commit `b170204d17df40c5240637437747d647026c895e`.
The v9 API-assembly check stopped in `litchi-opc` with two library errors and
149 test errors. The captured diagnostic log is retained outside the checkout
at `(deleted-coord-artifact)/check-0.log`
(68,469 bytes, SHA-256
`cb259594b244f0f1544fe7bd38c131024340c85badb0185d07ac4c43f882e21d`).

The library fixes import the existing `Arc` type for content-type source
materialization and add a narrow `OwnedRelationships::preflight_under` parent
module seam. The package replacement path now asks the relationship owner to
re-admit its private XML token under the recipient limits; the XML field and
payload boundary remain private.

All reported test consumers now use the required source constructor policy:
`SourceLimits`, `SourceCacheLimits` and `SourceReadPolicy` are imported in
their owning test modules, `PACKAGE_URI` is imported where the relationship
fixture needs it, and the denied redundant `crate::SourceCacheLimits`
qualifications are removed. Existing limits, typed errors, and assertions are
unchanged.

## Validation

- `rustfmt +1.95.0 --edition 2024 --config skip_children=true --config match_block_trailing_comma=true --check` passed on all 13 changed Rust files.
- `git diff --check` passed.
- Cargo was not run by design; this was an isolated source-only repair. The
  coordinator will run the focused OPC all-target check and the workspace gate
  after integration.
