# 0361i — Retain Pages footnote resource policies

- Date: 2026-10-21.
- Status: source slice validated; integration with the final W4/W4b workspace
  remains required.
- Contract: 0351 §2.5 and 0355a §1 require owner-derived limits at every raw
  parse, rewrite and output boundary.

The Pages footnote graph projection and body-footnote edit paths now derive
wire, text-rewrite and Snappy profiles from the admitted package limits. The
projection keeps the package archive's input, field, nesting and rewrite-work
ceilings while constructing codec options; it no longer silently falls back to
its fixed codec defaults. Body-footnote insertion, removal, metadata rewrites,
attachment-owner validation, storage-candidate discovery and text-component
rewrites use the retained package profile for every wire parse and Snappy
bound/compression call. The Pages text helper maps an unsupported text-wire
nesting profile to its typed limit error rather than treating it as malformed
content.

The migration preserves the existing content policy. Malformed or semantically
invalid wire remains `InvalidSource`; input/output bytes, fields, nesting,
rewrite work, allocation and archive/Snappy ceilings retain typed limit or
allocation errors. Unknown storage fields are still copied through the rewrite
path. A package whose archive nesting limit is below the text-wire schema's
minimum is refused as `WireNesting` before attempting to decode. Source-derived
limits are checked before the corresponding parse, output allocation or
compression operation; there is no default-policy retry after a limit error.

Focused tests cover exact and one-over input/output byte ceilings, the text-wire
minimum nesting requirement, unknown wire-field preservation and typed failure
from a too-small field budget. Existing Pages footnote unit tests and IWA
footnote graph tests continue to exercise malformed, unknown, reference,
attachment and body-marker behavior.

Validation at tip `3600111b59` is recorded in
`results/0361i-footnote-validation-receipt.json`:

- `cargo test --offline --locked -p litchi-pages --lib --all-features` — 211
  passed, 0 failed.
- `cargo test --offline --locked -p litchi-iwa --lib --all-features` — 1,695
  passed, 0 failed.
- `cargo clippy --offline --locked -p litchi-pages -p litchi-iwa --all-features
  --lib -- -D warnings` — passed with no warnings.

The workspace-wide gate remains an integration responsibility after the
coordinator's error-contract and W4/W4b changes. No native timing is claimed
here.
